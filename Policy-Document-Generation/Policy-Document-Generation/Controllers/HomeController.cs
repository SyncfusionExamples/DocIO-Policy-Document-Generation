using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting.Internal;
using Policy_Document_Generation.Models;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.XlsIO;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Reflection;

namespace Policy_Document_Generation.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _hostingEnvironment;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment hostingEnvironment)
        {
            _logger = logger;
            _hostingEnvironment = hostingEnvironment;
        }

        /// <summary>
        /// Main entry point for document generation. Supports multiple scenarios:
        /// 1. Single document with all policies (single file or ZIP with additional doc)
        /// 2. Separate documents per policy (ZIP archive)
        /// 3. Selective policies (specific policy numbers)
        /// 4. All policies (from Excel file)
        /// 
        /// OPTIMIZATION: Excel is read only ONCE, DataSet is created once and reused.
        /// For separate documents, uses merge-once-then-split strategy for better performance.
        /// </summary>
        public IActionResult GenerateDocument(ReportDataViewModel model)
        {
            try
            {
                // Step 1: Get template and Excel streams
                Stream documentStream = GetWordDocument(model.TemplateFile);
                Stream excelStream = GetExcel(model.TemplateFile, model.ExcelDataFile);

                // Step 2: Parse policy numbers based on selection mode
                List<string> policyNumbersList = ParsePolicyNumbers(model);

                // Step 3: If "all policies" mode is selected and separate files is requested,
                // fetch all policy numbers from Excel file
                if (model.SelectionMode == "all" && model.GenerateSeparateFiles)
                {
                    policyNumbersList = ExtractAllPolicyNumbersFromExcel(excelStream);
                    excelStream.Position = 0; // Reset stream position after reading
                }

                // Step 4: Check if separate files should be generated for each policy
                if (model.GenerateSeparateFiles && policyNumbersList.Count > 1)
                {
                    // Generate separate documents for each policy and return as ZIP
                    return GenerateSeparateDocuments(documentStream, excelStream, policyNumbersList, model);
                }
                else
                {
                    // Generate single document with all selected policies
                    return GenerateSingleDocument(documentStream, excelStream, policyNumbersList, model);
                }
            }
            catch (Exception ex)
            {
                TempData["Error"] = $"Error generating document: {ex.Message}";
                return RedirectToAction("Index");
            }
        }

        /// <summary>
        /// Parses policy numbers from the model based on selection mode
        /// </summary>
        private List<string> ParsePolicyNumbers(ReportDataViewModel model)
        {
            List<string> policyNumbers = new List<string>();

            if (model.SelectionMode == "all")
            {
                // Return null/empty to indicate all policies should be processed
                return policyNumbers; // Empty list means "all"
            }
            else if (model.SelectionMode == "specific")
            {
                if (!string.IsNullOrWhiteSpace(model.PolicyNumbers))
                {
                    // Split by newlines and commas, trim whitespace, remove empty entries
                    policyNumbers = model.PolicyNumbers
                        .Split(new[] { '\n', '\r', ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(p => p.Trim())
                        .Where(p => !string.IsNullOrWhiteSpace(p))
                        .Distinct() // Remove duplicates
                        .ToList();
                }
            }

            return policyNumbers;
        }

        /// <summary>
        /// Extracts all unique policy numbers from the first column of the Excel file
        /// Used when "all policies" mode is selected with separate file generation
        /// </summary>
        private List<string> ExtractAllPolicyNumbersFromExcel(Stream excelStream)
        {
            List<string> policyNumbers = new List<string>();

            try
            {
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    application.DefaultVersion = ExcelVersion.Xlsx;

                    IWorkbook workbook = application.Workbooks.Open(excelStream);
                    
                    // Get the first sheet (main data sheet)
                    if (workbook.Worksheets.Count > 0)
                    {
                        IWorksheet sheet = workbook.Worksheets[0];
                        
                        if (sheet?.UsedRange != null && sheet.UsedRange.LastRow > 1)
                        {
                            int headerRow = sheet.UsedRange.Row;
                            int lastRow = sheet.UsedRange.LastRow;

                            // Extract all values from the first column (skip header)
                            for (int row = headerRow + 1; row <= lastRow; row++)
                            {
                                string value = sheet[row, 1].Value?.ToString()?.Trim();
                                
                                // Add unique non-empty values
                                if (!string.IsNullOrWhiteSpace(value) && !policyNumbers.Contains(value))
                                {
                                    policyNumbers.Add(value);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error extracting policy numbers from Excel: {ex.Message}", ex);
            }

            return policyNumbers;
        }

        /// <summary>
        /// Generates a single document containing all selected policies.
        /// If an additional document type is selected, both are returned as a ZIP archive.
        /// Supports both DOCX and PDF output formats.
        /// </summary>
        private IActionResult GenerateSingleDocument(Stream documentStream, Stream excelStream, List<string> policyNumbers, ReportDataViewModel model)
        {
            try
            {
                // Create DataSet from Excel with relations
                DataSet dataSet = CreateMailMergeDataSet(excelStream, policyNumbers);

                if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                {
                    TempData["Error"] = "No data found for the selected policy numbers.";
                    return RedirectToAction("Index");
                }

                bool hasAdditionalDoc = !string.IsNullOrEmpty(model.DocumentType);
                bool isPdf = model.OutputFormat?.ToLower() == "pdf";
                string fileExtension = isPdf ? "pdf" : "docx";
                string baseName = policyNumbers.Count == 1 ? $"Policy_{policyNumbers[0]}" : $"Policies_{policyNumbers.Count}_Documents";

                // Generate main policy document bytes
                byte[] mainDocBytes = GenerateDocumentBytes(documentStream, dataSet, model, isPdf);

                if (!hasAdditionalDoc)
                {
                    // Return single file directly
                    string mimeType = isPdf ? "application/pdf"
                        : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                    return File(mainDocBytes, mimeType, $"{baseName}.{fileExtension}");
                }

                // Generate additional document — template & data resolved via priority chain
                Stream additionalTemplateStream = GetAdditionalDocumentTemplate(model.DocumentType, model.AdditionalDocTemplateFile);
                byte[] additionalDocBytes = null;

                if (additionalTemplateStream != null)
                {
                    DataSet additionalDataSet = GetAdditionalDocDataSet(model, excelStream, dataSet, policyNumbers);
                    additionalDocBytes = GenerateDocumentBytes(additionalTemplateStream, additionalDataSet, model, isPdf);
                    additionalTemplateStream.Dispose();
                }

                // Return both documents as a ZIP archive
                string additionalDocLabel = GetAdditionalDocumentLabel(model.DocumentType);
                using (MemoryStream zipStream = new MemoryStream())
                {
                    using (var archive = new System.IO.Compression.ZipArchive(zipStream, System.IO.Compression.ZipArchiveMode.Create, true))
                    {
                        // Add main policy document
                        var mainEntry = archive.CreateEntry($"{baseName}.{fileExtension}");
                        using (var entryStream = mainEntry.Open())
                            entryStream.Write(mainDocBytes, 0, mainDocBytes.Length);

                        // Add additional document
                        if (additionalDocBytes != null)
                        {
                            var addEntry = archive.CreateEntry($"{baseName}_{additionalDocLabel}.{fileExtension}");
                            using (var entryStream = addEntry.Open())
                                entryStream.Write(additionalDocBytes, 0, additionalDocBytes.Length);
                        }
                    }

                    zipStream.Position = 0;
                    return File(zipStream.ToArray(), "application/zip", $"{baseName}_Documents_{DateTime.Now:yyyyMMdd_HHmmss}.zip");
                }
            }
            catch (Exception ex)
            {
                TempData["Error"] = $"Error generating document: {ex.Message}";
                return RedirectToAction("Index");
            }
        }

        /// <summary>
        /// Performs mail merge on a Word template stream and returns the result as a byte array.
        /// Applies MultiPartDocument cover page and PDF conversion based on model settings.
        /// </summary>
        private byte[] GenerateDocumentBytes(Stream templateStream, DataSet dataSet, ReportDataViewModel model, bool convertToPdf)
        {
            using (WordDocument document = new WordDocument(templateStream, FormatType.Docx))
            {
                ExecuteMailMerge(document, dataSet);

                if (model.MultiPartDocument)
                {
                    AddCoverPageAndTOC(document);
                }

                using (MemoryStream outputStream = new MemoryStream())
                {
                    if (convertToPdf)
                    {
                        Syncfusion.DocIORenderer.DocIORenderer renderer = new Syncfusion.DocIORenderer.DocIORenderer();
                        Syncfusion.Pdf.PdfDocument pdfDocument = renderer.ConvertToPDF(document);
                        pdfDocument.Save(outputStream);
                        pdfDocument.Close();
                        renderer.Dispose();
                    }
                    else
                    {
                        document.Save(outputStream, FormatType.Docx);
                    }

                    return outputStream.ToArray();
                }
            }
        }

        /// <summary>
        /// Returns the file stream for the selected additional document template from wwwroot/Data.
        /// Maps document type keys to their corresponding template filenames.
        /// </summary>
        /// <summary>
        /// Resolves the Word template stream for the additional document.
        /// Priority: 1) User-uploaded custom template  2) Built-in default template from wwwroot/Data.
        /// </summary>
        private Stream GetAdditionalDocumentTemplate(string documentType, IFormFile uploadedTemplate = null)
        {
            // Priority 1: user-uploaded custom template
            if (uploadedTemplate != null && uploadedTemplate.Length > 0)
            {
                MemoryStream ms = new MemoryStream();
                uploadedTemplate.CopyTo(ms);
                ms.Position = 0;
                return ms;
            }

            // Priority 2: built-in default template
            if (string.IsNullOrEmpty(documentType))
                return null;

            string fileName = documentType switch
            {
                "claims-letter"  => "Claim Letter.docx",
                "endorsement"    => "EndorsementTemplate.docx",
                "certificate"    => "Certificate Of Insurance.docx",
                "renewal-notice" => "RenewalNoticeTemplate.docx",
                _                => null
            };

            if (fileName == null)
                return null;

            string filePath = Path.Combine(_hostingEnvironment.WebRootPath, "Data", fileName);

            if (!System.IO.File.Exists(filePath))
                return null;

            return new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        }

        /// <summary>
        /// Returns a clean display label for an additional document type key (used in file naming).
        /// </summary>
        private string GetAdditionalDocumentLabel(string documentType)
        {
            return documentType switch
            {
                "claims-letter"  => "ClaimsLetter",
                "endorsement"    => "Endorsement",
                "certificate"    => "Certificate",
                "renewal-notice" => "RenewalNotice",
                _                => "AdditionalDocument"
            };
        }

        /// <summary>
        /// Resolves the DataSet used for additional document mail merge.
        ///
        /// Resolution strategy:
        ///   1. If user uploaded a separate Excel for the additional doc → use it exclusively.
        ///   2. Otherwise, reuse the main Excel stream:
        ///      a. If the main DataSet already contains a sheet that matches the additional
        ///         document's expected sheet name (e.g. "Claims") → pass the full DataSet as-is
        ///         so the additional template can access that sheet's merge group.
        ///      b. If no matching sheet exists → fall back to the same main DataSet
        ///         (fields that exist will be merged; unrecognised fields are left blank).
        ///
        /// This means a single Excel with multiple sheets (Policies, Claims, Endorsements …)
        /// covers all document types without any extra file.
        /// </summary>
        private DataSet GetAdditionalDocDataSet(
            ReportDataViewModel model,
            Stream mainExcelStream,
            DataSet mainDataSet,
            List<string> policyNumbers)
        {
            // --- Priority 1: dedicated separate Excel uploaded by user ---
            if (model.AdditionalDocExcelFile != null && model.AdditionalDocExcelFile.Length > 0)
            {
                MemoryStream separateStream = new MemoryStream();
                model.AdditionalDocExcelFile.CopyTo(separateStream);
                separateStream.Position = 0;
                return CreateMailMergeDataSet(separateStream, policyNumbers);
            }

            // --- Priority 2: reuse main Excel (same stream, reset position) ---
            // The CreateMailMergeDataSet already loaded all sheets into mainDataSet.
            // We can pass it directly — ExecuteMailMerge will only use the tables
            // whose names match the «TableStart:Name» groups in the template.
            // Reset the main stream so it can be reopened if needed elsewhere.
            if (mainExcelStream.CanSeek)
                mainExcelStream.Position = 0;

            return mainDataSet;
        }

        /// <summary>
        /// OPTIMIZED: Generates separate documents using merge-once-then-split strategy.
        /// This approach is simpler, faster, and easier to understand:
        /// 1. Reads Excel once and creates DataSet
        /// 2. Performs mail merge once with ALL data (each record starts on new page)
        /// 3. Splits the merged document by page breaks
        /// 4. Returns ZIP archive with separate files
        /// </summary>
        private IActionResult GenerateSeparateDocuments(Stream documentStream, Stream excelStream, List<string> policyNumbers, ReportDataViewModel model)
        {
            try
            {
                bool isPdf = model.OutputFormat?.ToLower() == "pdf";
                string fileExtension = isPdf ? "pdf" : "docx";
                bool hasAdditionalDoc = !string.IsNullOrEmpty(model.DocumentType);

                // Step 1: Create DataSet once - reads Excel only once
                DataSet fullDataSet = CreateMailMergeDataSet(excelStream, policyNumbers);

                if (fullDataSet.Tables.Count == 0)
                {
                    TempData["Error"] = "No data found for the selected policy numbers.";
                    return RedirectToAction("Index");
                }

                using (MemoryStream zipStream = new MemoryStream())
                {
                    using (var archive = new System.IO.Compression.ZipArchive(zipStream, System.IO.Compression.ZipArchiveMode.Create, true))
                    {
                        // Step 2: Generate main policy documents (merge once, split by page breaks)
                        byte[] mainDocsZip = GenerateDocumentsWithPageBreakSplit(
                            documentStream, fullDataSet, policyNumbers, model, isPdf, fileExtension);

                        // Step 3: Add main documents to final ZIP
                        AddZipContentsToArchive(mainDocsZip, archive);

                        // Step 4: Generate additional documents if requested
                        if (hasAdditionalDoc)
                        {
                            Stream additionalTemplateStream = GetAdditionalDocumentTemplate(
                                model.DocumentType, model.AdditionalDocTemplateFile);

                            if (additionalTemplateStream != null)
                            {
                                // Get DataSet for additional document (reuses main DataSet or loads separate Excel)
                                excelStream.Position = 0;
                                DataSet additionalDataSet = GetAdditionalDocDataSet(model, excelStream, fullDataSet, policyNumbers);

                                string additionalLabel = GetAdditionalDocumentLabel(model.DocumentType);

                                byte[] additionalDocsZip = GenerateDocumentsWithPageBreakSplit(
                                    additionalTemplateStream, additionalDataSet, policyNumbers, 
                                    model, isPdf, fileExtension, additionalLabel);

                                AddZipContentsToArchive(additionalDocsZip, archive);
                            }
                        }
                    }

                    zipStream.Position = 0;
                    return File(zipStream.ToArray(), "application/zip", 
                        $"Policy_Documents_{DateTime.Now:yyyyMMdd_HHmmss}.zip");
                }
            }
            catch (Exception ex)
            {
                TempData["Error"] = $"Error generating separate documents: {ex.Message}";
                return RedirectToAction("Index");
            }
        }

        /// <summary>
        /// CORE OPTIMIZATION: Merge once with all data, then split by page breaks.
        /// This is much simpler and faster than filtering and merging multiple times.
        /// Returns ZIP bytes containing separate documents for each policy.
        /// </summary>
        private byte[] GenerateDocumentsWithPageBreakSplit(
            Stream templateStream,
            DataSet dataSet,
            List<string> policyNumbers,
            ReportDataViewModel model,
            bool isPdf,
            string fileExtension,
            string fileNameSuffix = "")
        {
            using (WordDocument document = new WordDocument(templateStream, FormatType.Docx))
            {
                // KEY: Ensure each record starts on a new page (creates page breaks automatically)
                document.MailMerge.StartAtNewPage = true;

                // Execute mail merge ONCE with all data
                ExecuteMailMerge(document, dataSet);

                // Add cover page if requested
                if (model.MultiPartDocument)
                {
                    AddCoverPageAndTOC(document);
                }

                // Split the merged document by page breaks and return as ZIP
                return SplitDocumentByPageBreaks(document, policyNumbers, isPdf, fileExtension, fileNameSuffix);
            }
        }

        /// <summary>
        /// Splits a merged document into separate files using page breaks as boundaries.
        /// Uses bookmark navigation to extract content between page breaks.
        /// Returns ZIP archive bytes containing all separated documents.
        /// </summary>
        private byte[] SplitDocumentByPageBreaks(
            WordDocument document,
            List<string> policyNumbers,
            bool convertToPdf,
            string fileExtension,
            string fileNameSuffix)
        {
            // Step 1: Find all page breaks in the document
            List<Entity> pageBreaks = document.FindAllItemsByProperty(
                EntityType.Break, "BreakType", "PageBreak");

            WSection section = document.Sections[0];
            WTextBody body = section.Body;
            int bookmarkIndex = 1;

            // Step 2: Insert bookmark at the document start
            WParagraph firstBookmarkPara = new WParagraph(document);
            firstBookmarkPara.AppendBookmarkStart($"Policy_Section_{bookmarkIndex}");
            body.ChildEntities.Insert(0, firstBookmarkPara);

            // Step 3: Insert bookmarks after each page break
            foreach (Entity entity in pageBreaks)
            {
                WParagraph breakParagraph = entity.Owner as WParagraph;
                if (breakParagraph == null) continue;

                int paraIndex = body.ChildEntities.IndexOf(breakParagraph);
                if (paraIndex < 0) continue;

                // Close current bookmark and start next one
                WParagraph bookmarkPara = new WParagraph(document);
                bookmarkPara.AppendBookmarkEnd($"Policy_Section_{bookmarkIndex}");
                bookmarkIndex++;
                bookmarkPara.AppendBookmarkStart($"Policy_Section_{bookmarkIndex}");

                body.ChildEntities.Insert(paraIndex + 1, bookmarkPara);
            }

            // Step 4: Insert bookmark at the document end
            WParagraph lastBookmarkPara = new WParagraph(document);
            lastBookmarkPara.AppendBookmarkEnd($"Policy_Section_{bookmarkIndex}");
            body.ChildEntities.Add(lastBookmarkPara);

            // Step 5: Extract each section and create ZIP archive
            using (MemoryStream zipStream = new MemoryStream())
            {
                using (var archive = new System.IO.Compression.ZipArchive(
                    zipStream, System.IO.Compression.ZipArchiveMode.Create, true))
                {
                    for (int i = 1; i <= bookmarkIndex && i <= policyNumbers.Count; i++)
                    {
                        string policyNumber = policyNumbers[i - 1];
                        
                        // Build filename: "Policy_12345.pdf" or "Policy_12345_ClaimsLetter.pdf"
                        string fileName = string.IsNullOrEmpty(fileNameSuffix)
                            ? $"Policy_{policyNumber}.{fileExtension}"
                            : $"Policy_{policyNumber}_{fileNameSuffix}.{fileExtension}";

                        try
                        {
                            // Navigate to bookmark and extract content
                            BookmarksNavigator navigator = new BookmarksNavigator(document);
                            navigator.MoveToBookmark($"Policy_Section_{i}", true, true);
                            WordDocumentPart documentPart = navigator.GetContent();

                            if (documentPart == null) continue;

                            // Save extracted section as separate file
                            using (WordDocument extractedDoc = documentPart.GetAsWordDocument())
                            using (MemoryStream docStream = new MemoryStream())
                            {
                                SaveDocumentToStream(extractedDoc, docStream, convertToPdf);
                                docStream.Position = 0;

                                var zipEntry = archive.CreateEntry(fileName);
                                using (var entryStream = zipEntry.Open())
                                {
                                    docStream.CopyTo(entryStream);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, $"Error processing policy {policyNumber}");
                        }
                    }
                }

                return zipStream.ToArray();
            }
        }

        /// <summary>
        /// Helper method: Extracts files from a ZIP byte array and adds them to target archive.
        /// Used to combine main documents and additional documents into a single ZIP.
        /// </summary>
        private void AddZipContentsToArchive(byte[] zipBytes, System.IO.Compression.ZipArchive targetArchive)
        {
            using (MemoryStream zipStream = new MemoryStream(zipBytes))
            using (var sourceArchive = new System.IO.Compression.ZipArchive(
                zipStream, System.IO.Compression.ZipArchiveMode.Read))
            {
                foreach (var entry in sourceArchive.Entries)
                {
                    var newEntry = targetArchive.CreateEntry(entry.FullName);
                    using (var sourceStream = entry.Open())
                    using (var targetStream = newEntry.Open())
                    {
                        sourceStream.CopyTo(targetStream);
                    }
                }
            }
        }

        /// <summary>
        /// Saves a WordDocument to a MemoryStream as either PDF or DOCX.
        /// </summary>
        private void SaveDocumentToStream(WordDocument document, MemoryStream stream, bool convertToPdf)
        {
            if (convertToPdf)
            {
                Syncfusion.DocIORenderer.DocIORenderer renderer = new Syncfusion.DocIORenderer.DocIORenderer();
                Syncfusion.Pdf.PdfDocument pdfDocument = renderer.ConvertToPDF(document);
                pdfDocument.Save(stream);
                pdfDocument.Close();
                renderer.Dispose();
            }
            else
            {
                document.Save(stream, FormatType.Docx);
            }
        }

        /// <summary>
        /// Adds cover page and table of contents to the document
        /// </summary>
        private void AddCoverPageAndTOC(WordDocument document)
        {
            WSection firstSection = document.Sections[0].Clone();
            firstSection.Body.ChildEntities.Clear();
            // Insert a new section at the beginning for cover page
            document.Sections.Insert(0, firstSection);
            IWParagraph coverPara = firstSection.AddParagraph();
            coverPara.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
            coverPara.AppendText("POLICY DOCUMENT").CharacterFormat.FontSize = 24;
            coverPara.AppendBreak(BreakType.LineBreak);
            coverPara.AppendBreak(BreakType.LineBreak);
            coverPara.AppendText($"Generated on: {DateTime.Now:MMMM dd, yyyy}").CharacterFormat.FontSize = 12;

            // Add page break after cover
            coverPara.AppendBreak(BreakType.PageBreak);
        }

        /// <summary>
        /// Reads Excel file and creates DataSet for mail merge
        /// Updated to handle multiple policy numbers
        /// Note: DataRelations are NOT needed for DocIO's ExecuteNestedGroup - only DataSet and commands ArrayList
        /// </summary>
        public DataSet CreateMailMergeDataSet(Stream excelStream, List<string> policyNumbers = null)
        {
            DataSet dataSet = new DataSet();

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook workbook = application.Workbooks.Open(excelStream);

                // Step 1: Identify table structure and relationships
                var tableStructure = AnalyzeExcelStructure(workbook);

                // Step 2: Create DataTables from sheets
                foreach (var tableInfo in tableStructure)
                {
                    DataTable dt = ReadExcelSheetToDataTable(workbook, tableInfo.SheetName, policyNumbers);
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        dataSet.Tables.Add(dt);
                    }
                }

                // Store table structure in DataSet extended properties for later use
                dataSet.ExtendedProperties["TableStructure"] = tableStructure;
            }

            return dataSet;
        }

        /// <summary>
        /// Reads a single Excel sheet into a DataTable with optional filtering by multiple policy numbers
        /// </summary>
        private DataTable ReadExcelSheetToDataTable(IWorkbook workbook, string sheetName, List<string> policyNumbers)
        {
            IWorksheet sheet = workbook.Worksheets[sheetName];
            if (sheet?.UsedRange == null)
                return null;

            IRange usedRange = sheet.UsedRange;
            int headerRow = usedRange.Row;
            int lastRow = usedRange.LastRow;
            int lastCol = usedRange.LastColumn;

            // Create DataTable
            DataTable dt = new DataTable(sheetName);

            // Add columns from header row
            for (int col = 1; col <= lastCol; col++)
            {
                string columnName = sheet[headerRow, col].Value ?? $"Column{col}";
                dt.Columns.Add(columnName);
            }

            // Determine filtering logic
            bool shouldFilter = policyNumbers != null && policyNumbers.Count > 0;

            // Add data rows (skip header)
            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                string firstColumnValue = sheet[row, 1].Value?.ToString()?.Trim();

                // Apply filter if policy numbers specified
                if (shouldFilter)
                {
                    // Check if first column value matches any of the policy numbers
                    if (string.IsNullOrWhiteSpace(firstColumnValue) ||
                        !policyNumbers.Any(p => p.Equals(firstColumnValue, StringComparison.OrdinalIgnoreCase)))
                    {
                        continue; // Skip this row
                    }
                }

                // Add row to DataTable
                DataRow dr = dt.NewRow();
                for (int col = 1; col <= lastCol; col++)
                {
                    dr[col - 1] = sheet[row, col].Value ?? string.Empty;
                }
                dt.Rows.Add(dr);
            }

            return dt;
        }
        /// <summary>
        /// Reads Excel sheets and captures only schema information
        /// (NO relationship inference here)
        /// </summary>
        private List<TableStructureInfo> AnalyzeExcelStructure(IWorkbook workbook)
        {
            var tables = new List<TableStructureInfo>();

            foreach (IWorksheet sheet in workbook.Worksheets)
            {
                if (sheet.UsedRange == null || sheet.UsedRange.LastRow < 2)
                    continue;

                var tableInfo = new TableStructureInfo
                {
                    SheetName = sheet.Name,
                    Columns = new List<string>()
                };

                int lastCol = sheet.UsedRange.LastColumn;
                for (int col = 1; col <= lastCol; col++)
                {
                    string columnName = sheet[1, col].Value?.Trim();
                    if (!string.IsNullOrEmpty(columnName))
                    {
                        tableInfo.Columns.Add(columnName);
                    }
                }

                // Store for reference only (not for hierarchy decisions)
                tableInfo.PrimaryKeyColumn = tableInfo.Columns.FirstOrDefault();
                tables.Add(tableInfo);
            }

            return tables;
        }

        private HashSet<string> GetTemplateMergeGroups(WordDocument document)
        {
            return new HashSet<string>(
                document.MailMerge.GetMergeGroupNames(),
                StringComparer.OrdinalIgnoreCase);
        }

        private bool IsUsedAsNestedGroup(WordDocument document, string tableName)
        {
            var groupNames = document.MailMerge.GetMergeGroupNames();
            return groupNames.Any(g =>
                string.Equals(g, tableName, StringComparison.OrdinalIgnoreCase));
        }


        /// <summary>
        /// Executes mail merge based on data structure (flat, nested, or multiple tables)
        /// Maintains parent-child relationships for nested groups
        /// </summary>
        public void ExecuteMailMerge(WordDocument document, DataSet dataSet)
        {
            if (dataSet.Tables.Count == 0)
                return;

            try
            {
                // Check if nested structure exists based on table structure analysis
                var tableStructure = dataSet.ExtendedProperties["TableStructure"] as List<TableStructureInfo>;
                bool hasNestedStructure = HasNestedStructure(
                    document,
                    dataSet,
                    tableStructure
                );

                document.MailMerge.StartAtNewPage = true;

                if (dataSet.Tables.Count > 1 && hasNestedStructure)
                {
                    // Execute nested mail merge - maintains parent-child relationships
                    ExecuteNestedMailMerge(document, dataSet, tableStructure);
                }
                else if (dataSet.Tables.Count == 1)
                {
                    // Single table - use ExecuteGroup for repeating data, Execute for single record
                    DataTable table = dataSet.Tables[0];
                    
                    if (table.Rows.Count > 1)
                    {
                        // Multiple records - use ExecuteGroup to repeat the region
                        document.MailMerge.ExecuteGroup(table);
                    }
                    else
                    {
                        document.MailMerge.StartAtNewPage = false;
                        // Single record - use Execute for simple field replacement
                        document.MailMerge.Execute(table);
                    }
                }
                else if (dataSet.Tables.Count > 1)
                {
                    // Multiple tables without nested structure - execute nested group with all tables
                    ExecuteMultipleTablesAsNestedGroup(document, dataSet);
                }
                document.UpdateDocumentFields();
            }
            catch (Exception ex)
            {
                throw new Exception($"Mail merge failed: {ex.Message}", ex);
            }
        }

        private bool HasNestedStructure(
            WordDocument document,
            DataSet dataSet,
            List<TableStructureInfo> tableStructure)
        {
            var groupNames = document.MailMerge.GetMergeGroupNames();

            if (groupNames == null || groupNames.Length < 2)
                return false;

            // Root group (first appearance)
            string rootGroup = groupNames[0];

            var parentTable = dataSet.Tables
                .Cast<DataTable>()
                .FirstOrDefault(t =>
                    string.Equals(t.TableName, rootGroup, StringComparison.OrdinalIgnoreCase));

            if (parentTable == null)
                return false;

            // Check if ANY other group can form a valid relation
            foreach (string group in groupNames.Skip(1))
            {
                var childTable = dataSet.Tables
                    .Cast<DataTable>()
                    .FirstOrDefault(t =>
                        string.Equals(t.TableName, group, StringComparison.OrdinalIgnoreCase));

                if (childTable == null)
                    continue;

                // Try to build a valid relation
                var relation = TryBuildRelation(parentTable, childTable, tableStructure);

                if (!string.IsNullOrEmpty(relation))
                    return true;   // ✅ TRUE nested structure
            }

            return false;
        }

        private string TryBuildRelation(
            DataTable parent,
            DataTable child,
            List<TableStructureInfo> tableStructure)
        {
            var parentColumns = parent.Columns.Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var childColumns = child.Columns.Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            // Find common key (example: PolicyNumber)
            var commonKey = parentColumns.Intersect(childColumns).FirstOrDefault();

            if (commonKey == null)
                return null;   // ❌ no relation

            return $"{commonKey} = %{child.TableName}.{commonKey}%";
        }

        /// <summary>
        /// Executes nested mail merge for hierarchical data structures with proper group handling
        /// Uses ArrayList of commands to define relationships - NO DataRelations needed
        /// </summary>
        private void ExecuteNestedMailMerge(WordDocument document, DataSet dataSet, List<TableStructureInfo> tableStructure)
        {
            try
            {
                ArrayList commands = new ArrayList();

                var mergeGroups = GetTemplateMergeGroups(document);

                if (mergeGroups.Count == 0)
                    throw new Exception("No BeginGroup found in Word template");

                // Parent (first group)
                string parentGroup = mergeGroups.First();
                var parentTable = tableStructure.First(t =>
                    string.Equals(t.SheetName, parentGroup, StringComparison.OrdinalIgnoreCase));

                commands.Add(new DictionaryEntry(parentTable.SheetName, string.Empty));

                // Children (remaining groups)
                foreach (string childGroup in mergeGroups.Skip(1))
                {
                    var childTable = tableStructure.FirstOrDefault(t =>
                        string.Equals(t.SheetName, childGroup, StringComparison.OrdinalIgnoreCase));

                    if (childTable == null)
                        continue;

                    string relationString = BuildRelationString(dataSet,parentTable.SheetName,childTable.SheetName);
                    commands.Add(new DictionaryEntry(childTable.SheetName, relationString));
                }

                document.MailMerge.ExecuteNestedGroup(dataSet, commands);
            }
            catch (Exception ex)
            {
                throw new Exception($"Nested mail merge failed: {ex.Message}", ex);
            }
        }
        /// <summary>
        /// Builds the relation string for nested mail merge command
        /// Format: "ChildKeyColumn = %ParentTableName.ParentKeyColumn%"
        /// The relation references the PARENT table, not the child table
        /// </summary>

        private string BuildRelationString(
            DataSet dataSet,
            string parentTableName,
            string childTableName)
        {
            DataTable parentTable = dataSet.Tables[parentTableName];
            DataTable childTable = dataSet.Tables[childTableName];

            if (parentTable == null || childTable == null)
                return string.Empty;

            // Find a common column between parent and child
            var parentColumns = parentTable.Columns
                .Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var childColumns = childTable.Columns
                .Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var commonKey = parentColumns
                .Intersect(childColumns)
                .FirstOrDefault();

            if (commonKey == null)
                return string.Empty;

            // ✅ Correct DocIO relation format
            return $"{commonKey} = %{childTableName}.{commonKey}%";
        }


        /// <summary>
        /// Executes mail merge for multiple independent tables using ExecuteNestedGroup
        /// Even without explicit parent-child relationships, ExecuteNestedGroup handles multiple tables
        /// </summary>
        private void ExecuteMultipleTablesAsNestedGroup(WordDocument document, DataSet dataSet)
        {
            try
            {
                ArrayList commands = new ArrayList();

                // Add all tables to commands with empty relation strings
                // This allows ExecuteNestedGroup to handle multiple independent merge regions
                foreach (DataTable table in dataSet.Tables)
                {
                    DictionaryEntry command = new DictionaryEntry(table.TableName, string.Empty);
                    commands.Add(command);
                }

                // Execute nested group - works for both related and independent tables
                document.MailMerge.ExecuteNestedGroup(dataSet, commands);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error executing nested group merge: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Retrieves a Word document stream from the uploaded file or a default template.
        /// </summary>

        private Stream GetWordDocument(IFormFile file)
        {
            // Case 1: Uploaded file exists and has content
            if (file != null && file.Length > 0)
            {
                string extension = Path.GetExtension(file.FileName).ToLower();
                string[] supportedExtensions = { ".doc", ".docx", ".dot", ".dotx", ".dotm", ".docm", ".xml", ".rtf" };
                // Validate the file extension
                if (supportedExtensions.Contains(extension))
                {
                    // Copy the uploaded file into an in-memory stream
                    MemoryStream stream = new MemoryStream();
                    file.CopyTo(stream);
                    // Reset stream position to the beginning for downstream reading
                    stream.Position = 0;
                    return stream;
                }
                else
                {
                    ViewBag.Message = "Please choose a Word format document to convert to PDF.";
                    return null;
                }
            }
            else
            {
                // Load default file from wwwroot\Data\
                string defaultFilePath = Path.Combine(_hostingEnvironment.WebRootPath, "Data", "Template.docx");
                return new FileStream(defaultFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            }
        }
        /// <summary>
        /// Retrieves an Json stream based on the uploaded files.
        /// </summary>
        private Stream GetExcel(IFormFile file, IFormFile jsonFile)
        {
            // If Word file was uploaded, JSON file must also be uploaded
            if (file != null && file.Length > 0)
            {
                // Ensure an Json file is also uploaded
                if (jsonFile != null && jsonFile.Length > 0)
                {
                    // Copy uploaded Json file into an in-memory stream.
                    MemoryStream stream = new MemoryStream();
                    jsonFile.CopyTo(stream);
                    // Reset stream position so it can be read from the beginning
                    stream.Position = 0;
                    return stream;
                }
                else
                {
                    ViewBag.Message = "Please upload a JSON data file along with the Word document.";
                    return null;
                }
            }
            else
            {
                // Both Word and JSON are defaults (no file uploaded)
                string defaultJsonPath = Path.Combine(_hostingEnvironment.WebRootPath, "Data", "Data.xlsx");
                return new FileStream(defaultJsonPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            }
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public class TableStructureInfo
        {
            public string SheetName { get; set; }
            public List<string> Columns { get; set; } = new List<string>();
            public List<string> KeyColumns { get; set; } = new List<string>();
            public string PrimaryKeyColumn { get; set; }
            public bool IsParent { get; set; }
            public string ParentTableName { get; set; }
            public string ParentKeyColumn { get; set; }
        }

    }

}
