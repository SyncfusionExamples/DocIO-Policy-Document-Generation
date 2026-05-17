using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting.Internal;
using Policy_Document_Generation.Models;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
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
        /// Excel is read only ONCE, DataSet is created once and reused.
        /// For separate documents, uses merge-once-then-split strategy for better performance.
        /// </summary>
        public IActionResult GenerateDocument(ReportDataViewModel model)
        {
            try
            {
                Stream documentStream = GetWordDocument(model.TemplateFile);
                Stream excelStream = GetExcel(model.TemplateFile, model.ExcelDataFile);

                // Step 2: Parse policy numbers from user input (specific mode)
                List<string> policyNumbersList = ParsePolicyNumbers(model);

                // Step 3: Determine filtering strategy
                bool needsFiltering = model.SelectionMode == "specific" && policyNumbersList.Count > 0;

                // Step 4: Create DataSet (filter during read if specific policies selected)
                DataSet dataSet = CreateMailMergeDataSet(excelStream, needsFiltering ? policyNumbersList : null);

                if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                {
                    TempData["Error"] = "No data found in Excel file.";
                    return RedirectToAction("Index");
                }

                // Step 5: For "all with separate files" - extract policy numbers from DataTable
                if (model.SelectionMode == "all" && model.GenerateSeparateFiles)
                {
                    // Extract from first column of first table (already in memory)
                    policyNumbersList = dataSet.Tables[0].AsEnumerable()
                        .Select(row => row[0]?.ToString()?.Trim())
                        .Where(p => !string.IsNullOrWhiteSpace(p))
                        .Distinct()
                        .ToList();
                }

                // Step 6: Generate separate or single document
                if (model.GenerateSeparateFiles && policyNumbersList.Count > 1)
                {
                    // Pass the already-created DataSet, not excelStream
                    return GenerateSeparateDocuments(documentStream, dataSet, excelStream, policyNumbersList, model);
                }
                else
                {
                    return GenerateSingleDocument(documentStream, dataSet, excelStream, policyNumbersList, model);
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
        /// Generates a single document containing all selected policies.
        /// If an additional document type is selected, both are returned as a ZIP archive.
        /// Supports both DOCX and PDF output formats.
        /// </summary>
        private IActionResult GenerateSingleDocument(Stream documentStream, DataSet dataSet, Stream excelStream, List<string> policyNumbers, ReportDataViewModel model)
        {
            try
            {
                if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                {
                    TempData["Error"] = "No data found for the selected policy numbers.";
                    return RedirectToAction("Index");
                }

                bool isPdf = model.OutputFormat?.ToLower() == "pdf";
                string fileExtension = isPdf ? "pdf" : "docx";
                string baseName = policyNumbers.Count == 1 ? $"Policy_{policyNumbers[0]}" : $"Policies_{policyNumbers.Count}_Documents";

                // Generate main policy document
                byte[] mainDocBytes = GenerateDocumentBytes(documentStream, dataSet, model, isPdf);

                // Return the document directly
                string mimeType = isPdf ? "application/pdf"
                    : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

                return File(mainDocBytes, mimeType, $"{baseName}.{fileExtension}");
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

                // Insert documents at bookmark locations after mail merge
                if (model.UseBookmarkDocuments && model.BookmarkDocuments != null && model.BookmarkDocuments.Count > 0)
                {
                    InsertBookmarkDocuments(document, model);
                }

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
        /// Inserts documents by finding text in main document, creating bookmark, then inserting content
        /// Document file names are used to search for matching text in the document
        /// </summary>
        private void InsertBookmarkDocuments(WordDocument document, ReportDataViewModel model)
        {
            if (!model.UseBookmarkDocuments || model.BookmarkDocuments == null || model.BookmarkDocuments.Count == 0)
                return;

            foreach (var bookmarkDocFile in model.BookmarkDocuments)
            {
                if (bookmarkDocFile.Length == 0)
                    continue;

                // Extract document name from filename (without extension) - this is what to search for
                string searchText = Path.GetFileNameWithoutExtension(bookmarkDocFile.FileName);

                try
                {
                    // Find ALL text occurrences in document using Find API
                    TextSelection[] textSelections = document.FindAll(searchText, false, false);

                    if (textSelections == null || textSelections.Length == 0)
                    {
                        _logger.LogWarning($"Text '{searchText}' not found in main document. Skipping document '{bookmarkDocFile.FileName}'");
                        continue;
                    }

                    _logger.LogInformation($"Found {textSelections.Length} occurrence(s) of '{searchText}'");

                    // Process in REVERSE order to avoid index shifting issues
                    for (int i = textSelections.Length - 1; i >= 0; i--)
                    {
                        TextSelection textSelection = textSelections[i];
                        WParagraph targetParagraph = textSelection.GetAsOneRange().OwnerParagraph;

                        // Get the ACTUAL container body
                        WTextBody body = targetParagraph.OwnerTextBody;

                        // Find index inside that body
                        int index = body.ChildEntities.IndexOf(targetParagraph);

                        // Insert new paragraph in SAME container
                        WParagraph newParagraph = new WParagraph(document);
                        body.ChildEntities.Insert(index + 1, newParagraph);

                        // Create bookmark on the new paragraph
                        string bookmarkName = $"InsertedDoc_{Guid.NewGuid().ToString().Substring(0, 8)}";
                        newParagraph.AppendBookmarkStart(bookmarkName);
                        newParagraph.AppendBookmarkEnd(bookmarkName);

                        _logger.LogInformation($"Created bookmark '{bookmarkName}' after text '{searchText}' (Occurrence {i + 1})");

                        // Insert document content at the bookmark
                        using (var ms = new MemoryStream())
                        {
                            bookmarkDocFile.CopyTo(ms);
                            ms.Position = 0;
                            InsertDocumentAtBookmark(document, bookmarkName, ms);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"Could not insert document for '{searchText}'");
                }
            }
        }

        /// <summary>
        /// Inserts document content at a specific bookmark location
        /// </summary>
        private void InsertDocumentAtBookmark(WordDocument mainDoc, string bookmarkName, Stream documentStream)
        {
            using (WordDocument bookmarkDoc = new WordDocument(documentStream, FormatType.Automatic))
            {
                // Create TextBodyPart from bookmark document
                TextBodyPart bodyPart = new TextBodyPart(bookmarkDoc);

                // Extract all content from bookmark document
                foreach (IWSection section in bookmarkDoc.Sections)
                {
                    foreach (IEntity entity in section.Body.ChildEntities)
                    {
                        bodyPart.BodyItems.Add(entity.Clone());
                    }
                }

                // Navigate to bookmark and insert content
                BookmarksNavigator navigator = new BookmarksNavigator(mainDoc);
                navigator.MoveToBookmark(bookmarkName, true, true);
                navigator.ReplaceBookmarkContent(bodyPart);

                _logger.LogInformation($"Successfully inserted document content at bookmark '{bookmarkName}'");
            }
        }

        /// <summary>
        /// OPTIMIZED: Generates separate documents using merge-once-then-split strategy.
        /// This approach is simpler, faster, and easier to understand:
        /// 1. Reads Excel once and creates DataSet
        /// 2. Performs mail merge once with ALL data (each record starts on new page)
        /// 3. Splits the merged document by page breaks
        /// 4. Returns ZIP archive with separate files
        /// </summary>
        private IActionResult GenerateSeparateDocuments(Stream documentStream, DataSet dataSet,Stream excelStream, List<string> policyNumbers, ReportDataViewModel model)
        {
            try
            {
                bool isPdf = model.OutputFormat?.ToLower() == "pdf";
                string fileExtension = isPdf ? "pdf" : "docx";

                using (MemoryStream zipStream = new MemoryStream())
                {
                    using (var archive = new System.IO.Compression.ZipArchive(zipStream, System.IO.Compression.ZipArchiveMode.Create, true))
                    {
                        // Step 2: Generate main policy documents (merge once, split by page breaks)
                        byte[] docsZip = GenerateDocumentsWithPageBreakSplit(
                            documentStream, dataSet, policyNumbers, model, isPdf, fileExtension);

                        // Step 3: Add main documents to final ZIP
                        AddZipContentsToArchive(docsZip, archive);                      
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
            using (WordDocument document = new WordDocument(templateStream, FormatType.Automatic))
            {
                // KEY: Ensure each record starts on a new page (creates page breaks automatically)
                document.MailMerge.StartAtNewPage = true;

                // Execute mail merge ONCE with all data
                ExecuteMailMerge(document, dataSet);

                // Insert documents at bookmark locations after mail merge
                if (model.UseBookmarkDocuments && model.BookmarkDocuments != null && model.BookmarkDocuments.Count > 0)
                {
                    InsertBookmarkDocuments(document, model);
                }

                // Split the merged document by page breaks and return as ZIP
                return SplitDocumentByPageBreaks(document, policyNumbers, isPdf, fileExtension, fileNameSuffix, model.MultiPartDocument);
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
            string fileNameSuffix,
             bool addCoverPage = false)
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
                                if(addCoverPage)
                                    AddCoverPageAndTOC(extractedDoc);
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
            document.Sections.Insert(0, firstSection);

            // === TITLE ===
            IWParagraph titlePara = firstSection.AddParagraph();
            titlePara.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
            titlePara.ParagraphFormat.BeforeSpacing = 72; // ~1 inch top margin

            IWTextRange titleText = titlePara.AppendText("POLICY DOCUMENT");
            titleText.CharacterFormat.FontSize = 28;
            titleText.CharacterFormat.Bold = true;
            titleText.CharacterFormat.FontName = "Arial";

            titlePara.AppendBreak(BreakType.LineBreak);

            // === SUBTITLE ===
            IWTextRange subtitleText = titlePara.AppendText("Insurance Coverage Summary");
            subtitleText.CharacterFormat.FontSize = 14;
            subtitleText.CharacterFormat.Italic = true;
            subtitleText.CharacterFormat.TextColor = Color.FromArgb(68, 114, 196);

            // === SPACING ===
            firstSection.AddParagraph().AppendBreak(BreakType.LineBreak);
            firstSection.AddParagraph().AppendBreak(BreakType.LineBreak);

            // === DOCUMENT INFO ===
            IWParagraph infoPara = firstSection.AddParagraph();
            infoPara.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;

            IWTextRange dateLabel = infoPara.AppendText("Generated Date: ");
            dateLabel.CharacterFormat.FontSize = 12;
            dateLabel.CharacterFormat.Bold = true;

            IWTextRange dateValue = infoPara.AppendText($"{DateTime.Now:MMMM dd, yyyy}");
            dateValue.CharacterFormat.FontSize = 12;

            infoPara.AppendBreak(BreakType.LineBreak);

            IWTextRange timeValue = infoPara.AppendText($"{DateTime.Now:hh:mm tt}");
            timeValue.CharacterFormat.FontSize = 10;
            timeValue.CharacterFormat.TextColor = Color.Gray;

            // === SPACING ===
            firstSection.AddParagraph().AppendBreak(BreakType.LineBreak);

            // === DISCLAIMER ===
            IWParagraph disclaimerPara = firstSection.AddParagraph();
            disclaimerPara.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;
            disclaimerPara.ParagraphFormat.LineSpacing = 14;

            IWTextRange disclaimer = disclaimerPara.AppendText(
                "This document contains confidential policy information.\n" +
                "Please review all terms and conditions carefully.\n" +
                "For questions, contact your insurance provider.");
            disclaimer.CharacterFormat.FontSize = 10;
            disclaimer.CharacterFormat.Italic = true;
            disclaimer.CharacterFormat.TextColor = Color.FromArgb(89, 89, 89);

            // === FOOTER LINE ===
            firstSection.AddParagraph().AppendBreak(BreakType.LineBreak);

            IWParagraph footerLine = firstSection.AddParagraph();
            footerLine.ParagraphFormat.HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment.Center;

            IWTextRange footer = footerLine.AppendText("_______________________________________");
            footer.CharacterFormat.FontSize = 10;
            footer.CharacterFormat.TextColor = Color.LightGray;

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

                // Read all sheets into DataTables
                foreach (IWorksheet sheet in workbook.Worksheets)
                {
                    if (sheet.UsedRange == null || sheet.UsedRange.LastRow < 2)
                        continue;

                    DataTable dt = ReadExcelSheetToDataTable(
                        workbook,
                        sheet.Name,
                        policyNumbers);

                    if (dt != null && dt.Rows.Count > 0)
                    {
                        dataSet.Tables.Add(dt);
                    }
                }
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

            // Check if any column name contains "policy" (case-insensitive)
            bool hasPolicyColumn = dt.Columns.Cast<DataColumn>()
                .Any(c => c.ColumnName.Contains("policy", StringComparison.OrdinalIgnoreCase));

            // Determine if filtering should be applied
            bool shouldFilter = hasPolicyColumn &&
                                policyNumbers != null &&
                                policyNumbers.Count > 0;

            // Find the policy column index if filtering is needed
            int policyColumnIndex = -1;
            if (shouldFilter)
            {
                var policyColumn = dt.Columns.Cast<DataColumn>()
                    .FirstOrDefault(c => c.ColumnName.Contains("policy", StringComparison.OrdinalIgnoreCase));
                policyColumnIndex = dt.Columns.IndexOf(policyColumn);
            }

            // Add data rows (skip header)
            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                // Apply filter if table has policy column and specific policies are selected
                if (shouldFilter && policyColumnIndex >= 0)
                {
                    string policyValue = sheet[row, policyColumnIndex + 1].Value?.ToString()?.Trim();

                    if (string.IsNullOrWhiteSpace(policyValue) ||
                        !policyNumbers.Any(p => p.Equals(policyValue, StringComparison.OrdinalIgnoreCase)))
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
        /// Executes mail merge based on data structure (flat, nested, or multiple tables)
        /// Maintains parent-child relationships for nested groups
        /// </summary>
        public void ExecuteMailMerge(WordDocument document, DataSet dataSet)
        {
            if (dataSet.Tables.Count == 0)
                return;

            try
            {
                document.MailMerge.StartAtNewPage = true;

                var groupNames = document.MailMerge.GetMergeGroupNames();

                // Case 1: No groups in template AND single table → Simple Execute
                if ((groupNames == null || groupNames.Length == 0) && dataSet.Tables.Count == 1)
                {
                    document.MailMerge.StartAtNewPage = false;
                    document.MailMerge.Execute(dataSet.Tables[0]);
                }
                // Case 2: Template has groups → ExecuteNestedGroup
                else if (groupNames != null && groupNames.Length > 0)
                {
                    ExecuteNestedGroupWithSmartCommands(document, dataSet, groupNames);
                }
                // Update document fields
                document.UpdateDocumentFields();
            }
            catch (Exception ex)
            {
                throw new Exception($"Mail merge failed: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Executes a nested mail merge operation using the provided document,
        /// dataset, and group names. Builds smart commands dynamically for nested groups.
        /// </summary>

        private void ExecuteNestedGroupWithSmartCommands(WordDocument document, DataSet dataSet, string[] groupNames)
        {
            try
            {
                // List to hold dynamically built nested group command
                ArrayList commands = new ArrayList();

                // Simply build commands based on group names
                // If groupNames is null/empty, commands will be empty and DocIO handles it
                if (groupNames != null && groupNames.Length > 0)
                {
                    BuildNestedCommands(document, dataSet, groupNames, commands);
                }
                // Execute the nested mail merge with the generated commands
                document.MailMerge.ExecuteNestedGroup(dataSet, commands);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error executing nested group merge: {ex.Message}", ex);
            }
        }
        /// <summary>
        /// Builds nested mail merge commands dynamically based on the provided group names.
        /// Each group represents a table in the dataset, and relationships are built sequentially.
        /// </summary
        private void BuildNestedCommands(WordDocument document, DataSet dataSet, string[] groupNames, ArrayList commands)
        {
            // Exit early if no group names provided
            if (groupNames == null || groupNames.Length == 0)
                return;
            // Get the first (root) group name
            string firstGroupName = groupNames[0];
            // Find matching table in dataset (case-insensitive
            var firstTable = dataSet.Tables.Cast<DataTable>()
                .FirstOrDefault(t => string.Equals(t.TableName, firstGroupName, StringComparison.OrdinalIgnoreCase));
            // If root table not found, log warning and stop processing
            if (firstTable == null)
            {
                _logger.LogWarning($"First group '{firstGroupName}' not found in dataset.");
                return;
            }

            // Add root parent with empty relation (like "Employees" in your example)
            commands.Add(new DictionaryEntry(firstTable.TableName, string.Empty));

            // Track current parent - changes at each level
            string currentParentTableName = firstTable.TableName;

            // Add child tables - each relates to IMMEDIATE parent (not root parent)
            for (int i = 1; i < groupNames.Length; i++)
            {
                string childGroupName = groupNames[i];
                var childTable = dataSet.Tables.Cast<DataTable>()
                    .FirstOrDefault(t => string.Equals(t.TableName, childGroupName, StringComparison.OrdinalIgnoreCase));

                if (childTable != null)
                {
                    // Build relation to immediate parent (previous table)
                    // Example: Orders relates to Customers (not to Employees)
                    string relationString = BuildRelationString(dataSet, currentParentTableName, childTable.TableName);

                    commands.Add(new DictionaryEntry(childTable.TableName, relationString));

                    if (string.IsNullOrEmpty(relationString))
                    {
                        _logger.LogInformation($"No common column found between '{currentParentTableName}' and '{childTable.TableName}'. Using empty relation.");
                    }

                    // Update parent for next iteration
                    // Next child will relate to THIS table
                    currentParentTableName = childTable.TableName;
                }
                else
                {
                    _logger.LogWarning($"Child group '{childGroupName}' not found in dataset.");
                }
            }
        }

        /// <summary>
        /// Builds a relation string between a parent and child table based on a common column name.
        /// This relation format is used by DocIO for nested mail merge operations.
        /// </summary>

        private string BuildRelationString(DataSet dataSet, string parentTableName, string childTableName)
        {
            // Retrieve parent and child tables
            DataTable parentTable = dataSet.Tables[parentTableName];
            DataTable childTable = dataSet.Tables[childTableName];

            // Return empty if either table is missing
            if (parentTable == null || childTable == null)
                return string.Empty;

            // Find common column
            var parentColumns = parentTable.Columns
                .Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var childColumns = childTable.Columns
                .Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            // Find the first common column between parent and child
            var commonKey = parentColumns
                .Intersect(childColumns)
                .FirstOrDefault();
            // If no common column found, return empty relation
            if (commonKey == null)
                return string.Empty;

            // DocIO relation format
            return $"{commonKey} = %{parentTableName}.{commonKey}%";
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
                string[] supportedExtensions = { ".doc", ".docx", ".dot", ".dotx", ".dotm", ".docm", ".xml", ".rtf", ".html", ".md" };
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
    }

}
