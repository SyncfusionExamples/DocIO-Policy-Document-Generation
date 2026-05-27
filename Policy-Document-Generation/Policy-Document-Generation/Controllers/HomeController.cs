using Microsoft.AspNetCore.Mvc;
using Policy_Document_Generation.Models;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Pdf.Security;
using Syncfusion.SmartDataExtractor;
using Syncfusion.XlsIO;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.IO.Compression;

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

                // Step 4: Create DataSet (filter during read if specific policies selected)
                DataSet dataSet = CreateMailMergeDataSet(excelStream, policyNumbersList);

                if (dataSet.Tables.Count == 0 || dataSet.Tables[0].Rows.Count == 0)
                {
                    TempData["Error"] = "No data found in Excel file.";
                    return RedirectToAction("Index");
                }

                // Step 5: For "all with separate files" - extract policy numbers from DataTable
                if (model.SelectionMode == "all" && model.GenerateSeparateFiles)
                {
                    policyNumbersList = dataSet.Tables[0].AsEnumerable()
                        .Select(row => row[0]?.ToString()?.Trim())
                        .Where(p => !string.IsNullOrWhiteSpace(p))
                        .Distinct()
                        .ToList();
                }

                // Step 6: Generate documents using unified method
                return CreateDocuments(documentStream, dataSet, policyNumbersList, model);
            }
            catch (Exception ex)
            {
                TempData["Error"] = $"Error generating document: {ex.Message}";
                return RedirectToAction("Index");
            }
        }
        /// <summary>
        /// Unified document generation method supporting both single and multiple file outputs.
        /// Performs mail merge once, then either returns single file or splits by page breaks into ZIP.
        /// </summary>
        private IActionResult CreateDocuments(Stream documentStream, DataSet dataSet, List<string> policyNumbers, ReportDataViewModel model)
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

                using (WordDocument document = new WordDocument(documentStream, FormatType.Automatic))
                {
                    // Execute mail merge once with all data
                    ExecuteMailMerge(document, dataSet);

                    // Insert bookmark documents if enabled
                    if (model.UseBookmarkDocuments && model.BookmarkDocuments != null && model.BookmarkDocuments.Count > 0)
                    {
                        InsertPlaceholderDocuments(document, model);
                    }

                    if (model.GenerateSeparateFiles && policyNumbers.Count > 1)
                    {
                        // Split into multiple documents
                        byte[] zipBytes = SplitDocumentByPageBreaks(
                            document,
                            policyNumbers,
                            isPdf,
                            fileExtension,
                            model.MultiPartDocument,
                            model.SignatureImage,
                            model.SignatureKeywords,
                            model.EnableDigitalSign);

                        if (zipBytes != null && zipBytes.Length > 0)
                        {
                            // Return ZIP file with all separated documents (signatures already applied)
                            return File(zipBytes, "application/zip",
                                $"Policy_Documents_{DateTime.Now:yyyyMMdd_HHmmss}.zip");
                        }
                        else
                        {
                            // Fallback: return as single document if split fails
                            TempData["Warning"] = "Unable to split documents. Returning as single file.";
                            using (MemoryStream outputStream = new MemoryStream())
                            {
                                SaveDocumentToStream(document, outputStream, isPdf);

                                // Apply digital signatures if enabled and PDF format
                                if (isPdf && model.EnableDigitalSign)
                                {
                                    outputStream.Position = 0;
                                    using (MemoryStream signedStream = ApplyDigitalSignatureIfEnabled(
                                        outputStream,
                                        model.SignatureImage,
                                        model.SignatureKeywords,
                                        true))
                                    {
                                        return File(signedStream.ToArray(),
                                            isPdf ? "application/pdf" : "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                            $"Policy_Documents.{fileExtension}");
                                    }
                                }
                                else
                                {
                                    return File(outputStream.ToArray(),
                                        isPdf ? "application/pdf" : "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        $"Policy_Documents.{fileExtension}");
                                }
                            }
                        }
                    }
                    // OUTPUT: Single merged document
                    else
                    {
                        if (model.MultiPartDocument)
                        {
                            AddCoverPage(document);
                        }

                        using (MemoryStream outputStream = new MemoryStream())
                        {
                            SaveDocumentToStream(document, outputStream, isPdf);
                            // Apply digital signatures if enabled and PDF format
                            if (isPdf && model.EnableDigitalSign)
                            {
                                outputStream.Position = 0;
                                using (MemoryStream signedStream = ApplyDigitalSignatureIfEnabled(
                                    outputStream,
                                    model.SignatureImage,
                                    model.SignatureKeywords,
                                    true))
                                {
                                    string fileName = policyNumbers.Count == 1
                                        ? $"Policy_{policyNumbers[0]}.pdf"
                                        : $"Policies_{policyNumbers.Count}_Documents.pdf";

                                    return File(signedStream.ToArray(), "application/pdf", fileName);
                                }
                            }
                            else
                            {
                                string fileName = policyNumbers.Count == 1
                                    ? $"Policy_{policyNumbers[0]}.{fileExtension}"
                                    : $"Policies_{policyNumbers.Count}_Documents.{fileExtension}";

                                string mimeType = isPdf
                                    ? "application/pdf"
                                    : "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

                                return File(outputStream.ToArray(), mimeType, fileName);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                TempData["Error"] = $"Error generating documents: {ex.Message}";
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
                    InsertPlaceholderDocuments(document, model);
                }

                if (model.MultiPartDocument)
                {
                    AddCoverPage(document);
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
        private void InsertPlaceholderDocuments(WordDocument document, ReportDataViewModel model)
        {
            if (!model.UseBookmarkDocuments || model.BookmarkDocuments == null || model.BookmarkDocuments.Count == 0)
                return;

            foreach (IFormFile bookmarkDocFile in model.BookmarkDocuments)
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
                        using (MemoryStream ms = new MemoryStream())
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
                WordDocumentPart wordDocumentPart = new WordDocumentPart(bookmarkDoc);
                // Navigate to bookmark and insert content
                BookmarksNavigator navigator = new BookmarksNavigator(mainDoc);
                navigator.MoveToBookmark(bookmarkName, true, true);
                navigator.ReplaceContent(wordDocumentPart);

                _logger.LogInformation($"Successfully inserted document content at bookmark '{bookmarkName}'");
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
            bool addCoverPage = false,
            IFormFile signatureImage = null,
            string signatureKeywords = null,
            bool enableDigitalSign = false)
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
                using (System.IO.Compression.ZipArchive archive = new System.IO.Compression.ZipArchive(
                    zipStream, System.IO.Compression.ZipArchiveMode.Create, true))
                {
                    for (int i = 1; i <= bookmarkIndex && i <= policyNumbers.Count; i++)
                    {
                        string policyNumber = policyNumbers[i - 1];

                        // Build filename: "Policy_12345.pdf"
                        string fileName = $"Policy_{policyNumber}.{fileExtension}";

                        try
                        {
                            // Navigate to bookmark and extract content
                            BookmarksNavigator navigator = new BookmarksNavigator(document);
                            navigator.MoveToBookmark($"Policy_Section_{i}", true, true);
                            WordDocumentPart documentPart = navigator.GetContent();

                            if (documentPart == null) continue;

                            // Extract content as new WordDocument
                            using (WordDocument extractedDoc = documentPart.GetAsWordDocument())
                            {
                                if (addCoverPage)
                                {
                                    AddCoverPage(extractedDoc);
                                }

                                if (convertToPdf)
                                {
                                    // Convert to PDF
                                    using (Syncfusion.DocIORenderer.DocIORenderer renderer = new Syncfusion.DocIORenderer.DocIORenderer())
                                    using (Syncfusion.Pdf.PdfDocument pdfDocument = renderer.ConvertToPDF(extractedDoc))
                                    using (MemoryStream pdfStream = new MemoryStream())
                                    {
                                        pdfDocument.Save(pdfStream);
                                        pdfDocument.Close();
                                        pdfStream.Position = 0;

                                        // Apply digital signatures to each PDF if enabled
                                        using (MemoryStream signedPdfStream = ApplyDigitalSignatureIfEnabled(
                                            pdfStream, signatureImage, signatureKeywords, enableDigitalSign))
                                        {
                                            // Add signed PDF to ZIP
                                            ZipArchiveEntry entry = archive.CreateEntry(fileName, System.IO.Compression.CompressionLevel.Fastest);
                                            using (Stream entryStream = entry.Open())
                                            {
                                                signedPdfStream.Position = 0;
                                                signedPdfStream.CopyTo(entryStream);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    // Save as Word document
                                    using (MemoryStream docStream = new MemoryStream())
                                    {
                                        extractedDoc.Save(docStream, FormatType.Docx);
                                        docStream.Position = 0;

                                        // Add Word document to ZIP
                                        ZipArchiveEntry entry = archive.CreateEntry(fileName, System.IO.Compression.CompressionLevel.Fastest);
                                        using (Stream entryStream = entry.Open())
                                        {
                                            docStream.CopyTo(entryStream);
                                        }
                                    }
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
            using (System.IO.Compression.ZipArchive sourceArchive = new System.IO.Compression.ZipArchive(
                zipStream, System.IO.Compression.ZipArchiveMode.Read))
            {
                foreach (ZipArchiveEntry entry in sourceArchive.Entries)
                {
                    ZipArchiveEntry newEntry = targetArchive.CreateEntry(entry.FullName);
                    using (Stream sourceStream = entry.Open())
                    using (Stream targetStream = newEntry.Open())
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
        private void AddCoverPage(WordDocument document)
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

            // Find policy column using case-insensitive comparison (handles PolicyNumber, Policy Number, policy_number, etc.)
            DataColumn policyColumn = dt.Columns.Cast<DataColumn>()
                .FirstOrDefault(c => c.ColumnName.Replace(" ", "").Replace("_", "").Equals("policynumber", StringComparison.OrdinalIgnoreCase));

            bool shouldFilter = policyColumn != null && policyNumbers != null && policyNumbers.Count > 0;
            int policyColumnIndex = shouldFilter ? dt.Columns.IndexOf(policyColumn) : -1;

            // Add data rows (skip header)
            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                // Apply filter if needed
                if (shouldFilter)
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

                string[] groupNames = document.MailMerge.GetMergeGroupNames();

                // Case 1: No groups in template AND single table → Simple Execute
                if ((groupNames == null || groupNames.Length == 0) && dataSet.Tables.Count == 1)
                {
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
                BuildNestedCommands(document, dataSet, groupNames, commands);
                
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
            DataTable firstTable = dataSet.Tables.Cast<DataTable>()
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
                DataTable childTable = dataSet.Tables.Cast<DataTable>()
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
            HashSet<string> parentColumns = parentTable.Columns
                .Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            HashSet<string> childColumns = childTable.Columns
                .Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            // Find the first common column between parent and child
            string commonKey = parentColumns
                .Intersect(childColumns)
                .FirstOrDefault();
            // If no common column found, return empty relation
            if (commonKey == null)
                return string.Empty;

            // DocIO relation format
            return $"{commonKey} = %{parentTableName}.{commonKey}%";
        }
        /// <summary>
        /// Conditionally applies digital signatures to PDF based on enableDigitalSign flag
        /// Returns PDF stream directly if signature is not needed for optimal performance
        /// </summary>
        private MemoryStream ApplyDigitalSignatureIfEnabled(MemoryStream inputStream, IFormFile signatureImage, string signatureKeywordsInput, bool enableDigitalSign)
        {
            Stream signatureStream = null;
            try
            {
                // Step 1: Early exit if digital signature is not enabled - return PDF stream directly
                if (!enableDigitalSign)
                {
                    _logger.LogInformation("Digital signature is not enabled. Returning PDF stream directly.");
                    inputStream.Position = 0;
                    return inputStream;
                }
                // Step 2: Check if signature image is available
                signatureStream = GetSignatureImageStream(signatureImage);
                if (signatureStream == null)
                {
                    _logger.LogWarning("Digital signature enabled but no signature image found. Returning PDF stream directly.");
                    inputStream.Position = 0;
                    return inputStream;
                }
                // Step 3: Only use ApplyDigitalSignatureIfEnabled when signature is actually needed
                _logger.LogInformation("Applying digital signatures using ApplyDigitalSignatureIfEnabled.");
                // Initialize the extractor with required detection settings
                var extractor = new DataExtractor { EnableFormDetection = false, EnableTableDetection = true, ConfidenceThreshold = 0.6 };
                // Extract PDF document from the input stream
                inputStream.Position = 0;
                PdfLoadedDocument PDFdocument = extractor.ExtractDataAsPdfDocument(inputStream);
                // Step 4: Apply signatures
                AddSignaturesToPDFDocument(PDFdocument, signatureStream, signatureKeywordsInput);
                // Step 5: Save PDF with signatures
                var outputMs = new MemoryStream();
                PDFdocument.Save(outputMs);
                PDFdocument.Close(true);
                outputMs.Position = 0;
                return outputMs;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing PDF document");
                throw;
            }
            finally
            {
                // Clean up signature stream
                signatureStream?.Dispose();
            }
        }
        /// <summary>
        /// Adds digital signatures to PDF document at locations matching specified keywords
        /// </summary>
        private void AddSignaturesToPDFDocument(PdfLoadedDocument PDFdocument, Stream signatureImageStream, string keywords)
        {
            // Use default keywords if none are provided
            string[] signatureKeywords = string.IsNullOrWhiteSpace(keywords)
                ? new[] { "Sign", "WITNESS", "CompanySignature" }
                : keywords.Split(',').Select(k => k.Trim()).ToArray();

            _logger.LogInformation($"Applying digital signatures for keywords: {string.Join(", ", signatureKeywords)}");

            // Iterate through each page in the document
            for (int pageIndex = 0; pageIndex < PDFdocument.Pages.Count; pageIndex++)
            {
                PdfPageBase page = PDFdocument.Pages[pageIndex];
                TextLineCollection textLines;

                // Extract text lines from the page
                page.ExtractText(out textLines);

                // Iterate through each word on the page
                foreach (TextLine line in textLines.TextLine)
                {
                    foreach (TextWord word in line.WordCollection)
                    {
                        // Skip words that do not match any signature keyword
                        if (!signatureKeywords.Any(k => word.Text.Contains(k, StringComparison.Ordinal)))
                            continue;
                        // Calculate signature position above the keyword
                        RectangleF bounds = word.Bounds;
                        float signatureX = bounds.X;
                        float signatureY = bounds.Y - bounds.Height - 10;
                        float signatureWidth = 80;
                        float signatureHeight = 20;
                        try
                        {
                            // Load digital certificate
                            using System.IO.FileStream cert = new System.IO.FileStream(Path.GetFullPath("PDF.pfx"), System.IO.FileMode.Open, System.IO.FileAccess.Read);
                            PdfCertificate pdfCert = new PdfCertificate(cert, "syncfusion");

                            // Create and configure the PDF signature
                            PdfSignature signature = new PdfSignature(PDFdocument, page, pdfCert, "Signature");
                            signature.Bounds = new RectangleF(signatureX, signatureY, signatureWidth, signatureHeight);

                            // Load signature image directly from stream
                            signatureImageStream.Position = 0; // Reset stream position for each use
                            PdfBitmap signatureImageBitmap = new PdfBitmap(signatureImageStream);
                            signature.Appearance.Normal.Graphics.DrawImage(signatureImageBitmap, 0, 0, signatureWidth, signatureHeight);

                            _logger.LogDebug($"Signature added at page {pageIndex + 1}, position ({signatureX}, {signatureY})");
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, $"Error adding signature at page {pageIndex + 1}");
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Gets signature image as a stream from user upload or default signature
        /// </summary>
        private Stream GetSignatureImageStream(IFormFile signatureImage)
        {
            // If user provided an image, return its stream
            if (signatureImage != null && signatureImage.Length > 0)
            {
                _logger.LogInformation("Using user-provided signature image stream.");
                var memoryStream = new MemoryStream();
                signatureImage.OpenReadStream().CopyTo(memoryStream);
                memoryStream.Position = 0;
                return memoryStream;
            }
            // No user image - use default signature from project
            string defaultImagePath = Path.Combine(_hostingEnvironment.ContentRootPath, "Signature.png");
            if (System.IO.File.Exists(defaultImagePath))
            {
                _logger.LogInformation($"Using default signature image from: {defaultImagePath}");
                try
                {
                    var fileStream = new FileStream(defaultImagePath, FileMode.Open, FileAccess.Read);
                    var memoryStream = new MemoryStream();
                    fileStream.CopyTo(memoryStream);
                    fileStream.Dispose();
                    memoryStream.Position = 0;
                    return memoryStream;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"Error loading default signature image: {defaultImagePath}");
                    return null;
                }
            }
            _logger.LogWarning($"Default signature image not found at: {defaultImagePath}");
            return null;
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
    }

}
