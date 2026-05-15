using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Policy_Document_Generation.Models
{
    public class ReportDataViewModel
    {
        /// <summary>
        /// Word or PowerPoint template file uploaded by user
        /// </summary>
        [Display(Name = "Template File")]
        public IFormFile TemplateFile { get; set; }

        /// <summary>
        /// Excel file containing data for mail merge
        /// </summary>
        [Display(Name = "Excel Data File")]
        public IFormFile ExcelDataFile { get; set; }

        /// <summary>
        /// Type of document to generate (Word, PowerPoint, PDF)
        /// </summary>
        [Display(Name = "Output Format")]
        public string OutputFormat { get; set; }

        /// <summary>
        /// Selection mode: "specific" for specific policies or "all" for all policies
        /// </summary>
        [Display(Name = "Selection Mode")]
        public string SelectionMode { get; set; } = "specific";

        /// <summary>
        /// Policy numbers entered by user (one per line or comma-separated)
        /// </summary>
        [Display(Name = "Policy Numbers")]
        public string PolicyNumbers { get; set; }

        /// <summary>
        /// Generate multi-part document with cover page and table of contents
        /// </summary>
        [Display(Name = "Multi-part Document")]
        public bool MultiPartDocument { get; set; }

        /// <summary>
        /// Generate separate files for each policy instead of combined file
        /// </summary>
        [Display(Name = "Generate Separate Files")]
        public bool GenerateSeparateFiles { get; set; }

        /// <summary>
        /// Enable insertion of documents at bookmark locations in the template
        /// </summary>
        [Display(Name = "Use Bookmark Documents")]
        public bool UseBookmarkDocuments { get; set; }

        /// <summary>
        /// Documents to be inserted at bookmark locations
        /// File names must match bookmark names in the template
        /// Supported formats: .docx, .rtf, .doc, .md, .txt, .html
        /// </summary>
        [Display(Name = "Bookmark Documents")]
        public IFormFileCollection BookmarkDocuments { get; set; }

    }
}
