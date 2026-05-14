using System.ComponentModel.DataAnnotations;

namespace Policy_Document_Generation.Models
{
    public class ReportDataViewModel
    {
        /// <summary>
        /// Word or PowerPoint template file uploaded by user
        /// </summary>
        [Display(Name = "Template File")]
        public IFormFile? TemplateFile { get; set; }

        /// <summary>
        /// Excel file containing data for mail merge
        /// </summary>
        [Display(Name = "Excel Data File")]
        [Required(ErrorMessage = "Excel data file is required")]
        public IFormFile? ExcelDataFile { get; set; }

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
        public string? PolicyNumbers { get; set; }

        /// <summary>
        /// Type of additional document to generate alongside the main policy document.
        /// Options: "claims-letter", "endorsement", "certificate", "renewal-notice", or empty for none.
        /// </summary>
        [Display(Name = "Additional Document Type")]
        public string? DocumentType { get; set; }

        /// <summary>
        /// Optional separate Excel file containing data specifically for the additional document.
        /// If not provided, the system uses the main ExcelDataFile (looks for a matching sheet by document type).
        /// </summary>
        [Display(Name = "Additional Document Data File")]
        public IFormFile? AdditionalDocExcelFile { get; set; }

        /// <summary>
        /// Optional custom Word template for the additional document.
        /// If not provided, the system uses the built-in default template for the selected document type.
        /// </summary>
        [Display(Name = "Additional Document Template")]
        public IFormFile? AdditionalDocTemplateFile { get; set; }

        /// <summary>
        /// Extra features flags
        /// </summary>
        public bool MultiPartDocument { get; set; }
        public bool GenerateSeparateFiles { get; set; }
    }
}
