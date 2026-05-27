# Syncfusion ASP.NET Core – Policy Document Generation Demo

This repository contains a complete showcase sample demonstrating how to build an **Automated Policy Document Generation System** using **Syncfusion DocIO** and **Syncfusion Excel** , **Syncfusion PDF** and **Syncfusion Smart Data Extractor** libraries in an ASP.NET Core MVC application. The sample illustrates how insurance professionals can streamline policy document creation by merging Excel data with Word templates, supporting both single and bulk document generation with mail merge capabilities.

---

## 📁 Project Structure

```
├── Controllers/
│   └── HomeController.cs
├── Models/
│   └── ReportDataViewModel.cs
├── Views/
│   ├── Home/
│   │   └── Index.cshtml
│   └── Shared/
├── wwwroot/
│   └── Data/
│       ├── Template.docx
│       ├── PolicyData.xlsx
└── README.md
```

---

## 🚀 Getting Started

### Prerequisites

- [.NET 7.0 SDK](https://dotnet.microsoft.com/download) or later
- Visual Studio 2022 or VS Code
- A valid **Syncfusion License Key** (or use the free Community License)

### 1. Clone the Repository

```bash
git clone https://github.com/SyncfusionExamples/DocIO-Policy-Document-Generation
cd Policy-Document-Generation
```

### 2. Install Dependencies

Restore all NuGet packages:

```bash
dotnet restore
```

### 3. Add Syncfusion License Key

In your `Program.cs`, register your Syncfusion license:

```csharp
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY");
```

### 4. Run the Application

```bash
dotnet run
```

Open your browser and navigate to:
```
https://localhost:5001
```

---

## 📋 How to Use

### Basic Workflow

1. **Upload Template Document**
   - Upload your Word template (.docx, .doc, .rtf) with merge fields
   - Drag & drop supported
   - Or use default template provided

2. **Upload Data File**
   - Upload Excel file (.xlsx, .xls, .csv, .xlts, .xlsm) containing policy data
   - Automatic policy number detection
   - Or use default data file

3. **Select Policies**
   - **Specific Policies** – Enter policy numbers (one per line or comma-separated)
   - **All Policies** – Generate documents for all records in Excel
   - Note: Data file must contain a Policy Number column for specific policy filtering

4. **Choose Output Format**
   - **DOCX** – Editable Word document
   - **PDF** – Distribution-ready format

5. **Configure Extra Features**
   - **Multi-part Document** – Add cover page & table of contents
   - **Separate Files** – Generate individual files per policy (ZIP archive)
   - **Insert Documents** – Embed additional documents at specific locations
   - **Digital Signature** *(PDF only)* 
     - Upload your signature image (PNG, JPG, etc.)
     - Enter keywords where you want signatures placed (e.g., "Signature, WITNESS")
     - Signatures will automatically appear above those keywords in the PDF

6. **Generate Documents**
   - Click "Generate Documents" button
   - Download automatically starts
   - Single file or ZIP archive based on selection

---

## 🎯 Use Cases

### Insurance Document Scenarios

- **Policy Certificate Generation** – Create personalized policy documents from template and data
- **Bulk Policy Printing** – Generate multiple policy documents in a single operation
- **Renewal Notices** – Merge policy data with renewal letter templates
- **Claims Documentation** – Create claim documents with dynamic data
- **Endorsement Generation** – Produce policy modification documents with updated terms

---

## 🔗 Resources

- [Syncfusion DocIO Getting Started](https://help.syncfusion.com/document-processing/word/word-library/net/getting-started)
- [Mail Merge](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-mailmerge)
- [Syncfusion Excel Getting Started](https://help.syncfusion.com/document-processing/excel/excel-library/net/overview)
- [Word to PDF Conversion](https://help.syncfusion.com/document-processing/word/conversions/word-to-pdf/overview)
- [Find and Replace](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-find-and-replace)
- [Work with Bookmarks](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-bookmarks)
- [Smart Data Extractor](https://help.syncfusion.com/document-processing/data-extraction/smart-data-extractor/overview)

---

## 📣 Try It Out

Clone the repository, run the sample, and discover how **Syncfusion DocIO** can revolutionize policy document generation in your insurance operations.

### Customization

This sample application is provided as a reference implementation and can be freely customized to suit your specific business requirements.

You can modify the templates, data sources, merge logic, and output formats based on your use case. If you have any questions, need clarification, or require assistance while customizing this sample, please feel free to contact our [Syncfusion Support Team](https://support.syncfusion.com/support/tickets/create) for guidance.

---

## 📄 License and Copyright

> This is a commercial product and requires a paid license for possession or use. Syncfusion® licensed software, including this component, is subject to the terms and conditions of Syncfusion®. To acquire a license, visit https://www.syncfusion.com/account/downloads.

Are you already a Syncfusion user? You can download the product setup [here](https://www.syncfusion.com/account/downloads). If you're not yet a Syncfusion user, you can download a [30-day free trial](https://www.syncfusion.com/downloads).

---

## 📞 Support

For technical support and questions:
- [Syncfusion Support Portal](https://support.syncfusion.com/support/tickets/create)
- [Documentation](https://help.syncfusion.com/)
- [Community Forums](https://www.syncfusion.com/forums)

---