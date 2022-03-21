using DHIReportExtension;
using DocumentFormat.OpenXml.Packaging;

string templatePath = @"C:\Users\togi\Documents\ST - Copy.docx";

using WordprocessingDocument document = WordprocessingDocument.Open(templatePath, true);

document.GenerateReport();