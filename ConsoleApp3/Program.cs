using ConsoleApp3;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

string templatePath = @"C:\Users\togi\Documents\ST - Copy.docx";
using WordprocessingDocument document = WordprocessingDocument.Open(templatePath, true);

ReportExtension.GenerateReport(document);