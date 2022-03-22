using DHIReportExtension;
using DocumentFormat.OpenXml.Packaging;

var copyPath = CleanTemplate();

using WordprocessingDocument document = WordprocessingDocument.Open(copyPath, true);

document.GenerateReport();









static string CleanTemplate()
{
    string originalPath = @"C:\Users\togi\Documents\ST.docx";
    string copyPath = @"C:\Users\togi\Documents\ST - Copy.docx";

    File.Delete(copyPath);
    File.Copy(originalPath, copyPath);

    return copyPath;
}