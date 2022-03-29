using DHIReportExtension;
using DHIWordExtension.Interpreter;
using DocumentFormat.OpenXml.Packaging;
using WI = DHIReportExtension.WordElementArrangeImage;

Controller.InterpretingTest("dupa");

string documentCopyPath = @"C:\Users\togi\Documents\ST - Copy.docx";
CleanTemplate_DevelopmentTool(documentCopyPath);
using WordprocessingDocument document = WordprocessingDocument.Open(documentCopyPath, true);

var chartPath = NPlotLib.Plot(800, 400);
ReportExtension.InsertAPicture(document, chartPath);

document.GenerateReport();




static void CleanTemplate_DevelopmentTool(string destinationPath)
{
    string originalPath = @"C:\Users\togi\Documents\ST.docx";

    File.Delete(destinationPath);
    File.Copy(originalPath, destinationPath);
}