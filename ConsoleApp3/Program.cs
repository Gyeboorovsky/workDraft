using DHIReportExtension;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Office.Interop.Word;
using WI = DHIReportExtension.WordElementArrangeImage;

string documentCopyPath = @"C:\Users\togi\Documents\ST - Copy.docx";
//string imagePath = @"C:\Users\togi\Documents\image3.png";

CleanTemplate_DevelopmentTool(documentCopyPath);

using WordprocessingDocument document = WordprocessingDocument.Open(documentCopyPath, true);

var chartPath = NPlotLib.Plot(800, 400);
ReportExtension.InsertAPicture(document, chartPath);

document.GenerateReport();


//////////////////////////////////////////

static void CleanTemplate_DevelopmentTool(string destinationPath)
{
    string originalPath = @"C:\Users\togi\Documents\ST.docx";

    File.Delete(destinationPath);
    File.Copy(originalPath, destinationPath);
}

// static void SetImageReference(WordprocessingDocument document)
// {
//     document.MainDocumentPart.Parts.Append(WI.Arrange_IdPartPair());
// }

// static void AddImageResources()
// {
//     string copyPath = @"C:\Users\togi\Documents\ST - Copy.docx";
//     string extractPath = @"C:\Users\togi\Documents\folder";
//
//     if (Directory.Exists(extractPath))
//         Directory.Delete(extractPath, true);
//
//     ExtractDocx(copyPath, extractPath);
//     CopyImageToDocxFiles();
//     CompressToDocx(extractPath);
//     
//     Directory.Delete(extractPath, true);
// }

// static void ExtractDocx(string zipPath, string extractPath)
// {
//     if(Directory.Exists(extractPath))
//         Directory.Delete(extractPath);
//     
//     System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, extractPath);
// }
//
// static void CompressToDocx(string startPath)
// {
//     File.Delete(@"C:\Users\togi\Documents\ST - Copy.docx");
//     string zipPath = @"C:\Users\togi\Documents\ST - Copy.docx";
//     System.IO.Compression.ZipFile.CreateFromDirectory(startPath, zipPath);
// }
//
// static void CopyImageToDocxFiles()
// {
//     const string from = @"C:\Users\togi\Documents\image3.png";
//     const string to = @"C:\Users\togi\Documents\folder\word\media\image3.png";
//
//     if (!Directory.Exists(@"C:\Users\togi\Documents\folder\word\media"))
//         Directory.CreateDirectory(@"C:\Users\togi\Documents\folder\word\media");
//     
//     File.Copy(from, to);
// }