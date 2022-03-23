using CSOpenXmlCreateChartInWord;
using DHIReportExtension;
using DocumentFormat.OpenXml.Packaging;

string pathWithoutExtension = @"C:\Users\togi\Documents\ST - Copy";
string extractPath = @"C:\Users\togi\Documents\folder";
string copyPath = @"C:\Users\togi\Documents\ST - Copy.docx";

CleanTemplate_DevelopmentTool(copyPath);

AddImage();

using WordprocessingDocument document = WordprocessingDocument.Open(copyPath, true);

document.GenerateReport();



static void CleanTemplate_DevelopmentTool(string destinationPath)
{
    string originalPath = @"C:\Users\togi\Documents\ST.docx";

    File.Delete(destinationPath);
    File.Copy(originalPath, destinationPath);
}

static void AddImage()
{
    string copyPath = @"C:\Users\togi\Documents\ST - Copy.docx";
    string extractPath = @"C:\Users\togi\Documents\folder";
    
    ExtractDocx(copyPath, extractPath);
    //image does not coping
    CopyImageToDocxFiles();
    CompressToDocx(extractPath);
}

static void ExtractDocx(string zipPath, string extractPath)
{
    System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, extractPath);
}

static void CompressToDocx(string startPath)
{
    string zipPath = @"C:\Users\togi\Documents\takiTamZip.docx";
    System.IO.Compression.ZipFile.CreateFromDirectory(startPath, zipPath);
}

static void CopyImageToDocxFiles()
{
    const string from = @"C:\Users\togi\Documents\image3.png";
    const string to = @"C:\Users\togi\Documents\folder\word\media";

}

static void PróbyZChartami(string copyPath)
{
    var chartSubArea = new List<ChartSubArea>();
    chartSubArea.Add(new ChartSubArea());
    var wordChartCreator = new CreateWordChart(copyPath);
    wordChartCreator.CreateChart(chartSubArea);
    Console.WriteLine("rower");
}