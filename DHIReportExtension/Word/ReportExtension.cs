using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WT = DHIReportExtension.WordElementArrangeTable;

namespace DHIReportExtension
{
    public static class ReportExtension
    {
        public static void GenerateReport(this WordprocessingDocument wdDoc)
        {
            Body rawBody = wdDoc.MainDocumentPart.Document.Body;
            
            var tags = new List<string>() {"Rower"};
            var tagName = "holder";
            
            var placeHolders = rawBody.GetPlaceHolders(tagName);
            
            placeHolders.ProcessBody(tags);
            
            wdDoc.Save();
        }

        private static void InsertTable(this Body rawBody, int columns, int rows)
        {
            var settings = new WordArrangeSettings();
            var tbl = WT.Arrange_Table(columns, rows, settings);
            
            var validator = new OpenXmlValidator();
            var objectValidationErrors = validator.Validate(tbl);
            
            tbl.FillTable();
            rawBody.Append(tbl);
        }
        
        private static void SubstituteText(this OpenXmlElement element, string textToSubstitute)
        {
            var originalItem = (Text) element.FindFirstOfDefaultOf<Text>()!;

            originalItem.Text = textToSubstitute;
        }
        
        private static void FillTable(this Table inputTable)
        {
            
            //TODO interpreter response instead of textToSubstitute
            var textToSubstitute = "1,ser,chedar,2,mleko,3.2,3,chleb,pszenny";
            //////////////////////////////////////////////////////////////////
            
            var tabRow = inputTable.FindFirstOfDefaultOf<TableRow>();
            var Rows = inputTable.FindAllOf<TableRow>();
            int numOfCols = tabRow.ChildElements.Count;
            
            
            var listOfWords = textToSubstitute.Split(",").ToList();
            int numOfRows = listOfWords.Count / numOfCols;
            string[,] tableArray = new string[numOfRows, numOfCols];

            for (int i = 0; i < listOfWords.Count; i++)
            {
                tableArray[i / numOfCols, i % numOfCols] = listOfWords[i];
            }

            var settings = new WordArrangeSettings();
            var table = WT.Arrange_Table(numOfCols, numOfRows, settings);
            var generatedTable = WT.Arrange_Table(numOfCols, numOfRows, settings);

        }

        public static void InsertAPicture(WordprocessingDocument wordprocessingDocument, string fileName)
        {
            MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            var settings = new WordArrangeSettings(relationshipId);
            var element = WordElementArrangeImage.Arrange_Drawing(settings);
            
            //// Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }

        private static void ProcessBody(this List<OpenXmlElement> placeholders, List<string> tags, string tagLocalName = "sdt")
        {
            //TODO substituting via interpreter
            var tempTextToSubstitute = "Mój przyjaciel prosiaczek";

            foreach (var t in placeholders)
                if (t.LocalName == tagLocalName && tags.Contains(t.InnerText))
                {
                    t.SubstituteText(tempTextToSubstitute);
                }
        }
    }
}

