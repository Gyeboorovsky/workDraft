using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace DHIReportExtension
{
    public static class ReportExtension
    {
        public static void GenerateReport(this WordprocessingDocument wdDoc)
        {
            Body rawBody = wdDoc.MainDocumentPart.Document.Body;

            //testing table
            rawBody.InsertTable(3, 2);
            
            var tags = new List<string>() {"Rower"};
            var tagName = "holder";
            
            var placeHolders = rawBody.GetPlaceHolders(tagName);
            
            placeHolders.ProcessBody(tags);
            
            wdDoc.Save();
        }

        private static void InsertTable(this Body rawBody, int columns, int rows)
        {
            var tbl = ArrangeWordElement.Arrange_Table(columns, rows);
            var validator = new OpenXmlValidator();
            
            var validationObject = validator.Validate(tbl);
            
            tbl.FillTable();
            rawBody.Append(tbl);
        }
        
        private static void Substitute(this OpenXmlElement element, string textToSubstitute)
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

            var table = ArrangeWordElement.Arrange_Table(numOfCols, numOfRows);

            var generatedTable = ArrangeWordElement.Arrange_Table(numOfCols, numOfRows);

        }
        
        private static void ProcessBody(this List<OpenXmlElement> placeholders, List<string> tags, string tagLocalName = "sdt")
        {
            //TODO substituting via interpreter
            ////////////////////////////////////////////////////////////////
            var tempTextToSubstitute = "Mój przyjaciel prosiaczek";

            foreach (var t in placeholders)
                if (t.LocalName == tagLocalName && tags.Contains(t.InnerText))
                {
                    t.Substitute(tempTextToSubstitute);
                }
        }
    }
}

