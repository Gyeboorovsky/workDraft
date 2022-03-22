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
            //Created object validator
            var wynikValidacji = validator.Validate(tbl);
            //rawBody.InsertAt(tab, rawBody.Count()-1);
            rawBody.Append(tbl);
        }

        private static List<OpenXmlElement> GetPlaceHolders(this OpenXmlElement elem, string tagName)
        {
            var list = new List<OpenXmlElement>();
            
            if (elem.HasTagName(tagName))
            {
                list.Add(elem);
                return list;
            }

            foreach (var childElement in elem.ChildElements)
            {
                list.AddRange(childElement.GetPlaceHolders(tagName));
            }

            return list;
        }

        private static bool HasTagName(this OpenXmlElement element, string tagName)
        {
            if (element.GetType() != typeof(SdtBlock))
                return false;
            var tag = (Tag)element.ChildElements.FirstOrDefault(x => x.GetType() == typeof(SdtProperties))?.ChildElements
                .FirstOrDefault(x => x.GetType() == typeof(Tag))!;
            return tag?.Val?.Value == tagName;
        }
        
        private static void FillWithChildren(this Body body, List<OpenXmlElement>? tree)
        {
            foreach (var element in tree)
            {
                body.AppendChild(element);
            }
        }

        private static OpenXmlElement? FindFirstOfDefaultOf<T>(this OpenXmlElement elem)
        where T : OpenXmlElement
        {
            foreach (var childElement in elem.ChildElements)
            {
                if (childElement.GetType() == typeof(T))
                    return childElement;
                
                var result = childElement.FindFirstOfDefaultOf<T>();
                if (result?.GetType() == typeof(T))
                    return result;
            }
            
            return null;
        }
        
        private static void Substitute(this OpenXmlElement element, string textToSubstitute)
        {
            //TODO interpreter in this method?
            var originalItem = (Text) element.FindFirstOfDefaultOf<Text>()!;
            
            //TODO interpreter response instead of textToSubstitute
            
            originalItem.Text = textToSubstitute;
        }
        
        //todo substitute with table
        private static void SubstituteWithTable(this OpenXmlElement element, List<OpenXmlElement> list)
        {
            
            //TODO interpreter response instead of textToSubstitute
            ///////////////////////////////////////////////////////////
        
            var textToSubstitute = "1,ser,chedar,2,mleko,3.2,3,chleb,pszenny";
            var listOfWords = textToSubstitute.Split(",").ToList();
            var tableList = new List<List<string>>();
            int numOfCols = 3;
            int numOfRows = listOfWords.Count / numOfCols;
            
            

            var table = ArrangeWordElement.Arrange_Table(numOfCols, numOfRows);



        }

        // private static int FindElementIndexByLocalName(this List<OpenXmlElement> elementList, string localName)
        // {
        //     for (int i = 0; i < elementList.Count; i++)
        //     {
        //         if (elementList[i].LocalName == localName)
        //             return i;
        //     }
        //
        //     throw new Exception($"Element do not contain child element {localName}");
        // }
        
        private static void ProcessBody(this List<OpenXmlElement> placeholders, List<string> tags, string tagLocalName = "sdt")
        {
            //TODO substituting via interpreter
            var tempTextToSubstitute = "Mój przyjaciel prosiaczek";

            foreach (var t in placeholders)
                if (t.LocalName == tagLocalName && tags.Contains(t.InnerText))
                {
                    //t.Substitute(tempTextToSubstitute);
                }
        }
    }
}

