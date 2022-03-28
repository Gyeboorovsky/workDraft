using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
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

        private static void InsertImageObject(this Body body)
        {
            var settings = new WordArrangeSettings();
            var image = WordElementArrangeGeneral.Arrange_Paragraph(ArrangeElementType.Image, settings);

            var validator = new OpenXmlValidator();
            var objectValidationErrors = validator.Validate(image);
            
            body.AppendChild(image);
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
            //// Define the reference of the image.
            
            // var element =
            //     new Drawing(
            //         new DW.Inline(
            //             new DW.Extent() {Cx = 990000L, Cy = 792000L},
            //             new DW.EffectExtent()
            //             {
            //                 LeftEdge = 0L, TopEdge = 0L,
            //                 RightEdge = 0L, BottomEdge = 0L
            //             },
            //             new DW.DocProperties()
            //             {
            //                 Id = (UInt32Value) 1U,
            //                 Name = "Picture 1"
            //             },
            //             new DW.NonVisualGraphicFrameDrawingProperties(
            //                 new A.GraphicFrameLocks() {NoChangeAspect = true}),
            //             new A.Graphic(
            //                 new A.GraphicData(
            //                     new PIC.Picture(
            //                         new PIC.NonVisualPictureProperties(
            //                             new PIC.NonVisualDrawingProperties()
            //                             {
            //                                 Id = (UInt32Value) 0U,
            //                                 Name = "New Bitmap Image.jpg"
            //                             },
            //                             new PIC.NonVisualPictureDrawingProperties()),
            //                         new PIC.BlipFill(
            //                             new A.Blip(
            //                                 new A.BlipExtensionList(
            //                                     new A.BlipExtension()
            //                                     {
            //                                         Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
            //                                     })
            //                             )
            //                             {
            //                                 Embed = relationshipId,
            //                                 CompressionState =
            //                                     A.BlipCompressionValues.Print
            //                             },
            //                             new A.Stretch(
            //                                 new A.FillRectangle())),
            //                         new PIC.ShapeProperties(
            //                             new A.Transform2D(
            //                                 new A.Offset() {X = 0L, Y = 0L},
            //                                 new A.Extents() {Cx = 990000L, Cy = 792000L}),
            //                             new A.PresetGeometry(
            //                                 new A.AdjustValueList()
            //                             ) {Preset = A.ShapeTypeValues.Rectangle}))
            //                 ) {Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"})
            //         )
            //         {
            //             DistanceFromTop = (UInt32Value) 0U,
            //             DistanceFromBottom = (UInt32Value) 0U,
            //             DistanceFromLeft = (UInt32Value) 0U,
            //             DistanceFromRight = (UInt32Value) 0U, EditId = "50D07946"
            //         });

            //// Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }

        private static void ProcessBody(this List<OpenXmlElement> placeholders, List<string> tags, string tagLocalName = "sdt")
        {
            //TODO substituting via interpreter
            ////////////////////////////////////////////////////////////////
            var tempTextToSubstitute = "Mój przyjaciel prosiaczek";

            foreach (var t in placeholders)
                if (t.LocalName == tagLocalName && tags.Contains(t.InnerText))
                {
                    t.SubstituteText(tempTextToSubstitute);
                }
        }
    }
}

