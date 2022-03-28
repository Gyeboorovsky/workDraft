using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Highlight = DocumentFormat.OpenXml.Wordprocessing.Highlight;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Underline = DocumentFormat.OpenXml.Wordprocessing.Underline;
using WI = DHIReportExtension.WordElementArrangeImage;

namespace DHIReportExtension
{
    public static class WordElementArrangeGeneral
    {
        public static Paragraph Arrange_Paragraph(ArrangeElementType elementType, WordArrangeSettings settings)
        {
            var paragraph = new Paragraph();
            
            paragraph.Append(Arrange_ParagraphProperties());
            paragraph.Append(Arrange_Run(elementType, settings));

            return paragraph;
        }
        
        public static RunProperties Arrange_RunProperties()
        {
            var runProps = new RunProperties();
            runProps.Append(new Bold());
            runProps.Append(new BoldComplexScript());
            runProps.Append(new Italic());
            runProps.Append(new ItalicComplexScript());
            runProps.Append(new Color(){Val = new HexBinaryValue("000000")});
            runProps.Append(new FontSize(){Val = new StringValue("44")});
            runProps.Append(new FontSizeComplexScript(){Val = new StringValue("44")});
            runProps.Append(new Highlight(){Val = new EnumValue<HighlightColorValues>(HighlightColorValues.None)});
            runProps.Append(new Underline(){Val = new EnumValue<UnderlineValues>(UnderlineValues.None)});

            return runProps;
        }

        public static ParagraphProperties Arrange_ParagraphProperties()
        {
            var paragraphProp = new ParagraphProperties();
            paragraphProp.Append(Arrange_ParagraphMarkRunProperties());
            
            return paragraphProp;
        }

        public static ParagraphMarkRunProperties Arrange_ParagraphMarkRunProperties()
        {
            var pmrp = new ParagraphMarkRunProperties();
            pmrp.Append(new Bold());
            pmrp.Append(new BoldComplexScript());
            pmrp.Append(new Italic());
            pmrp.Append(new ItalicComplexScript());
            pmrp.Append(new Color(){Val = new HexBinaryValue("000000")});
            pmrp.Append(new FontSize(){Val = new StringValue("44")});
            pmrp.Append(new FontSizeComplexScript(){Val = new StringValue("44")});
            pmrp.Append(new Highlight(){Val = new EnumValue<HighlightColorValues>(HighlightColorValues.None)});
            pmrp.Append(new Underline(){Val = new EnumValue<UnderlineValues>(UnderlineValues.None)});

            return pmrp;
        }

        public static Text Arrange_Text()
        {
            return new Text();
        }

        public static Run Arrange_Run(ArrangeElementType elementType, WordArrangeSettings settings)
        {
            var run = new Run();
            
            run.Append(Arrange_RunProperties());

            switch (elementType)
            {
                case ArrangeElementType.Table:
                    run.Append(Arrange_Text());
                    break;
                case ArrangeElementType.Image:
                    run.Append(WI.Arrange_Drawing(settings));
                    break;
            }
            
            return run;
        }
    }

    public class MyImagePart : ImagePart
    {
        public string ContentType { get; set; }
        public Uri Uri { get; set; }
    }
}