using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Words.Fields;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Vml.Wordprocessing;
using TopBorder = DocumentFormat.OpenXml.Drawing.TopBorder;

namespace DHIReportExtension
{
    public class WordArrangeSettings
    {
        
        public WordArrangeSettings(string imageRelationshipId)
        {
            Blip = new BlipSettings(imageRelationshipId);
        }
        
        public WordArrangeSettings()
        {
            Blip = new BlipSettings();
        }
        
        public BlipSettings Blip;
        public InlineSettings Inline = new ();
        public GraphicDataSettings GraphicData = new ();
        public PresetGeometrySettings PresetGeometry = new ();
        public ExtentsSettings Extents = new ();
        public OffsetSettings Offset = new ();
        public BlipExtensionSettings BlipExtension = new ();
        public NonVisualDrawingPropertiesSettings NonVisualDrawingProperties = new ();
        public GraphicFrameLockSettings GraphicFrameLock = new ();
        public DocPropertiesSettings DocProperties = new ();
        public EffectExtentSettings EffectExtent = new ();
        public ExtentsSettings Extent = new ();
        public IdPartPairSettings IdPartPair = new ();
        public ImagePartSettings ImagePart = new ();

        public TableStyleSettings TableStyle = new ();
        public TableLockSettings TableLock = new ();
        public TableBordersSettings TableBorders = new ();
    }

    public class TableBordersSettings
    {
        public TopBorderSettings TopBorder = new ();
        public RightBorderSettings RightBorder = new ();
        public BottomBorderSettings BottomBorder = new ();
        public LeftBorderSettings LeftBorder = new ();
    }

    public class TopBorderSettings
    {
        public BorderValues Val = BorderValues.Thick;
        public string Color = "CC0000";
    }

    public class RightBorderSettings
    {
        public BorderValues Val = BorderValues.Thick;
        public string Color = "CC0000";
    }

    public class BottomBorderSettings
    {
        public BorderValues Val = BorderValues.Thick;
        public string Color = "CC0000";
    }

    public class LeftBorderSettings
    {
        public BorderValues Val = BorderValues.Thick;
        public string Color = "CC0000";
    }
    
    public class TableLockSettings
    {
        public string Val = "04A0";
    }

    public class TableStyleSettings
    {
        public string Val = "TableGrid";
    }

    public class ImagePartSettings
    {
        public string ContentType = "image/png";
        public static string UriString = "/word/media/image3.png";
        public Uri Uri = new Uri(UriString, UriKind.Relative);
    }

    public class IdPartPairSettings
    {
        public string Id = null;
    }

    public class ExtentSettings
    {
        public Int64Value Cx = 990000L;
        public Int64Value Cy = 792000L;
    }

    public class EffectExtentSettings
    {
        public Int64Value LeftEdge = 0L;
        public Int64Value TopEdge = 0L;
        public Int64Value RightEdge = 0L;
        public Int64Value BottomEdge = 0L;
    }

    public class DocPropertiesSettings
    {
        public UInt32Value Id = 1U;
        public string Name = "Picture 1";
    }

    public class GraphicFrameLockSettings
    {
        public bool NoChangeAspect = true;
    }

    public class NonVisualDrawingPropertiesSettings
    {
        public UInt32Value Id = 0U;
        public string Name = "New Bitmap Image.jpg";
    }
    
    public class BlipExtensionSettings
    {
        public string Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}";
    }

    public class BlipSettings
    {
        public string Embed;
        
        public BlipSettings(string relationshipId)
        {
            Embed = relationshipId;
        }
        
        public BlipSettings()
        {
        }

        public BlipCompressionValues CompressionState = BlipCompressionValues.Print;
    }

    public class OffsetSettings
    {
        public Int64Value? X = 0L;
        public Int64Value? Y = 0L;
    }

    public class ExtentsSettings
    {
        public Int64Value? Cx = 990000L;
        public Int64Value? Cy = 792000L;
    }

    public class InlineSettings
    {
        public UInt32Value DistanceFromTop = 0U;
        public UInt32Value DistanceFromBottom = 0U;
        public UInt32Value DistanceFromLeft = 0U;
        public UInt32Value DistanceFromRight = 0U;
        public string EditId = "50D07946";
    }

    public class GraphicDataSettings
    {
        public string Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    }
    
    public class PresetGeometrySettings
    {
        public ShapeTypeValues Preset = ShapeTypeValues.Rectangle;
    }
}