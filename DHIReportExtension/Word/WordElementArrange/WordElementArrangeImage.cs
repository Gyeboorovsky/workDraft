using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NonVisualGraphicFrameDrawingProperties = DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties;
using NonVisualDrawingProperties = DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties;
using NonVisualPictureProperties = DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties;
using NonVisualPictureDrawingProperties = DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties;
using ShapeProperties = DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties;
using BlipFill = DocumentFormat.OpenXml.Drawing.Pictures.BlipFill;
using Picture = DocumentFormat.OpenXml.Drawing.Pictures.Picture;

namespace DHIReportExtension
{
    public class WordElementArrangeImage
    {
        public static Drawing Arrange_Drawing(WordArrangeSettings settings)
        {
            var drawing = new Drawing();
            
            drawing.Append(Arrange_Inline(settings));

            return drawing;
        }

        public static Inline Arrange_Inline(WordArrangeSettings settings)
        {
            var inline = new Inline();

            inline.AppendChild(Arrange_Extent(settings));
            inline.AppendChild(Arrange_EffectExtent(settings));
            inline.AppendChild(Arrange_DocProperties(settings));
            inline.AppendChild(Arrange_NonVisualGraphicFrameDrawingProperties());
            inline.AppendChild(Arrange_Graphic(settings));

            return inline;
        }

        public static EffectExtent Arrange_EffectExtent(WordArrangeSettings settings)
        {
            var effectExtent = new EffectExtent();

            effectExtent.BottomEdge = settings.EffectExtent.BottomEdge;
            effectExtent.RightEdge = settings.EffectExtent.RightEdge;
            effectExtent.LeftEdge = settings.EffectExtent.LeftEdge;
            effectExtent.TopEdge = settings.EffectExtent.TopEdge;
            
            return effectExtent;
        }
        
        public static NonVisualGraphicFrameDrawingProperties Arrange_NonVisualGraphicFrameDrawingProperties()
        {
            var nonVisualGraphicFrameDrawingProperties = new NonVisualGraphicFrameDrawingProperties();

            nonVisualGraphicFrameDrawingProperties.AppendChild(Arrange_GraphicFrameLocks());

            return nonVisualGraphicFrameDrawingProperties;
        }

        public static GraphicFrameLocks Arrange_GraphicFrameLocks()
        {
            var graphicFrameLocks = new GraphicFrameLocks();

            return graphicFrameLocks;
        }
        
        public static Graphic Arrange_Graphic(WordArrangeSettings settings)
        {
            var graphic = new Graphic();

            graphic.AppendChild(Arrange_GraphicData(settings));

            return graphic;
        }
        
        public static GraphicData Arrange_GraphicData(WordArrangeSettings settings)
        {
            var graphicData = new GraphicData();

            graphicData.Uri = settings.GraphicData.Uri;

            graphicData.AppendChild(Arrange_Picture(settings));

            return graphicData;
        }

        public static Picture Arrange_Picture(WordArrangeSettings settings)
        {
            var picture = new Picture();

            picture.AppendChild(Arrange_NonVisualPictureProperties(settings));
            picture.AppendChild(Arrange_BlipFill(settings));
            picture.AppendChild(Arrange_ShapeProperties(settings));

            return picture;
        }

        public static BlipExtensionList Arrange_BlipExtensionList(WordArrangeSettings settings)
        {
            var blipExtensionList = new BlipExtensionList();
            
            blipExtensionList.AppendChild(Arrange_BlipExtension(settings));

            return blipExtensionList;
        }
        
        public static BlipExtension Arrange_BlipExtension(WordArrangeSettings settings)
        {
            var blipExtension = new BlipExtension
            {
                Uri = settings.BlipExtension.Uri
            };

            return blipExtension;
        }

        public static NonVisualPictureProperties Arrange_NonVisualPictureProperties(WordArrangeSettings settings)
        {
            var nonVisualPictureProperties = new NonVisualPictureProperties();

            nonVisualPictureProperties.AppendChild(Arrange_NonVisualDrawingProperties(settings));
            nonVisualPictureProperties.AppendChild(Arrange_NonVisualPictureDrawingProperties());

            return nonVisualPictureProperties;
        }

        public static NonVisualPictureDrawingProperties Arrange_NonVisualPictureDrawingProperties()
        {
            var nonVisualPictureDrawingProperties = new NonVisualPictureDrawingProperties();

            return nonVisualPictureDrawingProperties;
        }
        
        public static NonVisualDrawingProperties Arrange_NonVisualDrawingProperties(WordArrangeSettings settings)
        {
            var nonVisualDrawingProperties = new NonVisualDrawingProperties();
            
            //TODO unhardcode this name as in docProperties
            nonVisualDrawingProperties.Name = settings.NonVisualDrawingProperties.Name;
            nonVisualDrawingProperties.Id = settings.NonVisualDrawingProperties.Id;

            return nonVisualDrawingProperties;
        }

        public static BlipFill Arrange_BlipFill(WordArrangeSettings settings)
        {
            var blipFill = new BlipFill();

            blipFill.AppendChild(Arrange_Blip(settings));
            blipFill.AppendChild(Arrange_Stretch());

            return blipFill;
        }
        
        public static Blip Arrange_Blip(WordArrangeSettings settings)
        {
            var blip = new Blip();

            blip.AppendChild(Arrange_BlipExtensionList(settings));
            
            blip.Embed = settings.Blip.Embed;
            blip.CompressionState = settings.Blip.CompressionState;
            
            return blip;
        }

        public static IdPartPair Arrange_IdPartPair(WordArrangeSettings settings)
        {
            var idPartPair = new IdPartPair(settings.IdPartPair.Id, Arrange_ImagePart(settings));

            return idPartPair;
        }
        
        public static ImagePart Arrange_ImagePart(WordArrangeSettings settings)
        {
            ImagePart imagePart = new MyImagePart()
            {
                ContentType = settings.ImagePart.ContentType,
                Uri = settings.ImagePart.Uri
            };

            return imagePart;
        }

        public static Stretch Arrange_Stretch()
        {
            var stretch = new Stretch();

            stretch.AppendChild(Arrange_FillRectangle());
            
            return stretch;
        }
        
        public static FillRectangle Arrange_FillRectangle()
        {
            var fillRectange = new FillRectangle();

            return fillRectange;
        }

        public static ShapeProperties Arrange_ShapeProperties(WordArrangeSettings settings)
        {
            var shapeProperties = new ShapeProperties();

            shapeProperties.AppendChild(Arrange_Transform2D(settings));
            shapeProperties.AppendChild(Arrange_PresetGeometry(settings));

            return shapeProperties;
        }
        
        public static Transform2D Arrange_Transform2D(WordArrangeSettings settings)
        {
            var transform2d = new Transform2D();

            transform2d.AppendChild(Arrange_Offset(settings));
            transform2d.AppendChild(Arrange_Extents(settings));

            return transform2d;
        }

        private static Extents Arrange_Extents(WordArrangeSettings settings)
        {
            var extents = new Extents();

            extents.Cy = settings.Extents.Cy;
            extents.Cx = settings.Extents.Cx;

            return extents;
        }

        public static Extent Arrange_Extent(WordArrangeSettings settings)
        {
            var extent = new Extent();

            extent.Cy = settings.Extent.Cy;
            extent.Cx = settings.Extent.Cx;

            return extent;
        }
        
        public static Offset Arrange_Offset(WordArrangeSettings settings)
        {
            var offset = new Offset();

            offset.Y = settings.Offset.Y;
            offset.X = settings.Offset.X;
            
            return offset;
        }
        
        public static PresetGeometry Arrange_PresetGeometry(WordArrangeSettings settings)
        {
            var presetGeometry = new PresetGeometry();

            presetGeometry.Preset = settings.PresetGeometry.Preset;

            presetGeometry.AppendChild(Arrange_AdjustValueList());

            return presetGeometry;
        }

        public static AdjustValueList Arrange_AdjustValueList()
        {
            var adjustValueList = new AdjustValueList();

            return adjustValueList;
        }

        public static DocProperties Arrange_DocProperties(WordArrangeSettings settings)
        {
            var docProperties = new DocProperties();

            docProperties.Name = settings.DocProperties.Name;
            docProperties.Id = settings.DocProperties.Id;
            
            return docProperties;
        }
    }
}