using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using GridColumn = DocumentFormat.OpenXml.Wordprocessing.GridColumn;
using InsideHorizontalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder;
using InsideVerticalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder;
using LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder;
using RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;
using TableGrid = DocumentFormat.OpenXml.Wordprocessing.TableGrid;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TableStyle = DocumentFormat.OpenXml.Wordprocessing.TableStyle;
using TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder;
using WG = DHIReportExtension.WordElementArrangeGeneral;

namespace DHIReportExtension
{
    public class WordElementArrangeTable
    {
        public static Table Arrange_Table(int cols, int rows, WordArrangeSettings settings)
        {
            var tbl = new Table();

            tbl.Append(Arrange_TableProperties(settings));
            tbl.Append(Arrange_TableGrid(cols));

            for (int i = 0; i < rows; i++)
                tbl.Append(Arrange_TableRow(cols, settings));

            return tbl;
        }

        public static TableProperties Arrange_TableProperties(WordArrangeSettings settings)
        {
            var tblProp = new TableProperties();
            
            tblProp.AppendChild(Arrange_TableStyle(settings));
            tblProp.AppendChild(Arrange_TableWidth());
            tblProp.AppendChild(Arrange_TableLook(settings));

            return tblProp;
        }

        public static TableStyle Arrange_TableStyle(WordArrangeSettings settings)
        {
            return new TableStyle()
            {
                Val = settings.TableStyle.Val
            };
        }

        public static TableWidth Arrange_TableWidth()
        {
            return new TableWidth();
        }

        public static TableLook Arrange_TableLook(WordArrangeSettings settings)
        {
            return new TableLook()
            {
                Val = settings.TableLock.Val
            };
        }

        public static TableGrid Arrange_TableGrid(int cols)
        {
            var tblGrid = new TableGrid();

            for (int i = 0; i < cols; i++)
            {
                tblGrid.Append(Arrange_GridColumn());
            }
            
            return tblGrid;
        }

        public static GridColumn Arrange_GridColumn()
        {
            return new GridColumn();
        }

        public static TableRow Arrange_TableRow(int cols, WordArrangeSettings settings)
        {
            var tblRow = new TableRow();
            
            for (int i = 0; i < cols; i++)
            {
                tblRow.Append(Arrange_TableCell(settings));
            }

            return tblRow;
        }
        
        public static TableCell Arrange_TableCell(WordArrangeSettings settings)
        {
            var tblCell = new TableCell();
            
            tblCell.Append(Arrange_TableCellProperties());
            tblCell.Append(WG.Arrange_Paragraph(ArrangeElementType.Table, settings));
            
            return tblCell;
        }

        public static TableBorders Arrange_TableBorders(WordArrangeSettings settings)
        {
            var border = new TableBorders();

            border.AppendChild(new TopBorder()
            {
                Val = new EnumValue<BorderValues>((BorderValues)settings.TableBorders.TopBorder.Val),
                Color = settings.TableBorders.TopBorder.Color
            });

            border.AppendChild(new RightBorder()
            {
                Val = new EnumValue<BorderValues>((BorderValues)settings.TableBorders.RightBorder.Val),
                Color = settings.TableBorders.RightBorder.Color
            });

            border.AppendChild(new BottomBorder()
            {
                Val = new EnumValue<BorderValues>((BorderValues)settings.TableBorders.BottomBorder.Val),
                Color = settings.TableBorders.BottomBorder.Color
            });

            border.AppendChild(new LeftBorder()
            {
                Val = new EnumValue<BorderValues>((BorderValues)settings.TableBorders.LeftBorder.Val),
                Color = settings.TableBorders.LeftBorder.Color
            });

            border.AppendChild(new InsideHorizontalBorder()
            {
                Val = new EnumValue<BorderValues>(BorderValues.Thick),
                Color = "CC0000"
            });

            border.AppendChild(new InsideVerticalBorder()
            {
                Val = new EnumValue<BorderValues>(BorderValues.Thick),
                Color = "CC0000"
            });

            return border;
        }
        
        public static TableCellProperties Arrange_TableCellProperties()
        {
            var tblCellProp = new TableCellProperties();
            
            tblCellProp.Append(Arrange_TableCellWidth());

            return tblCellProp;
        }

        public static TableCellWidth Arrange_TableCellWidth()
        {
            return new TableCellWidth()
            {
                Type = TableWidthUnitValues.Dxa,
                Width = "2000"
            };
        }

    }
}