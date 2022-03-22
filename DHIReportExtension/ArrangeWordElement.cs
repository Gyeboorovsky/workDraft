using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DHIReportExtension
{
    public static class ArrangeWordElement
    {
        public static Table Arrange_Table(int cols, int rows)
        {
            var tbl = new Table();

            tbl.Append(Arrange_TableProperties());
            tbl.Append(Arrange_TableGrid(cols));

            for (int i = 0; i < rows; i++)
                tbl.Append(Arrange_TableRow(cols));

            return tbl;
        }

        public static TableProperties Arrange_TableProperties()
        {
            var tblProp = new TableProperties();
            
            tblProp.AppendChild(Arrange_TableStyle());
            tblProp.AppendChild(Arrange_TableWidth());
            tblProp.AppendChild(Arrange_TableLook());

            return tblProp;
        }

        public static TableStyle Arrange_TableStyle()
        {
            return new TableStyle()
            {
                Val = "TableGrid"
            };
        }

        public static TableWidth Arrange_TableWidth()
        {
            return new TableWidth();
        }

        public static TableLook Arrange_TableLook()
        {
            return new TableLook()
            {
                Val = "04A0"
            };
        }
        
        public static Paragraph Arrange_Paragraph()
        {
            var paragraph = new Paragraph();
            
            paragraph.Append(Arrange_ParagraphProperties());
            paragraph.Append(Arrange_Run());

            return paragraph;
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

        public static TableRow Arrange_TableRow(int cols)
        {
            var tblRow = new TableRow();
            
            for (int i = 0; i < cols; i++)
            {
                tblRow.Append(Arrange_TableCell());
            }

            return tblRow;
        }
        
        public static TableCell Arrange_TableCell()
        {
            var tblCell = new TableCell();
            
            tblCell.Append(Arrange_TableCellProperties());
            tblCell.Append(Arrange_Paragraph());
            
            return tblCell;
        }

        public static TableBorders Arrange_TableBorders()
        {
            var border = new TableBorders();

            border.AppendChild(new TopBorder()
            {
                Val = new EnumValue<BorderValues>(BorderValues.Thick),
                Color = "CC0000"
            });

            border.AppendChild(new RightBorder()
            {
                Val = new EnumValue<BorderValues>(BorderValues.Thick),
                Color = "CC0000"
            });

            // border.AppendChild(new BottomBorder()
            // {
            //     Val = new EnumValue<BorderValues>(BorderValues.Thick),
            //     Color = "CC0000"
            // });

            border.AppendChild(new LeftBorder()
            {
                Val = new EnumValue<BorderValues>(BorderValues.Thick),
                Color = "CC0000"
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

        public static Run Arrange_Run()
        {
            var run = new Run();
            
            run.Append(Arrange_RunProperties());
            run.Append(Arrange_Text());
            
            return run;
        }
    }
}