﻿using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Label_tool.Views
{
    public class ExcelLabelViewHelper
    {
        private double _h3High = 0.0;
        private double _c3High = 0.0;

        public ExcelLabelViewHelper() 
        {
            //Constructor
        }
        public void setCell3High(double[] arrCell3High)
        {
            this._h3High = arrCell3High[0];
            this._c3High = arrCell3High[1];
        }
        public void setBottomLine(Excel.Range cell)
        {
            cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].ColorIndex = Excel.XlRgbColor.rgbBlack; // Set border color to black
            cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].TintAndShade = 0;
            cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin; // Set border weight
        }

        public void SetBorders(Excel.Range cell)
        {
            cell.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell.Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = Excel.XlRgbColor.rgbBlack;
            cell.Borders[Excel.XlBordersIndex.xlEdgeLeft].TintAndShade = 0;
            cell.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;

            cell.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell.Borders[Excel.XlBordersIndex.xlEdgeRight].Color = Excel.XlRgbColor.rgbBlack;
            cell.Borders[Excel.XlBordersIndex.xlEdgeRight].TintAndShade = 0;
            cell.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

            cell.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell.Borders[Excel.XlBordersIndex.xlEdgeTop].Color = Excel.XlRgbColor.rgbBlack;
            cell.Borders[Excel.XlBordersIndex.xlEdgeTop].TintAndShade = 0;
            cell.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;

            cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = Excel.XlRgbColor.rgbBlack;
            cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].TintAndShade = 0;
            cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
        }

        public void setH1(Excel.Range cell)
        {
            cell.RowHeight = 24.9;
            cell.Font.Size = 13;
            cell.Font.Bold = true;
            cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }

        public void setH2(Excel.Range cell, string txt)
        {
            cell.Value = txt;
            cell.RowHeight = 18;
            cell.Font.Size = 12;
            cell.Font.Bold = true;
            cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
        }
        public void setC2(Excel.Range cell, string txt)
        {
            cell.Value = txt;
            cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            setBottomLine(cell);
        }

        public void setH3(Excel.Range cell, string txt)
        {
            cell.Value = txt;
            cell.RowHeight = this._h3High;
            cell.WrapText = true;
            cell.Font.Bold = true;
            cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cell.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            SetBorders(cell);
        }

        public void setC3(Excel.Range cell, string txt)
        {
            cell.Value = txt;
            cell.RowHeight = this._c3High;
            cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            SetBorders(cell);
        }

        public void setInterval(Excel.Range cell)
        {
            cell.RowHeight = 6.0;
        }
    }
}
