using Excel_tool.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Excel_tool.Views;
using System.Threading;

namespace Excel_tool.Controllers
{
    public class ExcelController
    {
        private ExcelModel _model;
        private Excel.Application _excelApp;
        private Excel.Workbook _workbook;
        private Excel.Worksheet _worksheet;
        private Excel.Worksheet _newWorksheet;
        private Excel.Range _range;
        private ExcelViewHelper _viewHelper = new ExcelViewHelper();

        public ExcelController(ExcelModel model)
        {
            _model = model;
        }
        public void setSignPic(PictureBox pictureBox)
        {
            _model.pictureBox = pictureBox;
        }
        public void setDate(string szDate)
        {
            _model.szDate = szDate;
        }
        public void setInspector(string szInspector)
        {
            _model.szInspector = szInspector;
        }

        public void setPiezoVer(string szPiezoV)
        {
            _model.szPiezoV = szPiezoV;
        }

        public void setDriverVer(string szDriverV)
        {
            _model.szDriverV = szDriverV;
        }

        public void BrowseFile()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (DialogResult.OK == dlg.ShowDialog())
            {
                _model.szFilePath = dlg.FileName;
            }
        }
        public bool IsFileLocked(string filePath)
        {
            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                return true;
            }
            return false;
        }
        public void OpenExcelFile()
        {
            if (IsFileLocked(_model.szFilePath))
            {
                MessageBox.Show("The file is currently open or locked", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                _excelApp = new Excel.Application();
                _workbook = _excelApp.Workbooks.Open(_model.szFilePath);
                _worksheet = (Excel.Worksheet)_workbook.Sheets[1];
                _range = _worksheet.UsedRange;

                int rows = _range.Rows.Count;
                int cols = _range.Columns.Count;

                _model.dataTable = new System.Data.DataTable();
                // Add columns to the DataTable
                for (int col = 1; col <= cols; col++)
                {
                    _model.dataTable.Columns.Add((col + 0x40).ToString(), typeof(string));
                }

                // Add rows to the DataTable
                for (int row = 1; row <= rows; row++)
                {
                    var newRow = _model.dataTable.NewRow();
                    for (int col = 1; col <= cols; col++)
                    {
                        var cellValue = (_range.Cells[row, col] as Excel.Range).Value2;
                        newRow[col - 1] = (null == cellValue) ? "" : cellValue;
                    }
                    _model.dataTable.Rows.Add(newRow);
                }

                // Find the worksheet or create a new one
                foreach (Excel.Worksheet sheet in _workbook.Worksheets)
                {
                    if ("Label_to_Print" == sheet.Name)
                    {
                        _newWorksheet = sheet;
                        _newWorksheet.Cells.Clear();
                        break;
                    }
                }

                if (null == _newWorksheet)
                {
                    _newWorksheet = _workbook.Worksheets.Add(After: _workbook.Worksheets[_workbook.Worksheets.Count]);
                    _newWorksheet.Name = "Label_to_Print";
                    _newWorksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                    _newWorksheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                    _newWorksheet.VPageBreaks.Add(_newWorksheet.Columns[14]);
                    _newWorksheet.HPageBreaks.Add(_newWorksheet.Rows[31]);

                    _newWorksheet.PageSetup.LeftMargin = _excelApp.InchesToPoints(0.64);
                    _newWorksheet.PageSetup.RightMargin = _excelApp.InchesToPoints(0.64);
                    _newWorksheet.PageSetup.TopMargin = _excelApp.InchesToPoints(1.91);
                    _newWorksheet.PageSetup.BottomMargin = _excelApp.InchesToPoints(1.91);
                    _newWorksheet.PageSetup.HeaderMargin = _excelApp.InchesToPoints(0.76);
                    _newWorksheet.PageSetup.FooterMargin = _excelApp.InchesToPoints(0.76);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to open Excel file. Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }
        public string GetFilePath()
        {
            return _model.szFilePath;
        }

        public string CheckValue()
        {
            string result = _model.dataTable.Rows[4 - 1][3 - 1].ToString();
            if ("Y22-088-USH05" == result)
            {
                return $"{result}: Supported";
            }
            else
            {
                return $"{result}: Not supported";
            }
        }

        public int GetRowCount()
        {
            return _model.dataTable.Rows.Count - 9 + 1;
        }

        public void CloseExcelFile()
        {
            // Dispose of dataTable (from the model)
            if (null != _model.dataTable)
            {
                _model.dataTable.Dispose();
                _model.dataTable = null;
            }

            // Release COM objects
            if (null != _range)
            {
                Marshal.ReleaseComObject(_range);
                _range = null;
            }
            if (null != _worksheet)
            {
                Marshal.ReleaseComObject(_worksheet);
                _worksheet = null;
            }
            if (null != _newWorksheet)
            {
                Marshal.ReleaseComObject(_newWorksheet);
                _newWorksheet = null;
            }
            if (null != _workbook)
            {
                _workbook.Close(false); // Close without saving
                Marshal.ReleaseComObject(_workbook);
                _workbook = null;
            }
            if (null != _excelApp)
            {
                Marshal.ReleaseComObject(_excelApp);
                _excelApp = null;
            }

            // Trigger garbage collection to free resources
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        
        public void GenerateReport(int startIdx, int endIdx)
        {
            try
            {
                for (int idx = startIdx; idx <= endIdx; idx++)
                {
                    int group = (idx - 1) / 2;
                    int rowStartIdx = 1 + (group * 15);
                    int colStartIdx = (0 == idx % 2) ? 8 : 1;
                    int dataIdx = 9 + (idx - 1);
                    string szLEDResult = "FAIL";

                    // Setting Excel columns, rows, headers
                    SetExcelLayout(_newWorksheet, rowStartIdx, colStartIdx, dataIdx, ref szLEDResult);

                    // Paste image
                    PasteImage(_newWorksheet, rowStartIdx, colStartIdx);

                }

                // Save the workbook
                _newWorksheet.SaveAs(_model.szFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"GenerateReport(), An unexpected error occurred: " + ex.Message);
            }
            finally
            {
                this.CloseExcelFile();
            }


        }

        private void PasteImage(Excel.Worksheet worksheet, int rowStartIdx, int colStartIdx)
        {
            Clipboard.SetImage(_model.pictureBox.Image);
            Excel.Range cell = worksheet.Cells[rowStartIdx + 12, colStartIdx + 5];
            ExcelViewHelper viewHelper = new ExcelViewHelper();
            viewHelper.setBottomLine(cell);

            // Paste the image in the cell
            Thread.Sleep(50);
            worksheet.Paste();
            Excel.Shapes shapes = worksheet.Shapes;
            Excel.Shape shape = shapes.Item(shapes.Count);
            shape.Top = (float)cell.Top - 3;
            shape.Left = (float)cell.Left + 5;
        }

        private void SetExcelLayout(Excel.Worksheet newWorksheet, int rowStartIdx, int colStartIdx, int dataIdx, ref string szLEDResult)
        {
            try
            {
                // start ========================================
                for (int colIdx = 0; 6 > colIdx; colIdx++)
                {
                    Excel.Range columnRange = newWorksheet.Columns[colStartIdx + colIdx];
                    columnRange.ColumnWidth = 10.5;
                }

                Excel.Range rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx, colStartIdx], newWorksheet.Cells[rowStartIdx, colStartIdx + 5]];
                rangeToMerge.Merge();
                rangeToMerge.Value = "US handpiece H05 test report";
                _viewHelper.setH1(rangeToMerge);
                _viewHelper.setH2(newWorksheet.Cells[rowStartIdx + 1, colStartIdx], "Type:");
                _viewHelper.setC2(newWorksheet.Cells[rowStartIdx + 1, colStartIdx + 1], "US H05");
                _viewHelper.setH2(newWorksheet.Cells[rowStartIdx + 1, colStartIdx + 4], "Date:");
                _viewHelper.setC2(newWorksheet.Cells[rowStartIdx + 1, colStartIdx + 5], _model.szDate);
                _viewHelper.setH2(newWorksheet.Cells[rowStartIdx + 2, colStartIdx], "SN:");
                _viewHelper.setC2(newWorksheet.Cells[rowStartIdx + 2, colStartIdx + 1], _model.dataTable.Rows[dataIdx - 1][1].ToString());
                _viewHelper.setH2(newWorksheet.Cells[rowStartIdx + 2, colStartIdx + 4], "Inspector:");
                _viewHelper.setC2(newWorksheet.Cells[rowStartIdx + 2, colStartIdx + 5], _model.szInspector);
                // interval ======================================
                _viewHelper.setInterval(newWorksheet.Cells[rowStartIdx + 3, colStartIdx]);
                // header ========================================
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx], newWorksheet.Cells[rowStartIdx + 5, colStartIdx]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "US cup surface");

                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 1], newWorksheet.Cells[rowStartIdx + 5, colStartIdx + 1]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Finishing housing");

                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 2], newWorksheet.Cells[rowStartIdx + 5, colStartIdx + 2]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Finishing gluing");

                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 3], newWorksheet.Cells[rowStartIdx + 5, colStartIdx + 3]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Batch    label");

                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 4], newWorksheet.Cells[rowStartIdx + 5, colStartIdx + 4]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Water proofness");

                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 5], newWorksheet.Cells[rowStartIdx + 5, colStartIdx + 5]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "LED indication");

                // content ===================================
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx], "OK");
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 1], "OK");
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 2], "OK");
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 3], "OK");
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 4], ("OK" == (string)_model.dataTable.Rows[dataIdx - 1][10]) ? "PASS" : "FAIL");
                if ("OK" == (string)_model.dataTable.Rows[dataIdx - 1][2] &&
                    "OK" == (string)_model.dataTable.Rows[dataIdx - 1][3] &&
                    "OK" == (string)_model.dataTable.Rows[dataIdx - 1][4] &&
                    "OK" == (string)_model.dataTable.Rows[dataIdx - 1][5])
                {
                    szLEDResult = "PASS";
                }
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 5], szLEDResult);
                // interval =================================
                _viewHelper.setInterval(newWorksheet.Cells[rowStartIdx + 7, colStartIdx]);
                // header ===================================
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 8, colStartIdx], newWorksheet.Cells[rowStartIdx + 9, colStartIdx]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Piezo   board");

                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 8, colStartIdx + 1], newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 1]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Driver   board");

                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 8, colStartIdx + 2], newWorksheet.Cells[rowStartIdx + 8, colStartIdx + 3]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Output power @2W/cm2");
                _viewHelper.setH3(newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 2], "1MHz");
                _viewHelper.setH3(newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 3], "3MHz");

                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 8, colStartIdx + 4], newWorksheet.Cells[rowStartIdx + 8, colStartIdx + 5]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Contact control");
                _viewHelper.setH3(newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 4], "1MHz");
                _viewHelper.setH3(newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 5], "3MHz");
                // content ================================
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx], _model.szPiezoV);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx + 1], _model.szDriverV);
                float fV = float.Parse(_model.dataTable.Rows[dataIdx - 1][8].ToString());
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx + 2], fV.ToString("F2"));
                fV = float.Parse(_model.dataTable.Rows[dataIdx - 1][9].ToString());
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx + 3], fV.ToString("F2"));
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx + 4], ("OK" == (string)_model.dataTable.Rows[dataIdx - 1][6]) ? "PASS" : "FAIL");
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx + 5], ("OK" == (string)_model.dataTable.Rows[dataIdx - 1][7]) ? "PASS" : "FAIL");
                // interval =============================
                _viewHelper.setInterval(newWorksheet.Cells[rowStartIdx + 11, colStartIdx]);
                // sign =================================
                newWorksheet.Cells[rowStartIdx + 12, colStartIdx + 3] = "Signature for approval:";
                Excel.Range cell = newWorksheet.Cells[rowStartIdx + 12, colStartIdx + 5];
                _viewHelper.setBottomLine(cell);
            }
            catch ( Exception ex)
            {
                MessageBox.Show($"idx:{dataIdx}, An unexpected error occurred: " + ex.Message);
            }

        }
    }
}
