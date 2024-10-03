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
using Excel_tool.Views;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices.ComTypes;

namespace Excel_tool.Controllers
{
    public class ExcelController
    {
        private IView _view;
        private ExcelModel _model;
        private Excel.Application _excelApp = null;
        private Excel.Workbook _workbook = null;
        private Excel.Worksheet _worksheet = null;
        private Excel.Worksheet _newWorksheet = null;
        private Excel.Range _range = null;
        private ExcelViewHelper _viewHelper = null;
        private KernelTbl[] arrKTbl = null;
        private uint kTblIdx = 0;
        private const string STR_PASS = "PASS";
        private const string STR_OK = "OK";

        public ExcelController(IView view)
        {
            this._model = new ExcelModel();
            this._view = view;
            this._viewHelper = new ExcelViewHelper();

            //register event 
            this._view.browseExcelFile += View_BrowseExcelFile;
            this._view.openExcelFile += View_OpenExcelFile;
            this._view.closeExcelFile += View_CloseExcelFile;
            this._view.browseImageFile += View_BrowseImageFile;
            this._view.generateReport += View_GenerateReport;

            //init kernel table
            arrKTbl = new KernelTbl[2];
            arrKTbl[0] = new KernelTbl("Y22-088-USH05", parseDataHP, generateLabelHP, new int[] { 9, 0 }, 12, new double[] {18.0, 25.0 });
            arrKTbl[1] = new KernelTbl("Y22-079-USC", parseDataUSC, generateLabelUSC, new int[] { 7, 6 }, 18, new double[] {15.2, 16.2 });
        }
        private bool isArtNoSupported(string szArtNo)
        {
            for (uint i = 0; arrKTbl.Length > i; i++)
            {
                if (arrKTbl[i].szArtNo == szArtNo)
                {
                    kTblIdx = i;
                    this._viewHelper.setCell3High(arrKTbl[kTblIdx].arrCellRowHigh);
                    return true;
                }
            }
            return false;
        }
        private bool isFileLocked(string filePath)
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
        private void GenerateReport(int startIdx, int endIdx)
        {
            try
            {
                arrKTbl[kTblIdx].generateLabel(startIdx, endIdx);

                // Save the workbook
                _newWorksheet.SaveAs(_model.szFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"GenerateReport(), An unexpected error occurred: " + ex.Message);
            }
            finally
            {
                this.View_CloseExcelFile(this, EventArgs.Empty);
            }
        }
        private void View_GenerateReport(object sender, EventArgs e)
        {
            this._model.szDate = this._view.getDate();
            this._model.szInspector = this._view.getInspector();
            this._model.pictureBox = this._view.getImageObj();
            this._model.szPiezoV = this._view.getPiezoVer();
            this._model.szDriverV = this._view.getDeiverVer();

            int startIdx = this._view.getStartIdx();
            int endIdx = this._view.getEndIdx();

            this.GenerateReport(startIdx, endIdx);
        }
        private void View_BrowseExcelFile(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (DialogResult.OK == dlg.ShowDialog())
            {
                this._model.szFilePath = dlg.FileName;
                this._view.setExcelPath(dlg.FileName);
            }
        }
        private void View_OpenExcelFile(object sender, EventArgs e)
        {
            if (this.isFileLocked(_model.szFilePath))
            {
                MessageBox.Show("The file is currently open or locked", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string szArtNo = "";
            try
            {
                _excelApp = new Excel.Application();
                _workbook = _excelApp.Workbooks.Open(_model.szFilePath);
                _worksheet = (Excel.Worksheet)_workbook.Sheets[1];
                _range = _worksheet.UsedRange;

                //Art No. position
                szArtNo = _worksheet.Cells[4, 3].Value?.ToString();
                if (false == this.isArtNoSupported(szArtNo))
                {
                    this._view.setOpenResult($"{szArtNo}: Not supported", 0);
                    return;
                }

                arrKTbl[kTblIdx].parseData();
            }
            catch (Exception ex)
            {
                this.View_CloseExcelFile(this, EventArgs.Empty);
                MessageBox.Show("Failed to open Excel file. Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this._view.setOpenResult($"{szArtNo}: Not supported", 0);
                return;
            }
            int rowCnt = _model.arrDataTable[0].Rows.Count - arrKTbl[kTblIdx].arrStartIdx[0] + 1;
            this._view.setOpenResult($"{szArtNo}: Supported", rowCnt);
            return;
        }
        private void View_CloseExcelFile(object sender, EventArgs e)
        {
            this._view.resetOpenResult();

            // Dispose of dataTable (from the model)
            if (null != _model.arrDataTable[0])
            {
                _model.arrDataTable[0].Dispose();
                _model.arrDataTable[0] = null;
            }
            if (null != _model.arrDataTable[1])
            {
                _model.arrDataTable[1].Dispose();
                _model.arrDataTable[1] = null;
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
        private void View_BrowseImageFile(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // Set filter options and filter index
                openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif";
                openFileDialog.FilterIndex = 1;

                // Set default file extension
                openFileDialog.DefaultExt = "jpg";

                // Show the dialog and get the result
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the selected file path
                    string szFilePath = openFileDialog.FileName;

                    // Load the image into the PictureBox
                    this._view.setImageFile(szFilePath);
                    //pictureBox1.Image = System.Drawing.Image.FromFile(filePath);
                }
            }
        }
        //=========================================================
        // kernel table start
        //=========================================================
        class KernelTbl
        {
            public string szArtNo;
            public int[] arrStartIdx = new int[2];
            public int picRowIdx;
            public double[] arrCellRowHigh = new double[2];
            public Func<int> parseData;
            public Func<int, int, int> generateLabel;
            public KernelTbl(string szArtNo, Func<int> parseData, Func<int, int, int> generateLabel, int[] arrStartIdx, int picRowIdx, double[] arrCellRowHigh)
            {
                this.szArtNo = szArtNo;
                this.parseData = parseData;
                this.generateLabel = generateLabel;
                this.arrStartIdx = arrStartIdx;
                this.picRowIdx = picRowIdx;
                this.arrCellRowHigh = arrCellRowHigh;
            }
        }
        private void PasteImage(Excel.Worksheet worksheet, int rowStartIdx, int colStartIdx)
        {
            //Clipboard needs to use STA: Single Thread Apartment
            Task.Run(() =>
            {
                var thread = new Thread(() => Clipboard.SetImage(_model.pictureBox.Image));
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
            });

            Excel.Range cell = worksheet.Cells[rowStartIdx + arrKTbl[kTblIdx].picRowIdx, colStartIdx + 5];
            this._viewHelper.setBottomLine(cell);
            
            // Paste the image in the cell
            Thread.Sleep(50);
            worksheet.Paste();
            Excel.Shapes shapes = worksheet.Shapes;
            Excel.Shape shape = shapes.Item(shapes.Count);
            shape.Top = (float)cell.Top - 2;
            shape.Left = (float)cell.Left + 9;
        }
        private int parseDataHP()
        {
            _worksheet = (Excel.Worksheet)_workbook.Sheets[1];
            _range = _worksheet.UsedRange;

            int rows = _range.Rows.Count;
            int cols = _range.Columns.Count;

            _model.arrDataTable[0] = new System.Data.DataTable();
            // Add columns to the DataTable
            for (int col = 1; col <= cols; col++)
            {
                _model.arrDataTable[0].Columns.Add((col + 0x40).ToString(), typeof(string));
            }

            // Add rows to the DataTable
            for (int row = 1; row <= rows; row++)
            {
                var newRow = _model.arrDataTable[0].NewRow();
                for (int col = 1; col <= cols; col++)
                {
                    var cellValue = (_range.Cells[row, col] as Excel.Range).Value2;
                    newRow[col - 1] = (null == cellValue) ? "" : cellValue;
                }
                _model.arrDataTable[0].Rows.Add(newRow);
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
                _newWorksheet.VPageBreaks.Add(_newWorksheet.Columns[13 + 1]);
                //_newWorksheet.HPageBreaks.Add(_newWorksheet.Rows[31]);

                _newWorksheet.PageSetup.LeftMargin = _excelApp.InchesToPoints(0.45);
                _newWorksheet.PageSetup.RightMargin = _excelApp.InchesToPoints(0.1);
                _newWorksheet.PageSetup.TopMargin = _excelApp.InchesToPoints(0.5);
                _newWorksheet.PageSetup.BottomMargin = _excelApp.InchesToPoints(0.3);
                _newWorksheet.PageSetup.HeaderMargin = _excelApp.InchesToPoints(0.0);
                _newWorksheet.PageSetup.FooterMargin = _excelApp.InchesToPoints(0.0);
            }
            return 0;
        }
        private int generateLabelHP(int startIdx, int endIdx)
        {
            for (int idx = startIdx; idx <= endIdx; idx++)
            {
                this._view.setCurrentlyIdx(idx);
                int group = (idx - 1) / 2;
                int rowStartIdx = 1 + (group * 18);
                int colStartIdx = (0 == idx % 2) ? 8 : 1;
                int dataIdx = arrKTbl[kTblIdx].arrStartIdx[0] + (idx - 1);

                // Setting Excel columns, rows, headers
                SetExcelLayoutHP(_newWorksheet, rowStartIdx, colStartIdx, dataIdx);

                // Paste image
                PasteImage(_newWorksheet, rowStartIdx, colStartIdx);

                // set page break
                if (0 == idx % 4)
                {
                    int rowIdx = rowStartIdx + 13 + 5;
                    _newWorksheet.HPageBreaks.Add(_newWorksheet.Rows[rowIdx]);
                }
            }
            return 0;
        }
        private void SetExcelLayoutHP(Excel.Worksheet newWorksheet, int rowStartIdx, int colStartIdx, int dataIdx)
        {
            try
            {
                // start ========================================
                for (int colIdx = 0; 6 > colIdx; colIdx++)
                {
                    Excel.Range columnRange = newWorksheet.Columns[colStartIdx + colIdx];
                    columnRange.ColumnWidth = 11.0;
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
                _viewHelper.setC2(newWorksheet.Cells[rowStartIdx + 2, colStartIdx + 1], _model.arrDataTable[0].Rows[dataIdx - 1][1].ToString());
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
                _viewHelper.setH3(rangeToMerge, "Batch      label");

                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 4], newWorksheet.Cells[rowStartIdx + 5, colStartIdx + 4]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Water proofness");

                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 5], newWorksheet.Cells[rowStartIdx + 5, colStartIdx + 5]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "LED indication");

                // content ===================================
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx], STR_OK);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 1], STR_OK);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 2], STR_OK);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 3], STR_OK);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 4], STR_PASS);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 5], STR_PASS);
                // interval =================================
                _viewHelper.setInterval(newWorksheet.Cells[rowStartIdx + 7, colStartIdx]);
                // header ===================================
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 8, colStartIdx], newWorksheet.Cells[rowStartIdx + 9, colStartIdx]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Piezo     board");

                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 8, colStartIdx + 1], newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 1]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Driver    board");

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
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx + 2], float.Parse((string)_model.arrDataTable[0].Rows[dataIdx - 1][8]).ToString("F2") + " W");
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx + 3], float.Parse((string)_model.arrDataTable[0].Rows[dataIdx - 1][9]).ToString("F2") + " W");
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx + 4], STR_PASS);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx + 5], STR_PASS);
                // interval =============================
                _viewHelper.setInterval(newWorksheet.Cells[rowStartIdx + 11, colStartIdx]);
                // sign =================================
                newWorksheet.Cells[rowStartIdx + 12, colStartIdx + 3] = "Signature for approval:";
                Excel.Range cell = newWorksheet.Cells[rowStartIdx + 12, colStartIdx + 5];
                _viewHelper.setBottomLine(cell);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"idx:{dataIdx}, An unexpected error occurred: " + ex.Message);
            }
        }
        private int parseDataUSC()
        {
            for (int i = 1; 2 >= i; i++)
            {
                _worksheet = (Excel.Worksheet)_workbook.Sheets[i];
                _range = _worksheet.UsedRange;

                int rows = _range.Rows.Count;
                int cols = _range.Columns.Count;

                _model.arrDataTable[i - 1] = new System.Data.DataTable();
                // Add columns to the DataTable
                for (int col = 1; col <= cols; col++)
                {
                    _model.arrDataTable[i - 1].Columns.Add((col + 0x40).ToString(), typeof(string));
                }

                // Add rows to the DataTable
                for (int row = 1; row <= rows; row++)
                {
                    var newRow = _model.arrDataTable[i - 1].NewRow();
                    for (int col = 1; col <= cols; col++)
                    {
                        var cellValue = (_range.Cells[row, col] as Excel.Range).Value2;
                        newRow[col - 1] = (null == cellValue) ? "" : cellValue;
                    }
                    _model.arrDataTable[i - 1].Rows.Add(newRow);
                }
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
                _newWorksheet.PageSetup.Zoom = 100;
                _newWorksheet.VPageBreaks.Add(_newWorksheet.Columns[13 + 1]);
                //_newWorksheet.HPageBreaks.Add(_newWorksheet.Rows[44 + 1]);

                _newWorksheet.PageSetup.LeftMargin = _excelApp.InchesToPoints(0.3);
                _newWorksheet.PageSetup.RightMargin = _excelApp.InchesToPoints(0.1);
                _newWorksheet.PageSetup.TopMargin = _excelApp.InchesToPoints(0.1);
                _newWorksheet.PageSetup.BottomMargin = _excelApp.InchesToPoints(0.1);
                _newWorksheet.PageSetup.HeaderMargin = _excelApp.InchesToPoints(0.0);
                _newWorksheet.PageSetup.FooterMargin = _excelApp.InchesToPoints(0.0);
            }
            return 0;
        }
        private int generateLabelUSC(int startIdx, int endIdx)
        {
            for (int idx = startIdx; idx <= endIdx; idx++)
            {
                this._view.setCurrentlyIdx(idx);

                int group = (idx - 1) / 2;
                int rowStartIdx = 1 + (group * 22);
                int colStartIdx = (0 == idx % 2) ? 8 : 1;
                int[] arrDataIdx = new int[2];
                arrDataIdx[0] = arrKTbl[kTblIdx].arrStartIdx[0] + (idx - 1);
                arrDataIdx[1] = arrKTbl[kTblIdx].arrStartIdx[1] + (idx - 1);

                // Setting Excel columns, rows, headers
                SetExcelLayoutUSC(_newWorksheet, rowStartIdx, colStartIdx, arrDataIdx);

                // Paste image
                PasteImage(_newWorksheet, rowStartIdx, colStartIdx);

                // set page break
                if (0 == idx % 4)
                {
                    int rowIdx = rowStartIdx + 19 + 3;
                    _newWorksheet.HPageBreaks.Add(_newWorksheet.Rows[rowIdx]);
                }

            }
            return 0;
        }
        private void SetExcelLayoutUSC(Excel.Worksheet newWorksheet, int rowStartIdx, int colStartIdx, int[] arrDataIdx)
        {
            try
            {
                // start ========================================
                for (int colIdx = 0; 6 > colIdx; colIdx++)
                {
                    Excel.Range columnRange = newWorksheet.Columns[colStartIdx + colIdx];
                    columnRange.ColumnWidth = 11.4;
                }

                Excel.Range rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx, colStartIdx], newWorksheet.Cells[rowStartIdx, colStartIdx + 5]];
                rangeToMerge.Merge();
                rangeToMerge.Value = "Device performance and electrical safety test report";
                _viewHelper.setH1(rangeToMerge);
                _viewHelper.setH2(newWorksheet.Cells[rowStartIdx + 1, colStartIdx], "Type:");
                _viewHelper.setC2(newWorksheet.Cells[rowStartIdx + 1, colStartIdx + 1], "US compact");
                _viewHelper.setH2(newWorksheet.Cells[rowStartIdx + 1, colStartIdx + 4], "Date:");
                _viewHelper.setC2(newWorksheet.Cells[rowStartIdx + 1, colStartIdx + 5], _model.szDate);
                _viewHelper.setH2(newWorksheet.Cells[rowStartIdx + 2, colStartIdx], "SN:");
                _viewHelper.setC2(newWorksheet.Cells[rowStartIdx + 2, colStartIdx + 1], _model.arrDataTable[0].Rows[arrDataIdx[0] - 1][1].ToString());
                _viewHelper.setH2(newWorksheet.Cells[rowStartIdx + 2, colStartIdx + 4], "Inspector:");
                _viewHelper.setC2(newWorksheet.Cells[rowStartIdx + 2, colStartIdx + 5], _model.szInspector);
                // interval =======================================
                _viewHelper.setInterval(newWorksheet.Cells[rowStartIdx + 3, colStartIdx]);
                // header1 ========================================
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx], newWorksheet.Cells[rowStartIdx + 6, colStartIdx]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "RTC         test");
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 1], newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 1]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "EEPROM test");
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 2], newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 2]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Power  button        test");
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 3], newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 3]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "LCD touch pannel test");
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 4], newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 4]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "LCD color patern test");
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 5], newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 5]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "LCD backlight test");
                // content1 =======================================
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 7, colStartIdx], STR_PASS);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 7, colStartIdx + 1], STR_PASS);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 7, colStartIdx + 2], STR_PASS);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 7, colStartIdx + 3], STR_PASS);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 7, colStartIdx + 4], STR_PASS);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 7, colStartIdx + 5], STR_PASS);
                // interval =======================================
                _viewHelper.setInterval(newWorksheet.Cells[rowStartIdx + 8, colStartIdx]);
                // header2 ========================================
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 9, colStartIdx], newWorksheet.Cells[rowStartIdx + 11, colStartIdx]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "USB port   test");
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 1], newWorksheet.Cells[rowStartIdx + 11, colStartIdx + 1]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Handpiece port test");
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 2], newWorksheet.Cells[rowStartIdx + 11, colStartIdx + 2]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Speaker    test");
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 3], newWorksheet.Cells[rowStartIdx + 11, colStartIdx + 3]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Discharging White");
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 4], newWorksheet.Cells[rowStartIdx + 11, colStartIdx + 4]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Charging Green");
                _viewHelper.setH3(newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 5], "Burn-in");
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 10, colStartIdx + 5], newWorksheet.Cells[rowStartIdx + 11, colStartIdx + 5]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Battery (Dis)");
                // content2 =======================================
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 12, colStartIdx], STR_PASS);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 12, colStartIdx + 1], STR_PASS);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 12, colStartIdx + 2], STR_PASS);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 12, colStartIdx + 3], STR_PASS);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 12, colStartIdx + 4], STR_PASS);
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 12, colStartIdx + 5], STR_PASS);
                // interval =======================================
                _viewHelper.setInterval(newWorksheet.Cells[rowStartIdx + 13, colStartIdx]);
                // header3 ========================================
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 14, colStartIdx], newWorksheet.Cells[rowStartIdx + 14, colStartIdx + 2]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Mains On Applied part (MAP test)");
                rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 14, colStartIdx + 3], newWorksheet.Cells[rowStartIdx + 14, colStartIdx + 5]];
                rangeToMerge.Merge();
                _viewHelper.setH3(rangeToMerge, "Leakage Current test");
                _viewHelper.setH3(newWorksheet.Cells[rowStartIdx + 15, colStartIdx], "NC");
                _viewHelper.setH3(newWorksheet.Cells[rowStartIdx + 15, colStartIdx + 1], "NC reverse");
                _viewHelper.setH3(newWorksheet.Cells[rowStartIdx + 15, colStartIdx + 2], "SFC");
                _viewHelper.setH3(newWorksheet.Cells[rowStartIdx + 15, colStartIdx + 3], "NC");
                _viewHelper.setH3(newWorksheet.Cells[rowStartIdx + 15, colStartIdx + 4], "NC reverse");
                _viewHelper.setH3(newWorksheet.Cells[rowStartIdx + 15, colStartIdx + 5], "Result");
                // content3 =======================================
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 16, colStartIdx], float.Parse((string)_model.arrDataTable[1].Rows[arrDataIdx[1] - 1][4 - 1]).ToString("F2") + " uA");
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 16, colStartIdx + 1], float.Parse((string)_model.arrDataTable[1].Rows[arrDataIdx[1] - 1][5 - 1]).ToString("F2") + " uA");
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 16, colStartIdx + 2], float.Parse((string)_model.arrDataTable[1].Rows[arrDataIdx[1] - 1][6 - 1]).ToString("F2") + " uA");
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 16, colStartIdx + 3], float.Parse((string)_model.arrDataTable[1].Rows[arrDataIdx[1] - 1][2 - 1]).ToString("F2") + " uA");
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 16, colStartIdx + 4], float.Parse((string)_model.arrDataTable[1].Rows[arrDataIdx[1] - 1][3 - 1]).ToString("F2") + " uA");
                _viewHelper.setC3(newWorksheet.Cells[rowStartIdx + 16, colStartIdx + 5], STR_PASS);
                // interval =======================================
                _viewHelper.setInterval(newWorksheet.Cells[rowStartIdx + 17, colStartIdx]);
                // sign ===========================================
                newWorksheet.Cells[rowStartIdx + 18, colStartIdx + 3] = "Signature for approval:";
                Excel.Range cell = newWorksheet.Cells[rowStartIdx + 18, colStartIdx + 5];
                _viewHelper.setBottomLine(cell);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"idx:{arrDataIdx[0]}, An unexpected error occurred: " + ex.Message);
            }
        }
    }
}
