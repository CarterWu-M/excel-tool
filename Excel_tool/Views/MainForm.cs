using Excel_tool.Controllers;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_tool
{
    public partial class MainForm : Form
    {
        private ExcelController _controller;
        public MainForm(ExcelController controller)
        {
            InitializeComponent();
            _controller = controller;
        }

        public Excel.Application excelApp = null;
        public Excel.Workbook workbook = null;
        public Excel.Worksheet worksheet = null;
        public Excel.Worksheet newWorksheet = null;
        public Excel.Range range = null;
        public string gFilePath = "";
        public System.Data.DataTable dataTable = null;

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            _controller.setDate(txtDate.Text);
            _controller.setInspector(txtInspector.Text);
            _controller.setPiezoVer(txtInspector.Text);
            _controller.setDriverVer(txtDriverV.Text);
            _controller.setSignPic(pictureBox1);

            int startIdx = cbStartIdx.SelectedIndex + 1;
            int endIdx = cbEndIdx.SelectedIndex + 1;

            _controller.GenerateReport(startIdx, endIdx);

            // Update the UI
            txtResult.Text = "Report Generated";
            txtResult.BackColor = SystemColors.Window;


            //int idx = 0;
            //try
            //{
            //    int startIdx = cbStartIdx.SelectedIndex + 1;
            //    int endIdx = cbEndIdx.SelectedIndex + 1;

            //    for (idx = startIdx; endIdx >= idx; idx++)
            //    {
            //        int group = (idx - 1) / 2;
            //        int rowStartIdx = 1 + (group * 15);
            //        int colStartIdx = (0 == idx % 2) ? 8 : 1;
            //        int dataIdx = 9 + (idx - 1);
            //        string szLEDResult = "FAIL";


            //        // start ========================================
            //        for (int colIdx = 0; 6 > colIdx; colIdx++)
            //        {
            //            Excel.Range columnRange = newWorksheet.Columns[colStartIdx + colIdx];
            //            columnRange.ColumnWidth = 10.5;
            //        }

            //        Excel.Range rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx, colStartIdx], newWorksheet.Cells[rowStartIdx, colStartIdx + 5]];
            //        rangeToMerge.Merge();
            //        rangeToMerge.Value = "US handpiece H05 test report";
            //        setH1(rangeToMerge);
            //        setH2(newWorksheet.Cells[rowStartIdx + 1, colStartIdx], "Type:");
            //        setC2(newWorksheet.Cells[rowStartIdx + 1, colStartIdx + 1], "US H05");
            //        setH2(newWorksheet.Cells[rowStartIdx + 1, colStartIdx + 4], "Date:");
            //        setC2(newWorksheet.Cells[rowStartIdx + 1, colStartIdx + 5], txtDate.Text);
            //        setH2(newWorksheet.Cells[rowStartIdx + 2, colStartIdx], "SN:");
            //        setC2(newWorksheet.Cells[rowStartIdx + 2, colStartIdx + 1], dataTable.Rows[dataIdx - 1][1].ToString());
            //        setH2(newWorksheet.Cells[rowStartIdx + 2, colStartIdx + 4], "Inspector:");
            //        setC2(newWorksheet.Cells[rowStartIdx + 2, colStartIdx + 5], txtInspector.Text);
            //        // interval ======================================
            //        setInterval(newWorksheet.Cells[rowStartIdx + 3, colStartIdx]);
            //        // header ========================================
            //        rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx], newWorksheet.Cells[rowStartIdx + 5, colStartIdx]];
            //        rangeToMerge.Merge();
            //        setH3(rangeToMerge, "US cup surface");

            //        rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 1], newWorksheet.Cells[rowStartIdx + 5, colStartIdx + 1]];
            //        rangeToMerge.Merge();
            //        setH3(rangeToMerge, "Finishing housing");

            //        rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 2], newWorksheet.Cells[rowStartIdx + 5, colStartIdx + 2]];
            //        rangeToMerge.Merge();
            //        setH3(rangeToMerge, "Finishing gluing");

            //        rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 3], newWorksheet.Cells[rowStartIdx + 5, colStartIdx + 3]];
            //        rangeToMerge.Merge();
            //        setH3(rangeToMerge, "Batch    label");

            //        rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 4], newWorksheet.Cells[rowStartIdx + 5, colStartIdx + 4]];
            //        rangeToMerge.Merge();
            //        setH3(rangeToMerge, "Water proofness");

            //        rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 4, colStartIdx + 5], newWorksheet.Cells[rowStartIdx + 5, colStartIdx + 5]];
            //        rangeToMerge.Merge();
            //        setH3(rangeToMerge, "LED indication");

            //        // content ===================================
            //        setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx], "OK");
            //        setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 1], "OK");
            //        setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 2], "OK");
            //        setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 3], "OK");
            //        setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 4], ("OK" == (string)dataTable.Rows[dataIdx - 1][10]) ? "PASS" : "FAIL");
            //        if ("OK" == (string)dataTable.Rows[dataIdx - 1][2] &&
            //            "OK" == (string)dataTable.Rows[dataIdx - 1][3] &&
            //            "OK" == (string)dataTable.Rows[dataIdx - 1][4] &&
            //            "OK" == (string)dataTable.Rows[dataIdx - 1][5])
            //        {
            //            szLEDResult = "PASS";
            //        }
            //        setC3(newWorksheet.Cells[rowStartIdx + 6, colStartIdx + 5], szLEDResult);
            //        // interval =================================
            //        setInterval(newWorksheet.Cells[rowStartIdx + 7, colStartIdx]);
            //        // header ===================================
            //        rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 8, colStartIdx], newWorksheet.Cells[rowStartIdx + 9, colStartIdx]];
            //        rangeToMerge.Merge();
            //        setH3(rangeToMerge, "Piezo   board");

            //        rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 8, colStartIdx + 1], newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 1]];
            //        rangeToMerge.Merge();
            //        setH3(rangeToMerge, "Driver   board");

            //        rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 8, colStartIdx + 2], newWorksheet.Cells[rowStartIdx + 8, colStartIdx + 3]];
            //        rangeToMerge.Merge();
            //        setH3(rangeToMerge, "Output power @2W/cm2");
            //        setH3(newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 2], "1MHz");
            //        setH3(newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 3], "3MHz");

            //        rangeToMerge = newWorksheet.Range[newWorksheet.Cells[rowStartIdx + 8, colStartIdx + 4], newWorksheet.Cells[rowStartIdx + 8, colStartIdx + 5]];
            //        rangeToMerge.Merge();
            //        setH3(rangeToMerge, "Contact control");
            //        setH3(newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 4], "1MHz");
            //        setH3(newWorksheet.Cells[rowStartIdx + 9, colStartIdx + 5], "3MHz");
            //        // content ================================
            //        setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx], txtPiezoV.Text);
            //        setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx + 1], txtDriverV.Text);
            //        float fV = float.Parse(dataTable.Rows[dataIdx - 1][8].ToString());
            //        setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx + 2], fV.ToString("F2"));
            //        fV = float.Parse(dataTable.Rows[dataIdx - 1][9].ToString());
            //        setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx + 3], fV.ToString("F2"));
            //        setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx + 4], ("OK" == (string)dataTable.Rows[dataIdx - 1][6]) ? "PASS" : "FAIL");
            //        setC3(newWorksheet.Cells[rowStartIdx + 10, colStartIdx + 5], ("OK" == (string)dataTable.Rows[dataIdx - 1][7]) ? "PASS" : "FAIL");
            //        // interval =============================
            //        setInterval(newWorksheet.Cells[rowStartIdx + 11, colStartIdx]);
            //        // sign =================================
            //        newWorksheet.Cells[rowStartIdx + 12, colStartIdx + 3] = "Signature for approval:";
            //        Excel.Range cell = newWorksheet.Cells[rowStartIdx + 12, colStartIdx + 5];
            //        setBottomLine(cell);

            //        Clipboard.SetImage(pictureBox1.Image);
            //        cell.Select();
            //        Thread.Sleep(50);
            //        //newWorksheet.Pictures().Paste();
            //        newWorksheet.Paste();
            //        Excel.Shapes shapes = newWorksheet.Shapes;
            //        Excel.Shape shape = shapes.Item(shapes.Count);
            //        shape.Top = (float)cell.Top - 3;
            //        shape.Left = (float)cell.Left + 5;
            //        // end ==================================
            //    }
            //    txtResult.Text = $"No file is opened";
            //    txtResult.BackColor = SystemColors.Window;
            //    txtResult.Refresh();
            //    cbStartIdx.Items.Clear();
            //    cbEndIdx.Items.Clear();
            //    workbook.SaveAs(txtFilePath.Text);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show($"idx:{idx}, An unexpected error occurred: " + ex.Message);
            //}
            //finally
            //{
            //    if (null != range)
            //    {
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
            //        range = null;
            //    }
            //    if (null != worksheet)
            //    {
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            //        worksheet = null;
            //    }
            //    if (null != newWorksheet)
            //    {
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(newWorksheet);
            //        newWorksheet = null;
            //    }
            //    if (null != workbook)
            //    {
            //        workbook.Close(false);
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            //        workbook = null;
            //    }
            //    if (null != excelApp)
            //    {
            //        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            //        excelApp = null;
            //    }

            //    GC.Collect();
            //    GC.WaitForPendingFinalizers();
            //}

            //void setBottomLine(Excel.Range cell)
            //{
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].ColorIndex = Excel.XlRgbColor.rgbBlack; // Set border color to black
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].TintAndShade = 0;
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin; // Set border weight
            //}

            //void SetBorders(Excel.Range cell)
            //{
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeLeft].Color = Excel.XlRgbColor.rgbBlack;
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeLeft].TintAndShade = 0;
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = Excel.XlBorderWeight.xlThin;

            //    cell.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeRight].Color = Excel.XlRgbColor.rgbBlack;
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeRight].TintAndShade = 0;
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

            //    cell.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeTop].Color = Excel.XlRgbColor.rgbBlack;
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeTop].TintAndShade = 0;
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = Excel.XlBorderWeight.xlThin;

            //    cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = Excel.XlRgbColor.rgbBlack;
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].TintAndShade = 0;
            //    cell.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
            //}


            //void setH1(Excel.Range cell)
            //{
            //    cell.RowHeight = 24.9;
            //    cell.Font.Size = 13;
            //    cell.Font.Bold = true;
            //    cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //    cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //}

            //void setH2(Excel.Range cell, string txt)
            //{
            //    cell.Value = txt;
            //    cell.RowHeight = 18;
            //    cell.Font.Size = 12;
            //    cell.Font.Bold = true;
            //    cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //    cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            //}
            //void setC2(Excel.Range cell, string txt)
            //{
            //    cell.Value = txt;
            //    cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //    cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //    setBottomLine(cell);
            //}

            //void setH3(Excel.Range cell, string txt)
            //{
            //    cell.Value = txt;
            //    cell.RowHeight = 18;
            //    cell.WrapText = true;
            //    cell.Font.Bold = true;
            //    cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //    cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //    cell.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
            //    //cell.NumberFormat = "0.00";
            //    SetBorders(cell);
            //}

            //void setC3(Excel.Range cell, string txt)
            //{
            //    cell.Value = txt;
            //    cell.RowHeight = 25;
            //    cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //    cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //    SetBorders(cell);
            //}

            //void setInterval(Excel.Range cell)
            //{
            //    cell.RowHeight = 6.8;
            //}

        }

        private void pictureBox1_DragDrop(object sender, DragEventArgs e)
        {
            var data = e.Data.GetData(DataFormats.FileDrop);
            if (null != data)
            {
                var fileNames = data as string[];
                if (0 < fileNames.Length)
                {
                    pictureBox1.Image = Image.FromFile(fileNames[0]);
                }
            }    
        }
        private void pictureBox1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }
        private void MainForm_Load(object sender, EventArgs e)
        {
            pictureBox1.AllowDrop = true;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
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
                    string filePath = openFileDialog.FileName;

                    // Load the image into the PictureBox
                    pictureBox1.Image = System.Drawing.Image.FromFile(filePath);
                }
            }
        }

        private void btnFileBrowse_Click(object sender, EventArgs e)
        {
            _controller.BrowseFile();
            txtFilePath.Text = _controller.GetFilePath();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            _controller.OpenExcelFile();
            string result = _controller.CheckValue();

            txtResult.Text = result;
            txtResult.BackColor = (result.Contains("Supported")) ? Color.LightGreen : Color.LightPink;
            txtResult.Refresh();

            cbStartIdx.Items.Clear();
            cbEndIdx.Items.Clear();
            int j = _controller.GetRowCount();
            for (int i = 1; i <= j; i++)
            {
                cbStartIdx.Items.Add(i);
                cbEndIdx.Items.Add(i);
            }
            cbStartIdx.SelectedIndex = 0;
            cbEndIdx.SelectedIndex = j - 1;
        }

        static bool IsFileLocked(string filePath)
        {
            FileStream fileStream = null;

            try
            {
                fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                // The file is unavailable because it is:
                // 1. still being written to
                // 2. being processed by another thread
                // 3. locked by another process
                return true;
            }
            finally
            {
                if (fileStream != null)
                {
                    fileStream.Close();
                }
            }

            return false;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            _controller.CloseExcelFile();

            cbStartIdx.Items.Clear();
            cbEndIdx.Items.Clear();
            txtResult.Text = "No file is opened";
            txtResult.BackColor = SystemColors.Window;
            txtResult.Refresh();
        }
    }
}
