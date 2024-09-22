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
