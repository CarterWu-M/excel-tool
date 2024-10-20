using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel_Label_tool.Views;

namespace Excel_Label_tool
{
    public partial class MainForm : Form, IView
    {
        #region Fields
        private Timer _timer;
        private int _elapsedSeconds = 0;
        #endregion

        #region Constructor
        public MainForm()
        {
            InitializeComponent();

            this._timer = new Timer();
            this._timer.Interval = 1000;
            this._timer.Tick += Timer_tick;
        }
        #endregion

        #region IView
        // =====================================
        // event trigger
        // =====================================
        public event EventHandler browseExcelFile;
        public event EventHandler openExcelFile;
        public event EventHandler closeExcelFile;
        public event EventHandler browseImageFile;
        public event EventHandler generateReport;

        // =====================================
        // set API
        // =====================================
        public void setExcelPath(string szPath)
        {
            txtFilePath.Text = szPath;
        }
        public void setCurrentlyIdx(int idx)
        {
            txtCurrIdx.Invoke(new System.Action(() =>
            {
                txtCurrIdx.Text = idx.ToString();
            }));
        }
        public void setOpenResult(string szResult, int rowCnt)
        {
            txtResult.Text = szResult;
            txtResult.BackColor = (szResult.Contains("Supported")) ? Color.LightGreen : Color.LightPink;
            txtResult.Refresh();
            if (!szResult.Contains("Supported"))
            {
                return;
            }

            cbStartIdx.Items.Clear();
            cbEndIdx.Items.Clear();
            int j = rowCnt;
            for (int i = 1; i <= j; i++)
            {
                cbStartIdx.Items.Add(i);
                cbEndIdx.Items.Add(i);
            }
            cbStartIdx.SelectedIndex = 0;
            cbEndIdx.SelectedIndex = j - 1;
        }
        public void resetOpenResult()
        {
            this.Invoke(new System.Action(() =>
            {
                cbStartIdx.Items.Clear();
                cbEndIdx.Items.Clear();
                txtResult.Text = "No file is opened";
                txtResult.BackColor = SystemColors.Control;
                txtResult.Refresh();
            }));
        }
        public void setImageFile(string szPath)
        {
            pictureBox1.Image = System.Drawing.Image.FromFile(szPath);
        }

        // =====================================
        // get API
        // =====================================
        public string getDate()
        {
            return txtDate.Text;
        }
        public string getInspector()
        {
            return txtInspector.Text;
        }
        public string getPiezoVer()//for HP
        {
            return txtPiezoV.Text;
        }
        public string getDeiverVer()//for HP
        {
            return txtDriverV.Text;
        }
        public PictureBox getImageObj()
        {
            return pictureBox1;
        }
        public int getStartIdx()
        {
            return (int)cbStartIdx.Invoke(new Func<int>(() =>
            {
                return cbStartIdx.SelectedIndex + 1;
            }));
        }
        public int getEndIdx()
        {
            return (int)cbEndIdx.Invoke(new Func<int>(() =>
            {
                return cbEndIdx.SelectedIndex + 1;
            }));
        }
        #endregion

        #region Event
        private void Timer_tick(object sender, EventArgs e)
        {
            this._elapsedSeconds++;
            TimeSpan timeSpan = TimeSpan.FromSeconds(this._elapsedSeconds);

            this.Invoke(new System.Action(() =>
            {
                txtTime.Text = timeSpan.ToString(@"hh\:mm\:ss");
                txtTime.Refresh();
            }));
        }
        private async void btnGenerate_Click(object sender, EventArgs e)
        {
            this._elapsedSeconds = 0;
            txtTime.Text = "00:00:00";
            txtTime.Refresh();
            btnGenerate.BackColor = Color.Plum;
            btnGenerate.Enabled = false;
            btnGenerate.Refresh();
            this._timer.Start();

            //this uses MTA: Multithreaded Apartment
            await Task.Run(() =>
            {
                this.generateReport?.Invoke(this, EventArgs.Empty);
            });
  
            this.resetOpenResult();
            this._timer.Stop();
            btnGenerate.BackColor = Color.YellowGreen;
            btnGenerate.Enabled = true;
        }
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            this.browseImageFile?.Invoke(this, EventArgs.Empty);
        }
        private void btnFileBrowse_Click(object sender, EventArgs e)
        {
            this.browseExcelFile?.Invoke(this, EventArgs.Empty);
        }
        private void btnOpen_Click(object sender, EventArgs e)
        {
            this.openExcelFile?.Invoke(this, EventArgs.Empty);
        }
        private void btnClose_Click(object sender, EventArgs e)
        {
            this.closeExcelFile?.Invoke(this, EventArgs.Empty);
        }
        #endregion

        #region Image_Drag_Drop_Event
        // =============================================
        // image drag & drop event
        // =============================================
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
        #endregion
    }
}
