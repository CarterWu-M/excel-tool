using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel_tool.Models
{
    public class ExcelModel
    {
        public string szFilePath { get; set; }
        public string szDate { get; set; }
        public string szInspector { get; set; }
        public string szPiezoV { get; set; }
        public string szDriverV {  get; set; }
        public PictureBox pictureBox { get; set; }
        public System.Data.DataTable[] arrDataTable = new System.Data.DataTable[2];

        public ExcelModel() 
        {
            this.szFilePath = string.Empty;
            this.szDate = string.Empty;
            this.szInspector = string.Empty;
            this.szPiezoV = string.Empty;
            this.szDriverV = string.Empty;
            this.arrDataTable[0] = new System.Data.DataTable();
            this.arrDataTable[1] = new System.Data.DataTable();
            this.pictureBox = new PictureBox();
        }
    }
}
