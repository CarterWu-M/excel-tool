using System.Windows.Forms;
using System.Data;

namespace Excel_Label_tool.Models
{
    public class ExcelModel
    {
        #region Properties
        public string szFilePath { get; set; }
        public string szDate { get; set; }
        public string szInspector { get; set; }
        public string szPiezoV { get; set; }
        public string szDriverV {  get; set; }
        public PictureBox pictureBox { get; set; }
        public DataTable[] arrDataTable { get; set; }
        #endregion

        #region Constructors
        public ExcelModel() 
        {
            this.szFilePath = string.Empty;
            this.szDate = string.Empty;
            this.szInspector = string.Empty;
            this.szPiezoV = string.Empty;
            this.szDriverV = string.Empty;
            this.pictureBox = new PictureBox();
            arrDataTable = new DataTable[2];
            this.arrDataTable[0] = new DataTable();
            this.arrDataTable[1] = new DataTable();
        }
        #endregion
    }
}
