using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel_tool.Views
{
    public interface IView
    {
        //view -> event trigger -> controller
        event EventHandler browseExcelFile;
        event EventHandler openExcelFile;
        event EventHandler closeExcelFile;
        event EventHandler browseImageFile;
        event EventHandler generateReport;

        //controller -> set -> view
        void setCurrentlyIdx(int idx);
        void setExcelPath(string szPath);
        void setOpenResult(string szResult, int rowCnt);
        void resetOpenResult();
        void setImageFile(string szPath);

        //controller <- get <- view
        string getDate();
        string getInspector();
        string getPiezoVer();//for HP
        string getDeiverVer();//for HP
        PictureBox getImageObj();
        int getStartIdx();
        int getEndIdx();
    }
}
