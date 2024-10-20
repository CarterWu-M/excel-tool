using Excel_Label_tool.Controllers;
using System;
using System.Windows.Forms;

namespace Excel_Label_tool
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Initialize the view and controller
            MainForm view = new MainForm();
            ExcelController controller = new ExcelController(view);

            // Run the form (view)
            Application.Run(view);
        }
    }
}
