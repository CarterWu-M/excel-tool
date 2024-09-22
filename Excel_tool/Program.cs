using Excel_tool.Controllers;
using Excel_tool.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel_tool
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

            // Initialize the model and controller
            ExcelModel model = new ExcelModel();
            ExcelController controller = new ExcelController(model);

            // Run the form (view) and pass the controller
            Application.Run(new MainForm(controller));
        }
    }
}
