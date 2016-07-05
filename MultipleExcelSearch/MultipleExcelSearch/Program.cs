using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace MultipleExcelSearch
{
    static class Program
    {
        /// <summary>
        /// Der Haupteinstiegspunkt für die Anwendung.
        /// </summary>
        [STAThread]
/*
        static void Main()
        {
        }
*/

        static void Main(string[] args)
        {
            if (2 > args.Length)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());

            }
            else
            {
                string strPath = args[0];
                string strRef = args[1];

                Excel.Application oXL = new Excel.Application();
                Excel._Workbook oWB = oXL.Workbooks.Open(strPath);

                if (null == oWB)
                {
                    oXL.Quit();
                    return;
                }


                string[] strGoto = strRef.Split('!');
                Excel.Range oRG = oWB.Worksheets[strGoto[0]].Range[strGoto[1]];
                oXL.Goto(oRG);

                oXL.Visible = true;
                oXL.UserControl = true;
            }

            return;

        }
    }
}
