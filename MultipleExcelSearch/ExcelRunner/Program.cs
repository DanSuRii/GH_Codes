using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelRunner
{
    class Program
    {
        static void Main(string[] args)
        {
            if (2 > args.Length)
                return;

            /*
            if (false == System.Diagnostics.Debugger.IsAttached)
                System.Diagnostics.Debugger.Break();
              */

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

            //oXL->Wait();
            //oRG.Clear();
            oWB.Close();
            oXL.Quit();
        }
    }
}
