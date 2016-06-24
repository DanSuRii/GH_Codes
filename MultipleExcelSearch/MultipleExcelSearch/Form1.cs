using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace MultipleExcelSearch
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnAddFile_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK == openFileDialog1.ShowDialog())
            {
                foreach (string strSelect in openFileDialog1.FileNames)
                {
                    String strFileName = System.IO.Path.GetFileName(strSelect);
                    String strDirectory = System.IO.Path.GetDirectoryName(strSelect);

                    ListViewItem lvItem = new ListViewItem(strFileName);
                    lvItem.SubItems.Add(strDirectory);
                    lvItem.Tag = strSelect;

                    listFiles.Items.Add(lvItem);
                }
            }

        }

        private void btnAddFolder_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK == folderBrowserDialog1.ShowDialog())
            {
                String strPath = folderBrowserDialog1.SelectedPath;

                ListViewItem lvItem = new ListViewItem("*.xls;*.xlsx");
                lvItem.SubItems.Add(strPath);

                listFiles.Items.Add(lvItem);
            }
        }

        private void listFiles_DragDrop(object sender, DragEventArgs e)
        {
            listFiles.Items.Add(e.Data.ToString());
        }

        class TreeNodeExt : TreeNode
        {
            string strFilePath;
            public TreeNodeExt( string ExcelFilePath , TreeNode[] nodeArray )
                : base(System.IO.Path.GetFileName(ExcelFilePath), nodeArray)
            {
                strFilePath = ExcelFilePath;
            }

            public string GetFilePath() { return strFilePath; }
        }

        private bool ExcelFindInWorkbook(ref TreeNode toRet, Excel._Workbook oWB, string strToFind)
        {
            bool bGeFunden = false;

            List<TreeNode> listGeFunden = new List<TreeNode>();

            foreach (Excel._Worksheet oSheet in oWB.Sheets)
            {
                Excel.Range firstFind = null;
                Excel.Range currentFind = oSheet.Cells.Find(strToFind, Type.Missing
                , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart
                , Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false
                , Type.Missing, Type.Missing);

                while (currentFind != null)
                {
                    // Keep track of the first range you find. 
                    if (firstFind == null)
                    {
                        firstFind = currentFind;
                    }

                    // If you didn't move to a new range, you are done.
                    else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                          == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                    {
                        break;
                    }

                    /*
                    currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    currentFind.Font.Bold = true;
                    */

                    // TestCode
                    //var vFind = currentFind.get_Address(Excel.XlReferenceStyle.xlA1);
                    //currentFind.Select();                    
                    listGeFunden.Add(new TreeNode(oSheet.Name + "!" + currentFind.get_Address(Excel.XlReferenceStyle.xlA1)));

                    bGeFunden = true;

                    currentFind = oSheet.Cells.FindNext(currentFind);
                }
            }

            if (0 < listGeFunden.Count)
            {
                TreeNode[] nodeList = listGeFunden.ToArray();

                toRet = new TreeNodeExt(oWB.FullName, nodeList);
            }


            return bGeFunden;
        }

        private bool ExcelFindInWorkbook(Excel._Workbook oWB, string strToFind)
        {

            /*
             TODO: 
             OpenFileDialog Filter *.xls;*.xlsx
              
             MainThread 
             ...
             PostQueuedCompletionStatus( SuchenCode, WorkBook or FilePath )              
              
             Worker Thread,
             GetQueuedCompletionStatus
            {
                ...Arbeit...

                PostMainthreadResult( TreeNode[] );
            }
             */

            TreeNode toRet = null;
            bool bGeFunden = ExcelFindInWorkbook(ref toRet, oWB, strToFind);
            if(null != toRet) treeResult.Nodes.Add(toRet);

#if false
            bool bGeFunden = false;


            List<TreeNode> listGeFunden = new List<TreeNode>();

            foreach (Excel._Worksheet oSheet in oWB.Sheets)
            {
                Excel.Range firstFind = null;
                Excel.Range currentFind = oSheet.Cells.Find(strToFind, Type.Missing
                , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart
                , Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false
                , Type.Missing, Type.Missing);

                while (currentFind != null)
                {
                    // Keep track of the first range you find. 
                    if (firstFind == null)
                    {
                        firstFind = currentFind;
                    }

                    // If you didn't move to a new range, you are done.
                    else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                          == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                    {
                        break;
                    }

                    /*
                    currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    currentFind.Font.Bold = true;
                    */

                    // TestCode
                    //var vFind = currentFind.get_Address(Excel.XlReferenceStyle.xlA1);
                    //currentFind.Select();                    
                    listGeFunden.Add(new TreeNode(oSheet.Name + "!" + currentFind.get_Address(Excel.XlReferenceStyle.xlA1)));

                    bGeFunden = true;

                    currentFind = oSheet.Cells.FindNext(currentFind);
                }
            }

            if (0 < listGeFunden.Count)
            {
                TreeNode[] nodeList = listGeFunden.ToArray();

                TreeNode treeNode = new TreeNodeExt(oWB.FullName, nodeList);
                treeResult.Nodes.Add(treeNode);
            }

#endif
            return bGeFunden;
        }

        private void btnSuchen_Click(object sender, EventArgs e)
        {
            // TODO: stack up the Alter Suche Result at specific Window or etc.
            string strSucheText = SuchenBegriff.Text;

            if(bMultiThread.Checked)    DoSearchInThread(strSucheText);
            else                        DoSearch(strSucheText);

            /*

            Parallel.ForEach<ListViewItem, ListViewItem>(
                listFiles.Items,
                ()=> new ListViewItem() ,
                (item, loop, final) =>
                {
                    string fileName = item.SubItems[0].Text;
                    string path = item.SubItems[1].Text;

                    Excel._Workbook oWB = oXL.Workbooks.Open(System.IO.Path.Combine(path, fileName), ReadOnly: true);
                    //Excel._Workbook oWB = app.Workbooks.Open(path + "\\" + fileName);                
                    if (null == oWB)
                        return final;

                    if (ExcelFindInWorkbook(oWB, strSucheText))
                    {
                        //oXL.Visible = true;
                        //oXL.UserControl = true;
                    }
                    oWB.Close(SaveChanges: false);

                    return final;
                },
                (final) => { }
                );
             */

        }

        private void DoSearch(string strSucheText)
        {

            Excel.Application oXL = new Excel.Application();
            using (ExcelAppHolder oHolder = new ExcelAppHolder(oXL))
                foreach (ListViewItem item in listFiles.Items)
                {
                    string fileName = item.SubItems[0].Text;
                    string path = item.SubItems[1].Text;

                    Excel._Workbook oWB = oXL.Workbooks.Open(System.IO.Path.Combine(path, fileName), ReadOnly: true);
                    //Excel._Workbook oWB = app.Workbooks.Open(path + "\\" + fileName);                
                    if (null == oWB)
                        continue;


                    if (ExcelFindInWorkbook(oWB, strSucheText))
                    {
                        //oXL.Visible = true;
                        //oXL.UserControl = true;
                    }
                    oWB.Close(SaveChanges: false);
                }

        }

        private void DoSearchInThread(string strSucheText)
        {
#if true // Multiple Search in Apolications
            //collect the Filepaths
            List<string> strFileList = new List<string>();
            foreach (ListViewItem item in listFiles.Items)
            {
                string fileName = item.SubItems[0].Text;
                string path = item.SubItems[1].Text;

                strFileList.Add(System.IO.Path.Combine(path, fileName));
            }

            if (0 > strFileList.Count)
                return;

            TreeNode toAdd = new TreeNode("SucheAngriff:" + strSucheText);
            //System.Threading.SpinLock locker = new System.Threading.SpinLock();
            //bool bTake = false;
            object sync = new object();

            Parallel.ForEach<string, TreeNode>(
                strFileList,
                () => null,
                (strPath, loop, childNode) =>
                {
                    Excel.Application oXL = new Excel.Application();
                    using (ExcelAppHolder holderAPP = new ExcelAppHolder(oXL))
                    {
                        Excel._Workbook oWB = oXL.Workbooks.Open(strPath, ReadOnly: true);
                        if (null == oWB)
                            return null;
                        WorkbookHolder holderWB = new WorkbookHolder(oWB);

                        //DONE: Generate in Threadversion
                        if (ExcelFindInWorkbook(ref childNode, oWB, strSucheText))
                        {
                            //oXL.Visible = true;
                            //oXL.UserControl = true;
                        }
                        oWB.Close(SaveChanges: false);
                    }

                    return childNode;
                },
                (childNode) =>
                {
                    if (null != childNode)
                        lock (sync)
                        {
                            toAdd.Nodes.Add(childNode);
                        }
#if false
                //Lock here
                try
                {
                    //locker.Enter(ref bTake);
                    {
                        //lock (sync);
                        if (null != childNode) toAdd.Nodes.Add(childNode);
                    };
                }
                finally
                {
                    //locker.Exit();
                }

#endif
                }
            );

#endif

#if false //Multiple search in Workbook, failure
            List<Excel._Workbook> listWorkBook = new List<Excel._Workbook>();
            Excel.Application oXL = new Excel.Application();
            ExcelAppHolder oHolder = new ExcelAppHolder(oXL);
            TreeNode toAdd = new TreeNode("SucheAngriff:" + strSucheText);

            foreach (ListViewItem item in listFiles.Items)
            {
                string fileName = item.SubItems[0].Text;
                string path = item.SubItems[1].Text;

                Excel._Workbook oWB = oXL.Workbooks.Open(System.IO.Path.Combine(path, fileName), ReadOnly: true);
                //Excel._Workbook oWB = app.Workbooks.Open(path + "\\" + fileName);                
                if (null == oWB)
                    continue;

                WorkbookHolder holderWB = new WorkbookHolder(oWB);
                listWorkBook.Add(oWB);
            }
            Parallel.ForEach<Excel._Workbook, TreeNode>
                (
                    listWorkBook,
                    () => null,
                    (iObjWB, loop, childNode) =>
                    {
                        if (ExcelFindInWorkbook(ref childNode, iObjWB, strSucheText))
                        {
                            //oXL.Visible = true;
                            //oXL.UserControl = true;
                        }
                        iObjWB.Close(SaveChanges: false);
                        //oXL.Quit();

                        return childNode;
                    },
                    (childNode) =>
                    {
                        //Lock here
                        try
                        {
                            //locker.Enter(ref bTake);
                            {
                                //lock (sync);
                                if (null != childNode) toAdd.Nodes.Add(childNode);
                            };
                        }
                        finally
                        {
                            //locker.Exit();
                        }
                    }
               );

#endif

            treeResult.Nodes.Add(toAdd);

        }

        abstract class ObjHolder<T> : IDisposable
        {
            protected T oObj;

            protected ObjHolder(T rhs)
            {
                oObj = rhs;
            }
            ~ObjHolder()
            {
            }

            public abstract void Dispose();
        };

        class ExcelAppHolder : ObjHolder<Excel.Application>
        {
            public ExcelAppHolder(Excel.Application rhs) : base(rhs)
            {
            }

            public override void Dispose()
            {
                oObj.Quit();
            }
        }

        class WorkbookHolder : ObjHolder<Excel._Workbook>
        {
            public WorkbookHolder(Excel._Workbook rhs) : base(rhs)
            {
            }

            public override void Dispose()
            {
                oObj.Close(SaveChanges: false);
            }
        }

        private void treeResult_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            TreeNode node = treeResult.SelectedNode;
            if (null == node.Parent) //nur Arbeit in Child Node
                return;

            if (false == (node.Parent is TreeNodeExt))
                return;

            string strRef = node.Name;
            string strFilePath= ((TreeNodeExt)node.Parent).GetFilePath();

            Excel.Application oXL = new Excel.Application();
            Excel._Workbook oWB = oXL.Workbooks.Open(strFilePath);
            using (ExcelAppHolder holder = new ExcelAppHolder(oXL))
            using (WorkbookHolder wHolder = new WorkbookHolder(oWB))
                do
                {
                    if (null == oWB)
                        break;

                    /*
                    Excel._Worksheet oSheet = oWB.Sheets[];
                    Excel.Range oRG = oSheet.get_Range();
                    */
                 }while(false);

        }
    }
}
