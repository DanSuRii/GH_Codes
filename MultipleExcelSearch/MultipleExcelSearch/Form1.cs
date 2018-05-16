using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace MultipleExcelSearch
{
    public partial class Form1 : Form
    {
        private delegate void DAddTreeItem(TreeNode node);
        private delegate void DRemoveFromProgress(ListViewItem lvItem);
        //DAddTreeItem myTreeAddItem;

        public void AddTreeItem(TreeNode node)
        {
            treeResult.Nodes.Add(node);
        }
        private void RemoveFromProgress(ListViewItem lvItem)
        {
            listInProgress.Items.Remove(lvItem);
        }

        public Form1()
        {
            InitializeComponent();

            //myTreeAddItem = new DAddTreeItem(AddTreeItem);
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
                    lvItem.Tag =  new System.IO.FileInfo( strSelect );

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
            bool bGeFunden = CExcelFindInWorkbook.DoFind(ref toRet, oWB, strToFind);
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
            using (new ExcelAppHolder(oXL))
                foreach (ListViewItem item in listFiles.Items)
                {
                    string fileName = item.SubItems[0].Text;
                    string path = item.SubItems[1].Text;

                    Excel._Workbook oWB = oXL.Workbooks.Open(System.IO.Path.Combine(path, fileName), ReadOnly: true);
                    //Excel._Workbook oWB = app.Workbooks.Open(path + "\\" + fileName);                
                    if (null == oWB)
                        continue;

                    using(new WorkbookHolder(oWB))
                        if (ExcelFindInWorkbook(oWB, strSucheText))
                        {
                            //oXL.Visible = true;
                            //oXL.UserControl = true;
                        }
                   
                }

        }
        private void DoSearchInThread(string strSucheText)
        {
            //TreeNode toAdd = null; // if does not exists any result, Add tree item must crashed.
            ListViewItem lviProgress = listInProgress.Items.Add(strSucheText);
            TreeNode toAdd = new TreeNode("SucheAngriff:" + strSucheText);

            List<System.IO.FileInfo> listFI = new List<System.IO.FileInfo>();
            List<Excel._Workbook> listWorkBook = new List<Excel._Workbook>();

            foreach (ListViewItem item in listFiles.Items)
            {
                if (item.Tag is System.IO.FileInfo)
                {
                    System.IO.FileInfo fI = (System.IO.FileInfo)item.Tag;
                    listFI.Add(fI);
                }
            }

            //using (ExcelAppThreadSafe excel = new ExcelAppThreadSafe())
            {
                System.Threading.Tasks.Task<TreeNode>.Run(
                    ()=>
                    {
                        toAdd = ExcelAppThreadSafe.Instance.FindInWorkBooks(listFI, strSucheText);
                        //this->Invoke( Action()=> { } );
                        //treeResult.Nodes.Add(toAdd);                      
                        Invoke(new DAddTreeItem(AddTreeItem), toAdd);
                        Invoke(new DRemoveFromProgress(RemoveFromProgress), lviProgress);

                    }
                );
            }


            //treeResult.Nodes.Add(toAdd);
        }

        private void DoSearchInThread__OLD(string strSucheText)
        {
            //System.Threading.ThreadPool.QueueUserWorkItem();


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

            //made it Threadpool Version
            //https://msdn.microsoft.com/query/dev14.query?appId=Dev14IDEF1&l=DE-DE&k=k(System.Threading.ThreadPool);k(TargetFrameworkMoniker-.NETFramework,Version%3Dv4.5.2);k(DevLang-csharp)&rd=true
            //https://msdn.microsoft.com/de-de/library/dd321424(v=vs.110).aspx

            var arrExcel = new ExcelAppThreadSafe[4];

            System.Threading.ThreadPool.SetMaxThreads(4, 4);
            List<System.Threading.Tasks.Task<TreeNode>> listTask = new List<System.Threading.Tasks.Task<TreeNode> >();
            foreach ( var strPath in strFileList )
            {
                listTask.Add(
                System.Threading.Tasks.Task<TreeNode>.Run(
                    () =>
                    {
                        TreeNode childNode = null;

                        Excel.Application oXL = new Excel.Application();
                        using (ExcelAppHolder holderAPP = new ExcelAppHolder(oXL))
                        {
                            Excel._Workbook oWB = oXL.Workbooks.Open(strPath, ReadOnly: true);
                            if (null == oWB)
                                return null;
                            WorkbookHolder holderWB = new WorkbookHolder(oWB);

                            //DONE: Generate in Threadversion
                            if (CExcelFindInWorkbook.DoFind(ref childNode, oWB, strSucheText))
                            {
                                //oXL.Visible = true;
                                //oXL.UserControl = true;
                            }
                            oWB.Close(SaveChanges: false);
                        }

                        return childNode;
                    }
                    )
                );
            }
            System.Threading.Tasks.Task.WaitAll( listTask.ToArray() );

            foreach (var task in listTask)
            {
                TreeNode childNode = task.Result;
                if (null != childNode) toAdd.Nodes.Add(childNode);
            }



#if false //Uses Thread Building Block
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
#endif

#if false //Multiple search in Workbook, success, but very slow
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

            }
            foreach( Excel._Workbook oWB in oXL.Workbooks)
                listWorkBook.Add(oWB);

            Parallel.ForEach<Excel._Workbook, TreeNode>
                (
                    listWorkBook,
                    () => null,
                    (iObjWB, loop, childNode) =>
                    {
                        WorkbookHolder hWB = new WorkbookHolder(iObjWB);

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


        private void treeResult_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            TreeNode node = treeResult.SelectedNode;
            if (null == node)
                return;
            if (null == node.Parent) //nur Arbeit in Child Node
                return;
            if (false == (node.Parent is TreeNodeExt))
                return;

            string strRef = node.Text;
            string strFilePath= ((TreeNodeExt)node.Parent).GetFilePath();


            System.Diagnostics.Process.Start("ExcelRunner.exe", '"'+strFilePath + "\" "+ strRef);

#if false
            Excel.Application oXL = new Excel.Application();
            Excel._Workbook oWB = oXL.Workbooks.Open(strFilePath);
#if false
            using (ExcelAppHolder holder = new ExcelAppHolder(oXL))
            using (WorkbookHolder wHolder = new WorkbookHolder(oWB))
                do
                {
                    if (null == oWB)
                        break;

                    oXL.Visible = true;
                    oXL.UserControl = true;
                    /*
                    Excel._Worksheet oSheet = oWB.Sheets[];
                    Excel.Range oRG = oSheet.get_Range();
                    */
                } while (false);

#endif
            if (null == oWB)
            {
                oXL.Quit();
                return;
            }
            //oXL.Goto( ((Excel._Worksheet)oWB.Worksheets[""].Range["A1"]) );
            string[] strGoto = strRef.Split('!');
            Excel.Range oRG = oWB.Worksheets[strGoto[0]].Range[strGoto[1]];
            oXL.Goto(oRG);

            oXL.Visible = true;
            oXL.UserControl = true;

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRG);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oWB);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
            oRG = null;
            oWB = null;
            oXL = null;

#endif
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem eachItem in listFiles.SelectedItems)
            {
                listFiles.Items.Remove(eachItem);
            }
        }
    }

    class TreeNodeExt : TreeNode
    {
        string strFilePath;
        public TreeNodeExt(string ExcelFilePath, TreeNode[] nodeArray)
            : base(System.IO.Path.GetFileName(ExcelFilePath), nodeArray)
        {
            strFilePath = ExcelFilePath;
        }

        public string GetFilePath() { return strFilePath; }
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

    class EventReleaser : ObjHolder<System.Threading.EventWaitHandle>
    {
        public EventReleaser(EventWaitHandle rhs) : base(rhs)
        {
        }

        public override void Dispose()
        {
            oObj.Set();
        }
    }

    //TODO: Create class complete
    class ExcelAppThreadSafe : IDisposable
    {
        private ExcelAppThreadSafe()
        {
            for (int i = 0; i < arrWaitHandles.Count(); ++i)
            {
                if (null == arrExcels[i])
                    arrExcels[i] = new ExcelDatei();
                arrWaitHandles[i] = arrExcels[i].evtHandle;
            }

        }

        
        private static readonly ExcelAppThreadSafe instance = new ExcelAppThreadSafe();
        public static ExcelAppThreadSafe Instance
        {
            get
            {
                return instance;
            }
        }               
        


        class ExcelDatei
        {
            public System.Object objSync = new System.Object();
            public Excel._Application oXL = new Excel.Application();
            public System.Threading.EventWaitHandle evtHandle = new System.Threading.AutoResetEvent(true);

            public ExcelDatei()
            {
                oXL.ScreenUpdating = false;
                oXL.Visible = false;
            }

            public void Dispose()
            {                
                foreach (Excel._Workbook oWB in oXL.Workbooks)
                    oWB.Close(SaveChanges:false);
                    
                oXL.Quit();
            }            
        }
        const int _MaxCnt = 4;

        ExcelDatei[] arrExcels = new ExcelDatei[_MaxCnt];
        System.Threading.WaitHandle[] arrWaitHandles = new System.Threading.WaitHandle[_MaxCnt];

        static int Cnt_ = 0;
        System.Object objSync = new System.Object();

        System.Object objSyncDict = new System.Object();
        private Dictionary<string, System.Object> mapFileLocks = new Dictionary<string, System.Object>();


        public void Dispose()
        {
            foreach (var oED in arrExcels)
                oED.Dispose();
        }

        private Excel._Workbook Open(string strPath)
        {
            int Pos =System.Threading.Interlocked.Increment(ref Cnt_);
            Pos = Pos % _MaxCnt;

            if (null == arrExcels[Pos])
            lock(objSync)
            {
                //prevent intialize redeundent
               if (null == arrExcels[Pos]) arrExcels[Pos] = new ExcelDatei();
            }


            //lock (arrExcels[Pos].objSync)
            {
                //return arrExcels[Pos].oXL.Workbooks.Open(strPath, ReadOnly: true);
                Excel._Application oApp = arrExcels[Pos].oXL;
                Excel.Workbooks oWB = oApp.Workbooks;
                return oWB.Open(strPath, ReadOnly: true);
            }
        }

        internal void OpenWorkBooks( List<System.IO.FileInfo> listFileInfo )
        {
            Parallel.ForEach(listFileInfo, (fileInfo) =>
            {
                Open(fileInfo.FullName);
            });
            //전체 파일을 엑셀에 균등하게 나눠서 오픈 한 뒤 그 리스트를 넘겨준다.
            /*
            foreach ( var excel in arrExcels )
            {
                if(null != excel)
                {
                    foreach (Excel._Workbook oWB in excel.oXL.Workbooks)
                        listWB.Add(oWB);
                }                    
            } 
            */             

            return ;
        }

        class RetType
        {
            public int          nTidx = -1;
            public TreeNode     childNode = null;
        }

        internal void FindInWorkBooks( ref TreeNode toAdd, List<FileInfo> listFI, string strSucheText)
        {
            /*
            System.Threading.Tasks.Task.Run(
                ()=>
                {
                    SendMessage(FindInWorkBooks(listFI, strSucheText));
                }
            );
            */

        }

        internal TreeNode FindInWorkBooks(List<FileInfo> listFI, string strSucheText)
        {
            TreeNode toRet = new TreeNode("SucheAngriff:" + strSucheText);

            List<System.Threading.Tasks.Task<TreeNode>> listTask = new List<System.Threading.Tasks.Task<TreeNode>>();

            foreach (FileInfo fI in listFI)
            {
                listTask.Add(
                     System.Threading.Tasks.Task<TreeNode>.Run(
                     () =>
                     {
                         TreeNode childNode = null;
                         int nIdx = 0;
                         System.Object objFileSync = null;

                         lock (objSyncDict)
                         {
                             try
                             {
                                 objFileSync = mapFileLocks[fI.FullName];
                             }
                             catch (KeyNotFoundException)
                             {
                                 objFileSync = mapFileLocks[fI.FullName] = new System.Object();
                             }
                         }
                         lock (objFileSync) // file sync more unique than ExcelSync
                         {
                             lock (objSync)
                             {
                                 nIdx = System.Threading.WaitHandle.WaitAny(arrWaitHandles);
                             } // get the Arbeit Application


                             lock (arrExcels[nIdx].objSync)
                                 using (new EventReleaser(arrExcels[nIdx].evtHandle))
                                 {
                                     //System.Diagnostics.Debug.Print("Try Open Idx[{0}], File[{1}]", nIdx, fileInfo.FullName);
                                     Excel._Workbook oWB = arrExcels[nIdx].oXL.Workbooks.Open(fI.FullName, ReadOnly: true);
                                     //System.Diagnostics.Debug.Print("Successfully Open Idx[{0}], File[{1}], isNull[{2}]",
                                     //    nIdx, fileInfo.FullName, null == oWB);

                                     if (null == oWB)
                                         return childNode;

                                     using (new WorkbookHolder(oWB))
                                     {
                                         //System.Diagnostics.Debug.Print("Try To Find Idx[{0}], File[{1}], FilePath2[{2}]",
                                         //    nIdx, fileInfo.FullName, oWB.FullName);
                                         CExcelFindInWorkbook.DoFind(ref childNode, oWB, strSucheText);
                                         //System.Diagnostics.Debug.Print("Successfully To Find Idx[{0}], File[{1}], FilePath2[{2}]",
                                         //    nIdx, fileInfo.FullName, oWB.FullName);
                                     }
                                 }

                         }
                         return childNode;
                     }
                     ));
            }

            System.Threading.Tasks.Task.WaitAll(listTask.ToArray());

            foreach (var task in listTask)
            {
                TreeNode childNode = task.Result;
                if (null != childNode) toRet.Nodes.Add(childNode);
            }


            return toRet;
        }

        internal TreeNode FindInWorkBooks__Succ1(List<FileInfo> listFI, string strSucheText)
        {

            TreeNode toRet = new TreeNode("SucheAngriff:" + strSucheText);
            Parallel.ForEach< FileInfo  , RetType>(
                listFI
                , ()=>new RetType()
                , (fileInfo,loop,retIdxNode) =>
                {
                    int nIdx = 0;
                    lock (objSync)
                    {
                        nIdx = System.Threading.WaitHandle.WaitAny(arrWaitHandles);
                    }

                    retIdxNode.nTidx = nIdx;
                    lock(arrExcels[nIdx].objSync)
                        using ( new EventReleaser(arrExcels[nIdx].evtHandle))
                        {
                            //System.Diagnostics.Debug.Print("Try Open Idx[{0}], File[{1}]", nIdx, fileInfo.FullName);
                            Excel._Workbook oWB = arrExcels[nIdx].oXL.Workbooks.Open(fileInfo.FullName, ReadOnly: true);
                            //System.Diagnostics.Debug.Print("Successfully Open Idx[{0}], File[{1}], isNull[{2}]",
                            //    nIdx, fileInfo.FullName, null == oWB);

                            if (null == oWB)
                                return retIdxNode;

                            using (new WorkbookHolder(oWB))
                            {
                                //System.Diagnostics.Debug.Print("Try To Find Idx[{0}], File[{1}], FilePath2[{2}]",
                                //    nIdx, fileInfo.FullName, oWB.FullName);
                                CExcelFindInWorkbook.DoFind(ref retIdxNode.childNode, oWB, strSucheText);
                                //System.Diagnostics.Debug.Print("Successfully To Find Idx[{0}], File[{1}], FilePath2[{2}]",
                                //    nIdx, fileInfo.FullName, oWB.FullName);
                            }
                        }

                    return retIdxNode;
                }
                ,(retIdxNode) =>
                {
                    /* sometimes unable reach to here. Perarrel Library failure.
                    if (retIdxNode.nTidx != -1)
                    {
                        System.Diagnostics.Debug.Print("Try Switch On Idx[{0}]", retIdxNode.nTidx);
                        arrExcels[retIdxNode.nTidx].evtHandle.Set();
                        System.Diagnostics.Debug.Print("Switched On Idx[{0}]", retIdxNode.nTidx);
                    }
                    else
                        System.Diagnostics.Debug.Print("!!!!Achtung!!!! nTidx is -1");
                     */
                    if (null != retIdxNode.childNode)
                        toRet.Nodes.Add(retIdxNode.childNode);
                }
                );
            return toRet;
        }

        internal TreeNode FindInWorkBooks__SuccButSlow(List<FileInfo> listFI, string strSucheText)
        {
            OpenWorkBooks(listFI);

            TreeNode toRet = new TreeNode("SucheAngriff: " + strSucheText);
            Parallel.ForEach<ExcelDatei, TreeNode>(
                arrExcels
                ,()=>null
                ,(iObjExcel,loop,childNode) =>
                {                    
                    foreach( Excel._Workbook oWB in iObjExcel.oXL.Workbooks)
                        CExcelFindInWorkbook.DoFind( ref childNode, oWB, strSucheText );
                    return childNode;
                }
                ,(childNode)=>
                {
                    if(null != childNode) toRet.Nodes.Add(childNode);
                });

            return toRet;
        }
    }

    class CExcelFindInWorkbook
    {
        static public bool DoFind(ref TreeNode toRet, Excel._Workbook oWB, string strToFind)
        {
            bool bGeFunden = false;

            List<TreeNode> listGeFunden = new List<TreeNode>();

            foreach (Excel._Worksheet oSheet in oWB.Worksheets)
            {
                if (Excel.XlSheetType.xlWorksheet != oSheet.Type)
                    continue;

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
    }

}
