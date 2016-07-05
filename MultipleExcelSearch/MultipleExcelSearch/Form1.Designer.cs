namespace MultipleExcelSearch
{
    partial class Form1
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Windows Form-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.SuchenBegriff = new System.Windows.Forms.TextBox();
            this.SuchenText = new System.Windows.Forms.Label();
            this.btnSuchen = new System.Windows.Forms.Button();
            this.listFiles = new System.Windows.Forms.ListView();
            this.FileName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Path = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnAddFile = new System.Windows.Forms.Button();
            this.btnAddFolder = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.treeResult = new System.Windows.Forms.TreeView();
            this.bMultiThread = new System.Windows.Forms.CheckBox();
            this.File = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.Reference = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.listInProgress = new System.Windows.Forms.ListView();
            this.btnDelete = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // SuchenBegriff
            // 
            this.SuchenBegriff.Location = new System.Drawing.Point(112, 68);
            this.SuchenBegriff.Name = "SuchenBegriff";
            this.SuchenBegriff.Size = new System.Drawing.Size(174, 22);
            this.SuchenBegriff.TabIndex = 0;
            // 
            // SuchenText
            // 
            this.SuchenText.AutoSize = true;
            this.SuchenText.Location = new System.Drawing.Point(23, 68);
            this.SuchenText.Name = "SuchenText";
            this.SuchenText.Size = new System.Drawing.Size(83, 17);
            this.SuchenText.TabIndex = 1;
            this.SuchenText.Text = "SuchenText";
            // 
            // btnSuchen
            // 
            this.btnSuchen.Location = new System.Drawing.Point(292, 67);
            this.btnSuchen.Name = "btnSuchen";
            this.btnSuchen.Size = new System.Drawing.Size(96, 23);
            this.btnSuchen.TabIndex = 2;
            this.btnSuchen.Text = "doSuchen";
            this.btnSuchen.UseVisualStyleBackColor = true;
            this.btnSuchen.Click += new System.EventHandler(this.btnSuchen_Click);
            // 
            // listFiles
            // 
            this.listFiles.AllowDrop = true;
            this.listFiles.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.FileName,
            this.Path});
            this.listFiles.Location = new System.Drawing.Point(26, 201);
            this.listFiles.Name = "listFiles";
            this.listFiles.Size = new System.Drawing.Size(391, 188);
            this.listFiles.TabIndex = 3;
            this.listFiles.UseCompatibleStateImageBehavior = false;
            this.listFiles.View = System.Windows.Forms.View.Details;
            this.listFiles.DragDrop += new System.Windows.Forms.DragEventHandler(this.listFiles_DragDrop);
            // 
            // FileName
            // 
            this.FileName.Text = "fileName";
            this.FileName.Width = 314;
            // 
            // Path
            // 
            this.Path.Text = "Path";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(26, 178);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 17);
            this.label1.TabIndex = 4;
            this.label1.Text = "SelectedFile";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(450, 7);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(48, 17);
            this.label2.TabIndex = 4;
            this.label2.Text = "Result";
            // 
            // btnAddFile
            // 
            this.btnAddFile.Location = new System.Drawing.Point(341, 172);
            this.btnAddFile.Name = "btnAddFile";
            this.btnAddFile.Size = new System.Drawing.Size(75, 23);
            this.btnAddFile.TabIndex = 5;
            this.btnAddFile.Text = "AddFiles";
            this.btnAddFile.UseVisualStyleBackColor = true;
            this.btnAddFile.Click += new System.EventHandler(this.btnAddFile_Click);
            // 
            // btnAddFolder
            // 
            this.btnAddFolder.Location = new System.Drawing.Point(198, 172);
            this.btnAddFolder.Name = "btnAddFolder";
            this.btnAddFolder.Size = new System.Drawing.Size(88, 23);
            this.btnAddFolder.TabIndex = 5;
            this.btnAddFolder.Text = "AddFolder";
            this.btnAddFolder.UseVisualStyleBackColor = true;
            this.btnAddFolder.Visible = false;
            this.btnAddFolder.Click += new System.EventHandler(this.btnAddFolder_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "Excel Files(*.xls;*.xlsx)|*.xls;*.xlsx";
            this.openFileDialog1.Multiselect = true;
            // 
            // treeResult
            // 
            this.treeResult.Dock = System.Windows.Forms.DockStyle.Right;
            this.treeResult.Location = new System.Drawing.Point(498, 0);
            this.treeResult.Name = "treeResult";
            this.treeResult.Size = new System.Drawing.Size(281, 605);
            this.treeResult.TabIndex = 6;
            this.treeResult.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.treeResult_MouseDoubleClick);
            // 
            // bMultiThread
            // 
            this.bMultiThread.AutoSize = true;
            this.bMultiThread.Location = new System.Drawing.Point(112, 27);
            this.bMultiThread.Name = "bMultiThread";
            this.bMultiThread.Size = new System.Drawing.Size(100, 21);
            this.bMultiThread.TabIndex = 7;
            this.bMultiThread.Text = "Multithread";
            this.bMultiThread.UseVisualStyleBackColor = true;
            // 
            // File
            // 
            this.File.Text = "File";
            // 
            // Reference
            // 
            this.Reference.Text = "Reference";
            // 
            // listInProgress
            // 
            this.listInProgress.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.File,
            this.Reference});
            this.listInProgress.Enabled = false;
            this.listInProgress.Location = new System.Drawing.Point(26, 480);
            this.listInProgress.MultiSelect = false;
            this.listInProgress.Name = "listInProgress";
            this.listInProgress.Scrollable = false;
            this.listInProgress.Size = new System.Drawing.Size(390, 121);
            this.listInProgress.TabIndex = 3;
            this.listInProgress.UseCompatibleStateImageBehavior = false;
            this.listInProgress.View = System.Windows.Forms.View.Tile;
            // 
            // btnDelete
            // 
            this.btnDelete.Location = new System.Drawing.Point(149, 395);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(124, 23);
            this.btnDelete.TabIndex = 8;
            this.btnDelete.Text = "DeleteFiles";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(26, 457);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(129, 17);
            this.label3.TabIndex = 9;
            this.label3.Text = "Search in Progress";
            // 
            // Form1
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(779, 605);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.bMultiThread);
            this.Controls.Add(this.treeResult);
            this.Controls.Add(this.btnAddFolder);
            this.Controls.Add(this.btnAddFile);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listInProgress);
            this.Controls.Add(this.listFiles);
            this.Controls.Add(this.btnSuchen);
            this.Controls.Add(this.SuchenText);
            this.Controls.Add(this.SuchenBegriff);
            this.Controls.Add(this.btnDelete);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Name = "Form1";
            this.Text = "Multiple Excel Sucher";
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.listFiles_DragDrop);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox SuchenBegriff;
        private System.Windows.Forms.Label SuchenText;
        private System.Windows.Forms.Button btnSuchen;
        private System.Windows.Forms.ListView listFiles;
        private System.Windows.Forms.ColumnHeader FileName;
        private System.Windows.Forms.ColumnHeader Path;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnAddFile;
        private System.Windows.Forms.Button btnAddFolder;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TreeView treeResult;
        private System.Windows.Forms.CheckBox bMultiThread;
        private System.Windows.Forms.ColumnHeader File;
        private System.Windows.Forms.ColumnHeader Reference;
        private System.Windows.Forms.ListView listInProgress;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Label label3;
    }
}

