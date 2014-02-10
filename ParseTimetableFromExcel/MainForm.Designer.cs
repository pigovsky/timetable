namespace ParseTimetableFromExcel
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.workbookSheetRawDataGrid = new System.Windows.Forms.DataGridView();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.splitContainer = new System.Windows.Forms.SplitContainer();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importFromAnExcelFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exportToMysqlDatabaseToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importFromTNEUSiteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.workbookSheetRawDataGrid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).BeginInit();
            this.splitContainer.Panel1.SuspendLayout();
            this.splitContainer.Panel2.SuspendLayout();
            this.splitContainer.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // workbookSheetRawDataGrid
            // 
            this.workbookSheetRawDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.workbookSheetRawDataGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.workbookSheetRawDataGrid.Location = new System.Drawing.Point(0, 0);
            this.workbookSheetRawDataGrid.Name = "workbookSheetRawDataGrid";
            this.workbookSheetRawDataGrid.Size = new System.Drawing.Size(117, 237);
            this.workbookSheetRawDataGrid.TabIndex = 1;
            // 
            // dataGridView2
            // 
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView2.Location = new System.Drawing.Point(0, 0);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(163, 237);
            this.dataGridView2.TabIndex = 3;
            // 
            // splitContainer
            // 
            this.splitContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer.Location = new System.Drawing.Point(0, 24);
            this.splitContainer.Name = "splitContainer";
            // 
            // splitContainer.Panel1
            // 
            this.splitContainer.Panel1.Controls.Add(this.workbookSheetRawDataGrid);
            // 
            // splitContainer.Panel2
            // 
            this.splitContainer.Panel2.Controls.Add(this.dataGridView2);
            this.splitContainer.Size = new System.Drawing.Size(284, 237);
            this.splitContainer.SplitterDistance = 117;
            this.splitContainer.TabIndex = 4;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(284, 24);
            this.menuStrip1.TabIndex = 5;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.importFromAnExcelFileToolStripMenuItem,
            this.importFromTNEUSiteToolStripMenuItem,
            this.exportToMysqlDatabaseToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // importFromAnExcelFileToolStripMenuItem
            // 
            this.importFromAnExcelFileToolStripMenuItem.Name = "importFromAnExcelFileToolStripMenuItem";
            this.importFromAnExcelFileToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.importFromAnExcelFileToolStripMenuItem.Text = "Import from excel files...";
            this.importFromAnExcelFileToolStripMenuItem.Click += new System.EventHandler(this.importFromMsExcel);
            // 
            // exportToMysqlDatabaseToolStripMenuItem
            // 
            this.exportToMysqlDatabaseToolStripMenuItem.Name = "exportToMysqlDatabaseToolStripMenuItem";
            this.exportToMysqlDatabaseToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.exportToMysqlDatabaseToolStripMenuItem.Text = "Export to mysql database";
            this.exportToMysqlDatabaseToolStripMenuItem.Click += new System.EventHandler(this.exportToMysqlDatabaseToolStripMenuItem_Click);
            // 
            // importFromTNEUSiteToolStripMenuItem
            // 
            this.importFromTNEUSiteToolStripMenuItem.Name = "importFromTNEUSiteToolStripMenuItem";
            this.importFromTNEUSiteToolStripMenuItem.Size = new System.Drawing.Size(206, 22);
            this.importFromTNEUSiteToolStripMenuItem.Text = "Import from TNEU site";
            this.importFromTNEUSiteToolStripMenuItem.Click += new System.EventHandler(this.importFromTNEUSiteToolStripMenuItem_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.splitContainer);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainForm";
            this.Text = "Parse timetable from excel";
            ((System.ComponentModel.ISupportInitialize)(this.workbookSheetRawDataGrid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.splitContainer.Panel1.ResumeLayout(false);
            this.splitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).EndInit();
            this.splitContainer.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView workbookSheetRawDataGrid;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.SplitContainer splitContainer;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importFromAnExcelFileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exportToMysqlDatabaseToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importFromTNEUSiteToolStripMenuItem;
    }
}

