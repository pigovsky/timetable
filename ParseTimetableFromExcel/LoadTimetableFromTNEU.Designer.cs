namespace ParseTimetableFromExcel
{
    partial class LoadTimetableFromTNEU
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
            this.buttonLoad = new System.Windows.Forms.Button();
            this.listBoxTimetableFiles = new System.Windows.Forms.ListBox();
            this.buttonDownload = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonLoad
            // 
            this.buttonLoad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.buttonLoad.Location = new System.Drawing.Point(12, 302);
            this.buttonLoad.Name = "buttonLoad";
            this.buttonLoad.Size = new System.Drawing.Size(116, 23);
            this.buttonLoad.TabIndex = 0;
            this.buttonLoad.Text = "Load timetable list";
            this.buttonLoad.UseVisualStyleBackColor = true;
            this.buttonLoad.Click += new System.EventHandler(this.buttonLoad_Click);
            // 
            // listBoxTimetableFiles
            // 
            this.listBoxTimetableFiles.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listBoxTimetableFiles.FormattingEnabled = true;
            this.listBoxTimetableFiles.Location = new System.Drawing.Point(12, 12);
            this.listBoxTimetableFiles.Name = "listBoxTimetableFiles";
            this.listBoxTimetableFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.listBoxTimetableFiles.Size = new System.Drawing.Size(435, 277);
            this.listBoxTimetableFiles.TabIndex = 1;
            // 
            // buttonDownload
            // 
            this.buttonDownload.Location = new System.Drawing.Point(372, 302);
            this.buttonDownload.Name = "buttonDownload";
            this.buttonDownload.Size = new System.Drawing.Size(75, 23);
            this.buttonDownload.TabIndex = 2;
            this.buttonDownload.Text = "Download";
            this.buttonDownload.UseVisualStyleBackColor = true;
            this.buttonDownload.Click += new System.EventHandler(this.buttonDownload_Click);
            // 
            // LoadTimetableFromTNEU
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(459, 337);
            this.Controls.Add(this.buttonDownload);
            this.Controls.Add(this.listBoxTimetableFiles);
            this.Controls.Add(this.buttonLoad);
            this.Name = "LoadTimetableFromTNEU";
            this.Text = "LoadTimetableFromTNEU";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonLoad;
        private System.Windows.Forms.ListBox listBoxTimetableFiles;
        private System.Windows.Forms.Button buttonDownload;
    }
}