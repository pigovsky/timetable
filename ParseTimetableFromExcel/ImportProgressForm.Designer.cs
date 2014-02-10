namespace ParseTimetableFromExcel
{
    partial class ProgressForm
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
            this.importProgressBar = new System.Windows.Forms.ProgressBar();
            this.buttonStopImport = new System.Windows.Forms.Button();
            this.totalProgressBar = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // importProgressBar
            // 
            this.importProgressBar.Location = new System.Drawing.Point(12, 41);
            this.importProgressBar.Name = "importProgressBar";
            this.importProgressBar.Size = new System.Drawing.Size(496, 23);
            this.importProgressBar.TabIndex = 0;
            // 
            // buttonStopImport
            // 
            this.buttonStopImport.Location = new System.Drawing.Point(432, 74);
            this.buttonStopImport.Name = "buttonStopImport";
            this.buttonStopImport.Size = new System.Drawing.Size(75, 23);
            this.buttonStopImport.TabIndex = 1;
            this.buttonStopImport.Text = "Stop import";
            this.buttonStopImport.UseVisualStyleBackColor = true;
            // 
            // totalProgressBar
            // 
            this.totalProgressBar.Location = new System.Drawing.Point(13, 12);
            this.totalProgressBar.Name = "totalProgressBar";
            this.totalProgressBar.Size = new System.Drawing.Size(496, 23);
            this.totalProgressBar.TabIndex = 2;
            // 
            // ProgressForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(521, 102);
            this.Controls.Add(this.totalProgressBar);
            this.Controls.Add(this.buttonStopImport);
            this.Controls.Add(this.importProgressBar);
            this.Name = "ProgressForm";
            this.Text = "ImportProgressForm";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ProgressBar importProgressBar;
        private System.Windows.Forms.Button buttonStopImport;
        private System.Windows.Forms.ProgressBar totalProgressBar;
    }
}