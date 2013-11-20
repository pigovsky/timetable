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
            this.SuspendLayout();
            // 
            // importProgressBar
            // 
            this.importProgressBar.Location = new System.Drawing.Point(13, 13);
            this.importProgressBar.Name = "importProgressBar";
            this.importProgressBar.Size = new System.Drawing.Size(496, 23);
            this.importProgressBar.TabIndex = 0;
            // 
            // buttonStopImport
            // 
            this.buttonStopImport.Location = new System.Drawing.Point(433, 46);
            this.buttonStopImport.Name = "buttonStopImport";
            this.buttonStopImport.Size = new System.Drawing.Size(75, 23);
            this.buttonStopImport.TabIndex = 1;
            this.buttonStopImport.Text = "Stop import";
            this.buttonStopImport.UseVisualStyleBackColor = true;
            // 
            // ImportProgressForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(521, 81);
            this.Controls.Add(this.buttonStopImport);
            this.Controls.Add(this.importProgressBar);
            this.Name = "ImportProgressForm";
            this.Text = "ImportProgressForm";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ProgressBar importProgressBar;
        private System.Windows.Forms.Button buttonStopImport;
    }
}