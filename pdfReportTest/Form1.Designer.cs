namespace pdfReportTest
{
    partial class Form1
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
            this.btn_create_report = new System.Windows.Forms.Button();
            this.lblSemenAnalElapsed = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btn_create_report
            // 
            this.btn_create_report.Location = new System.Drawing.Point(91, 92);
            this.btn_create_report.Name = "btn_create_report";
            this.btn_create_report.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.btn_create_report.Size = new System.Drawing.Size(75, 23);
            this.btn_create_report.TabIndex = 0;
            this.btn_create_report.Text = "create report";
            this.btn_create_report.UseVisualStyleBackColor = true;
            this.btn_create_report.Click += new System.EventHandler(this.btn_create_report_Click);
            // 
            // lblSemenAnalElapsed
            // 
            this.lblSemenAnalElapsed.AutoSize = true;
            this.lblSemenAnalElapsed.Location = new System.Drawing.Point(12, 138);
            this.lblSemenAnalElapsed.Name = "lblSemenAnalElapsed";
            this.lblSemenAnalElapsed.Size = new System.Drawing.Size(35, 13);
            this.lblSemenAnalElapsed.TabIndex = 1;
            this.lblSemenAnalElapsed.Text = "label1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.lblSemenAnalElapsed);
            this.Controls.Add(this.btn_create_report);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_create_report;
        private System.Windows.Forms.Label lblSemenAnalElapsed;
    }
}

