namespace VBAConverterApp
{
    partial class VBAConverter
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
            this.btnProcess = new System.Windows.Forms.Button();
            this.pnlHeader = new System.Windows.Forms.Panel();
            this.btnMasterRecipeBrowse = new System.Windows.Forms.Button();
            this.txtMasterRecipePath = new System.Windows.Forms.TextBox();
            this.btnBomBrowse = new System.Windows.Forms.Button();
            this.txtBomPath = new System.Windows.Forms.TextBox();
            this.panDetail = new System.Windows.Forms.Panel();
            this.rtbProcess = new System.Windows.Forms.RichTextBox();
            this.pnlHeader.SuspendLayout();
            this.panDetail.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnProcess
            // 
            this.btnProcess.Location = new System.Drawing.Point(713, 71);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(75, 23);
            this.btnProcess.TabIndex = 0;
            this.btnProcess.Text = "Convert";
            this.btnProcess.UseVisualStyleBackColor = true;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // pnlHeader
            // 
            this.pnlHeader.Controls.Add(this.btnMasterRecipeBrowse);
            this.pnlHeader.Controls.Add(this.txtMasterRecipePath);
            this.pnlHeader.Controls.Add(this.btnBomBrowse);
            this.pnlHeader.Controls.Add(this.txtBomPath);
            this.pnlHeader.Controls.Add(this.btnProcess);
            this.pnlHeader.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlHeader.Location = new System.Drawing.Point(0, 0);
            this.pnlHeader.Name = "pnlHeader";
            this.pnlHeader.Size = new System.Drawing.Size(800, 100);
            this.pnlHeader.TabIndex = 1;
            // 
            // btnMasterRecipeBrowse
            // 
            this.btnMasterRecipeBrowse.Location = new System.Drawing.Point(638, 36);
            this.btnMasterRecipeBrowse.Name = "btnMasterRecipeBrowse";
            this.btnMasterRecipeBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnMasterRecipeBrowse.TabIndex = 4;
            this.btnMasterRecipeBrowse.Text = "Browse";
            this.btnMasterRecipeBrowse.UseVisualStyleBackColor = true;
            this.btnMasterRecipeBrowse.Click += new System.EventHandler(this.btnMasterRecipeBrowse_Click);
            // 
            // txtMasterRecipePath
            // 
            this.txtMasterRecipePath.Location = new System.Drawing.Point(12, 38);
            this.txtMasterRecipePath.Name = "txtMasterRecipePath";
            this.txtMasterRecipePath.Size = new System.Drawing.Size(620, 20);
            this.txtMasterRecipePath.TabIndex = 3;
            this.txtMasterRecipePath.Text = "Master Recipe File";
            // 
            // btnBomBrowse
            // 
            this.btnBomBrowse.Location = new System.Drawing.Point(638, 9);
            this.btnBomBrowse.Name = "btnBomBrowse";
            this.btnBomBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBomBrowse.TabIndex = 2;
            this.btnBomBrowse.Text = "Browse";
            this.btnBomBrowse.UseVisualStyleBackColor = true;
            this.btnBomBrowse.Click += new System.EventHandler(this.btnBomBrowse_Click);
            // 
            // txtBomPath
            // 
            this.txtBomPath.Location = new System.Drawing.Point(12, 12);
            this.txtBomPath.Name = "txtBomPath";
            this.txtBomPath.Size = new System.Drawing.Size(620, 20);
            this.txtBomPath.TabIndex = 1;
            this.txtBomPath.Text = "Bom File";
            // 
            // panDetail
            // 
            this.panDetail.Controls.Add(this.rtbProcess);
            this.panDetail.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panDetail.Location = new System.Drawing.Point(0, 100);
            this.panDetail.Name = "panDetail";
            this.panDetail.Size = new System.Drawing.Size(800, 660);
            this.panDetail.TabIndex = 2;
            // 
            // rtbProcess
            // 
            this.rtbProcess.BackColor = System.Drawing.Color.Black;
            this.rtbProcess.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rtbProcess.Location = new System.Drawing.Point(0, 0);
            this.rtbProcess.Name = "rtbProcess";
            this.rtbProcess.Size = new System.Drawing.Size(800, 660);
            this.rtbProcess.TabIndex = 0;
            this.rtbProcess.Text = "";
            // 
            // VBAConverter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 760);
            this.Controls.Add(this.panDetail);
            this.Controls.Add(this.pnlHeader);
            this.Name = "VBAConverter";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "VBA Converter";
            this.pnlHeader.ResumeLayout(false);
            this.pnlHeader.PerformLayout();
            this.panDetail.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.Panel pnlHeader;
        private System.Windows.Forms.TextBox txtBomPath;
        private System.Windows.Forms.Panel panDetail;
        private System.Windows.Forms.Button btnBomBrowse;
        private System.Windows.Forms.Button btnMasterRecipeBrowse;
        private System.Windows.Forms.TextBox txtMasterRecipePath;
        private System.Windows.Forms.RichTextBox rtbProcess;
    }
}

