﻿namespace TS_Post_Database_Inserter
{
    partial class Start
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Start));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.Launch = new System.Windows.Forms.Button();
            this.SrcPdfL = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.PHExcelL = new System.Windows.Forms.Label();
            this.SrcLabelsPresentL = new System.Windows.Forms.Label();
            this.MFol = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.CloseBtn = new System.Windows.Forms.Button();
            this.OpenMDIRBTN = new System.Windows.Forms.Button();
            this.PDFNum = new System.Windows.Forms.Label();
            this.moveLeftBtn = new System.Windows.Forms.Button();
            this.moveRightBtn = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.WaitLabel = new System.Windows.Forms.Label();
            this.sourceList = new System.Windows.Forms.CheckedListBox();
            this.downloadsList = new System.Windows.Forms.CheckedListBox();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Launch
            // 
            this.Launch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.Launch.Location = new System.Drawing.Point(595, 431);
            this.Launch.Name = "Launch";
            this.Launch.Size = new System.Drawing.Size(75, 23);
            this.Launch.TabIndex = 2;
            this.Launch.Text = "Launch";
            this.Launch.UseVisualStyleBackColor = true;
            this.Launch.Click += new System.EventHandler(this.Launch_Click);
            // 
            // SrcPdfL
            // 
            this.SrcPdfL.AutoSize = true;
            this.SrcPdfL.Location = new System.Drawing.Point(12, 104);
            this.SrcPdfL.Name = "SrcPdfL";
            this.SrcPdfL.Size = new System.Drawing.Size(159, 13);
            this.SrcPdfL.TabIndex = 1;
            this.SrcPdfL.Text = "Source PDF Location > SrcPdfL";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 20);
            this.label1.TabIndex = 3;
            this.label1.Text = "Settings";
            // 
            // PHExcelL
            // 
            this.PHExcelL.AutoSize = true;
            this.PHExcelL.Location = new System.Drawing.Point(12, 83);
            this.PHExcelL.Name = "PHExcelL";
            this.PHExcelL.Size = new System.Drawing.Size(62, 13);
            this.PHExcelL.TabIndex = 1;
            this.PHExcelL.Text = "Excel Label";
            // 
            // SrcLabelsPresentL
            // 
            this.SrcLabelsPresentL.AutoSize = true;
            this.SrcLabelsPresentL.Location = new System.Drawing.Point(13, 125);
            this.SrcLabelsPresentL.Name = "SrcLabelsPresentL";
            this.SrcLabelsPresentL.Size = new System.Drawing.Size(227, 13);
            this.SrcLabelsPresentL.TabIndex = 1;
            this.SrcLabelsPresentL.Text = "Number of Labels in SRC > SrcLabelsPresentL";
            // 
            // MFol
            // 
            this.MFol.AutoSize = true;
            this.MFol.Location = new System.Drawing.Point(13, 65);
            this.MFol.Name = "MFol";
            this.MFol.Size = new System.Drawing.Size(71, 13);
            this.MFol.TabIndex = 1;
            this.MFol.Text = "Master Folder";
            // 
            // CloseBtn
            // 
            this.CloseBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.CloseBtn.Location = new System.Drawing.Point(16, 431);
            this.CloseBtn.Name = "CloseBtn";
            this.CloseBtn.Size = new System.Drawing.Size(75, 23);
            this.CloseBtn.TabIndex = 2;
            this.CloseBtn.Text = "Close";
            this.CloseBtn.UseVisualStyleBackColor = true;
            this.CloseBtn.Click += new System.EventHandler(this.CloseBtn_Click);
            // 
            // OpenMDIRBTN
            // 
            this.OpenMDIRBTN.Location = new System.Drawing.Point(12, 39);
            this.OpenMDIRBTN.Name = "OpenMDIRBTN";
            this.OpenMDIRBTN.Size = new System.Drawing.Size(131, 23);
            this.OpenMDIRBTN.TabIndex = 0;
            this.OpenMDIRBTN.Text = "Open Master Folder";
            this.OpenMDIRBTN.UseVisualStyleBackColor = true;
            this.OpenMDIRBTN.Click += new System.EventHandler(this.OpenMDIR_Click);
            // 
            // PDFNum
            // 
            this.PDFNum.AutoSize = true;
            this.PDFNum.Location = new System.Drawing.Point(13, 150);
            this.PDFNum.Name = "PDFNum";
            this.PDFNum.Size = new System.Drawing.Size(165, 13);
            this.PDFNum.TabIndex = 1;
            this.PDFNum.Text = "Number of PDF found > PDFNum";
            // 
            // moveLeftBtn
            // 
            this.moveLeftBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.moveLeftBtn.Location = new System.Drawing.Point(301, 230);
            this.moveLeftBtn.Name = "moveLeftBtn";
            this.moveLeftBtn.Size = new System.Drawing.Size(82, 23);
            this.moveLeftBtn.TabIndex = 0;
            this.moveLeftBtn.Text = "<<";
            this.moveLeftBtn.UseVisualStyleBackColor = true;
            this.moveLeftBtn.Click += new System.EventHandler(this.MoveLeftBtn_Click);
            // 
            // moveRightBtn
            // 
            this.moveRightBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.moveRightBtn.Location = new System.Drawing.Point(301, 272);
            this.moveRightBtn.Name = "moveRightBtn";
            this.moveRightBtn.Size = new System.Drawing.Size(82, 23);
            this.moveRightBtn.TabIndex = 0;
            this.moveRightBtn.Text = ">>";
            this.moveRightBtn.UseVisualStyleBackColor = true;
            this.moveRightBtn.Click += new System.EventHandler(this.MoveRightBtn_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(16, 181);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(145, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "PDF List of files in main folder";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(403, 184);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(169, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "PDF List of files in download folder";
            // 
            // WaitLabel
            // 
            this.WaitLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.WaitLabel.AutoSize = true;
            this.WaitLabel.Enabled = false;
            this.WaitLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 100F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.WaitLabel.ForeColor = System.Drawing.Color.Crimson;
            this.WaitLabel.Location = new System.Drawing.Point(189, 301);
            this.WaitLabel.Name = "WaitLabel";
            this.WaitLabel.Size = new System.Drawing.Size(333, 153);
            this.WaitLabel.TabIndex = 7;
            this.WaitLabel.Text = "Wait";
            this.WaitLabel.Visible = false;
            // 
            // sourceList
            // 
            this.sourceList.FormattingEnabled = true;
            this.sourceList.Location = new System.Drawing.Point(19, 197);
            this.sourceList.Name = "sourceList";
            this.sourceList.Size = new System.Drawing.Size(246, 229);
            this.sourceList.TabIndex = 8;
            // 
            // downloadsList
            // 
            this.downloadsList.FormattingEnabled = true;
            this.downloadsList.Location = new System.Drawing.Point(406, 200);
            this.downloadsList.Name = "downloadsList";
            this.downloadsList.Size = new System.Drawing.Size(246, 229);
            this.downloadsList.TabIndex = 8;
            // 
            // Start
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(682, 468);
            this.Controls.Add(this.WaitLabel);
            this.Controls.Add(this.downloadsList);
            this.Controls.Add(this.sourceList);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.CloseBtn);
            this.Controls.Add(this.Launch);
            this.Controls.Add(this.PDFNum);
            this.Controls.Add(this.SrcLabelsPresentL);
            this.Controls.Add(this.SrcPdfL);
            this.Controls.Add(this.PHExcelL);
            this.Controls.Add(this.MFol);
            this.Controls.Add(this.OpenMDIRBTN);
            this.Controls.Add(this.moveRightBtn);
            this.Controls.Add(this.moveLeftBtn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Start";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "TouchStone Frieght LTD - Delivery Order Edittor";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Start_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label SrcPdfL;
        protected internal System.Windows.Forms.Label label1;
        public System.Windows.Forms.Label PHExcelL;
        private System.Windows.Forms.Label SrcLabelsPresentL;
        public System.Windows.Forms.Label MFol;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button CloseBtn;
        private System.Windows.Forms.Button OpenMDIRBTN;
        private System.Windows.Forms.Label PDFNum;
        private System.Windows.Forms.Button Launch;
        private System.Windows.Forms.Button moveLeftBtn;
        private System.Windows.Forms.Button moveRightBtn;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label WaitLabel;
        private System.Windows.Forms.CheckedListBox sourceList;
        private System.Windows.Forms.CheckedListBox downloadsList;
    }
}

