namespace TS_Post_Database_Inserter
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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.Launch = new System.Windows.Forms.Button();
            this.LpdfL = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.PHExcelL = new System.Windows.Forms.Label();
            this.PDFL = new System.Windows.Forms.Label();
            this.OpenMFol = new System.Windows.Forms.Button();
            this.MFol = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.SetupMasFold = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Launch
            // 
            this.Launch.Location = new System.Drawing.Point(337, 302);
            this.Launch.Name = "Launch";
            this.Launch.Size = new System.Drawing.Size(75, 23);
            this.Launch.TabIndex = 2;
            this.Launch.Text = "Launch";
            this.Launch.UseVisualStyleBackColor = true;
            this.Launch.Click += new System.EventHandler(this.Launch_Click);
            // 
            // LpdfL
            // 
            this.LpdfL.AutoSize = true;
            this.LpdfL.Location = new System.Drawing.Point(13, 137);
            this.LpdfL.Name = "LpdfL";
            this.LpdfL.Size = new System.Drawing.Size(57, 13);
            this.LpdfL.TabIndex = 1;
            this.LpdfL.Text = "Label PDF";
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
            this.PHExcelL.Location = new System.Drawing.Point(13, 116);
            this.PHExcelL.Name = "PHExcelL";
            this.PHExcelL.Size = new System.Drawing.Size(62, 13);
            this.PHExcelL.TabIndex = 1;
            this.PHExcelL.Text = "Excel Label";
            // 
            // PDFL
            // 
            this.PDFL.AutoSize = true;
            this.PDFL.Location = new System.Drawing.Point(13, 158);
            this.PDFL.Name = "PDFL";
            this.PDFL.Size = new System.Drawing.Size(57, 13);
            this.PDFL.TabIndex = 1;
            this.PDFL.Text = "Label PDF";
            // 
            // OpenMFol
            // 
            this.OpenMFol.Location = new System.Drawing.Point(16, 53);
            this.OpenMFol.Name = "OpenMFol";
            this.OpenMFol.Size = new System.Drawing.Size(131, 23);
            this.OpenMFol.TabIndex = 0;
            this.OpenMFol.Text = "Select Master Folder";
            this.OpenMFol.UseVisualStyleBackColor = true;
            this.OpenMFol.Click += new System.EventHandler(this.OpenMFol_Click);
            // 
            // MFol
            // 
            this.MFol.AutoSize = true;
            this.MFol.Location = new System.Drawing.Point(13, 95);
            this.MFol.Name = "MFol";
            this.MFol.Size = new System.Drawing.Size(71, 13);
            this.MFol.TabIndex = 1;
            this.MFol.Text = "Master Folder";
            // 
            // SetupMasFold
            // 
            this.SetupMasFold.Location = new System.Drawing.Point(16, 252);
            this.SetupMasFold.Name = "SetupMasFold";
            this.SetupMasFold.Size = new System.Drawing.Size(131, 23);
            this.SetupMasFold.TabIndex = 0;
            this.SetupMasFold.Text = "Setup a Master Folder";
            this.SetupMasFold.UseVisualStyleBackColor = true;
            this.SetupMasFold.Click += new System.EventHandler(this.SetupMasFod_Click);
            // 
            // Start
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(424, 337);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Launch);
            this.Controls.Add(this.PDFL);
            this.Controls.Add(this.LpdfL);
            this.Controls.Add(this.PHExcelL);
            this.Controls.Add(this.MFol);
            this.Controls.Add(this.SetupMasFold);
            this.Controls.Add(this.OpenMFol);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
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
        private System.Windows.Forms.Button Launch;
        private System.Windows.Forms.Label LpdfL;
        protected internal System.Windows.Forms.Label label1;
        public System.Windows.Forms.Label PHExcelL;
        private System.Windows.Forms.Label PDFL;
        private System.Windows.Forms.Button OpenMFol;
        public System.Windows.Forms.Label MFol;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button SetupMasFold;
    }
}

