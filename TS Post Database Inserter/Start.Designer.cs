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
            this.OpenMEF = new System.Windows.Forms.Button();
            this.ElL = new System.Windows.Forms.Label();
            this.Launch = new System.Windows.Forms.Button();
            this.OpenLPDF = new System.Windows.Forms.Button();
            this.LpdfL = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.OpenEF = new System.Windows.Forms.Button();
            this.PHExcelL = new System.Windows.Forms.Label();
            this.PDFL = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // OpenMEF
            // 
            this.OpenMEF.Location = new System.Drawing.Point(12, 136);
            this.OpenMEF.Name = "OpenMEF";
            this.OpenMEF.Size = new System.Drawing.Size(131, 23);
            this.OpenMEF.TabIndex = 0;
            this.OpenMEF.Text = "Select Master Excel File";
            this.OpenMEF.UseVisualStyleBackColor = true;
            this.OpenMEF.Click += new System.EventHandler(this.OpenMEF_Click);
            // 
            // ElL
            // 
            this.ElL.AutoSize = true;
            this.ElL.Location = new System.Drawing.Point(149, 141);
            this.ElL.Name = "ElL";
            this.ElL.Size = new System.Drawing.Size(97, 13);
            this.ElL.TabIndex = 1;
            this.ElL.Text = "Master Excel Label";
            // 
            // Launch
            // 
            this.Launch.Location = new System.Drawing.Point(12, 176);
            this.Launch.Name = "Launch";
            this.Launch.Size = new System.Drawing.Size(75, 23);
            this.Launch.TabIndex = 2;
            this.Launch.Text = "Launch";
            this.Launch.UseVisualStyleBackColor = true;
            this.Launch.Click += new System.EventHandler(this.Launch_Click);
            // 
            // OpenLPDF
            // 
            this.OpenLPDF.Location = new System.Drawing.Point(12, 90);
            this.OpenLPDF.Name = "OpenLPDF";
            this.OpenLPDF.Size = new System.Drawing.Size(131, 23);
            this.OpenLPDF.TabIndex = 0;
            this.OpenLPDF.Text = "Select Label PDF File";
            this.OpenLPDF.UseVisualStyleBackColor = true;
            this.OpenLPDF.Click += new System.EventHandler(this.OpenPDF_Click);
            // 
            // LpdfL
            // 
            this.LpdfL.AutoSize = true;
            this.LpdfL.Location = new System.Drawing.Point(149, 95);
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
            // OpenEF
            // 
            this.OpenEF.Location = new System.Drawing.Point(12, 47);
            this.OpenEF.Name = "OpenEF";
            this.OpenEF.Size = new System.Drawing.Size(131, 23);
            this.OpenEF.TabIndex = 0;
            this.OpenEF.Text = "Select Main Excel File";
            this.OpenEF.UseVisualStyleBackColor = true;
            this.OpenEF.Click += new System.EventHandler(this.OpenEF_Click);
            // 
            // PHExcelL
            // 
            this.PHExcelL.AutoSize = true;
            this.PHExcelL.Location = new System.Drawing.Point(149, 52);
            this.PHExcelL.Name = "PHExcelL";
            this.PHExcelL.Size = new System.Drawing.Size(62, 13);
            this.PHExcelL.TabIndex = 1;
            this.PHExcelL.Text = "Excel Label";
            // 
            // PDFL
            // 
            this.PDFL.AutoSize = true;
            this.PDFL.Location = new System.Drawing.Point(149, 118);
            this.PDFL.Name = "PDFL";
            this.PDFL.Size = new System.Drawing.Size(57, 13);
            this.PDFL.TabIndex = 1;
            this.PDFL.Text = "Label PDF";
            // 
            // Start
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(424, 223);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Launch);
            this.Controls.Add(this.PDFL);
            this.Controls.Add(this.LpdfL);
            this.Controls.Add(this.PHExcelL);
            this.Controls.Add(this.ElL);
            this.Controls.Add(this.OpenLPDF);
            this.Controls.Add(this.OpenEF);
            this.Controls.Add(this.OpenMEF);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Start";
            this.Text = "TouchStone Frieght LTD - Delivery Order Edittor";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Start_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button OpenMEF;
        private System.Windows.Forms.Button Launch;
        private System.Windows.Forms.Button OpenLPDF;
        private System.Windows.Forms.Label LpdfL;
        public System.Windows.Forms.Label ElL;
        protected internal System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button OpenEF;
        public System.Windows.Forms.Label PHExcelL;
        private System.Windows.Forms.Label PDFL;
    }
}

