namespace TS_Post_Database_Inserter
{
    partial class Completed
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
            this.OKBtn = new System.Windows.Forms.Button();
            this.OnScreenText = new System.Windows.Forms.Label();
            this.Progressbar = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // OKBtn
            // 
            this.OKBtn.Enabled = false;
            this.OKBtn.Location = new System.Drawing.Point(238, 82);
            this.OKBtn.Name = "OKBtn";
            this.OKBtn.Size = new System.Drawing.Size(75, 23);
            this.OKBtn.TabIndex = 0;
            this.OKBtn.Text = "OK";
            this.OKBtn.UseVisualStyleBackColor = true;
            this.OKBtn.Click += new System.EventHandler(this.OKBtn_Click);
            // 
            // OnScreenText
            // 
            this.OnScreenText.AutoSize = true;
            this.OnScreenText.Location = new System.Drawing.Point(10, 14);
            this.OnScreenText.Name = "OnScreenText";
            this.OnScreenText.Size = new System.Drawing.Size(122, 13);
            this.OnScreenText.TabIndex = 1;
            this.OnScreenText.Text = "Text will change by itself";
            this.OnScreenText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Progressbar
            // 
            this.Progressbar.Location = new System.Drawing.Point(25, 88);
            this.Progressbar.Maximum = 200;
            this.Progressbar.Name = "Progressbar";
            this.Progressbar.Size = new System.Drawing.Size(205, 10);
            this.Progressbar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.Progressbar.TabIndex = 2;
            // 
            // Completed
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(325, 117);
            this.ControlBox = false;
            this.Controls.Add(this.Progressbar);
            this.Controls.Add(this.OnScreenText);
            this.Controls.Add(this.OKBtn);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Completed";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Completed";
            this.Shown += new System.EventHandler(this.Completed_Shown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button OKBtn;
        private System.Windows.Forms.Label OnScreenText;
        public System.Windows.Forms.ProgressBar Progressbar;
    }
}