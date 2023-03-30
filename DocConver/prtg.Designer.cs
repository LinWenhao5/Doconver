namespace DocConver
{
    partial class prtg
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
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.iPath = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(40, 78);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(88, 60);
            this.button1.TabIndex = 30;
            this.button1.Text = "data lezen";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.readData);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(134, 77);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(103, 61);
            this.button2.TabIndex = 31;
            this.button2.Text = "Tabel leegmaken";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.emptyPrtg);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(267, 34);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(69, 21);
            this.button4.TabIndex = 34;
            this.button4.Text = "Selecteer";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.selectFile);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(37, 19);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 13);
            this.label3.TabIndex = 33;
            this.label3.Text = "input path";
            // 
            // iPath
            // 
            this.iPath.Location = new System.Drawing.Point(40, 35);
            this.iPath.Name = "iPath";
            this.iPath.Size = new System.Drawing.Size(232, 20);
            this.iPath.TabIndex = 32;
            // 
            // prtg
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(357, 159);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.iPath);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "prtg";
            this.Text = "RPTG";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox iPath;
    }
}