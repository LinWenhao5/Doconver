namespace DocConver
{
    partial class servicedesk
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
            this.components = new System.ComponentModel.Container();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.iPath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(13, 76);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(101, 55);
            this.button1.TabIndex = 0;
            this.button1.Text = "Uploaden";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.excelUploadentToDatabase);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(120, 76);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(102, 55);
            this.button2.TabIndex = 1;
            this.button2.Text = "Tabel leegmaken";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.emptyData);
            // 
            // iPath
            // 
            this.iPath.Location = new System.Drawing.Point(13, 28);
            this.iPath.Name = "iPath";
            this.iPath.Size = new System.Drawing.Size(232, 20);
            this.iPath.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "input path";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(241, 28);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(69, 20);
            this.button4.TabIndex = 12;
            this.button4.Text = "Selecteer";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.selectFile);
            // 
            // servicedesk
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(349, 158);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.iPath);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "servicedesk";
            this.Text = "servicedesk";
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.TextBox iPath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button4;
    }
}

