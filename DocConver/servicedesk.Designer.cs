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
            this.x = new System.Windows.Forms.NumericUpDown();
            this.y = new System.Windows.Forms.NumericUpDown();
            this.label5 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.iPath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.bedrijfNaam = new System.Windows.Forms.ListBox();
            this.Samenvatting = new System.Windows.Forms.RichTextBox();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.x)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.y)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(41, 173);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(102, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Uploaden";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.excelUploadentToDatabase);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(41, 243);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(132, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "TRUNCATE TABLE data";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.emptyData);
            // 
            // x
            // 
            this.x.Location = new System.Drawing.Point(41, 106);
            this.x.Name = "x";
            this.x.Size = new System.Drawing.Size(120, 20);
            this.x.TabIndex = 16;
            // 
            // y
            // 
            this.y.Location = new System.Drawing.Point(41, 147);
            this.y.Name = "y";
            this.y.Size = new System.Drawing.Size(120, 20);
            this.y.TabIndex = 17;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(41, 90);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(12, 13);
            this.label5.TabIndex = 18;
            this.label5.Text = "x";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(42, 131);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(12, 13);
            this.label7.TabIndex = 20;
            this.label7.Text = "y";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(41, 77);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(68, 13);
            this.label6.TabIndex = 21;
            this.label6.Text = "Eerste kolom";
            // 
            // iPath
            // 
            this.iPath.Location = new System.Drawing.Point(41, 54);
            this.iPath.Name = "iPath";
            this.iPath.Size = new System.Drawing.Size(232, 20);
            this.iPath.TabIndex = 8;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(40, 35);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(54, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "input path";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(279, 51);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(69, 23);
            this.button4.TabIndex = 12;
            this.button4.Text = "Selecteer";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.selectFile);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(569, 35);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 23;
            this.label1.Text = "bedrijf";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(413, 214);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(102, 23);
            this.button3.TabIndex = 24;
            this.button3.Text = "Create view";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.creatView);
            // 
            // bedrijfNaam
            // 
            this.bedrijfNaam.FormattingEnabled = true;
            this.bedrijfNaam.Location = new System.Drawing.Point(572, 51);
            this.bedrijfNaam.Name = "bedrijfNaam";
            this.bedrijfNaam.Size = new System.Drawing.Size(120, 186);
            this.bedrijfNaam.TabIndex = 25;
            // 
            // Samenvatting
            // 
            this.Samenvatting.Location = new System.Drawing.Point(362, 106);
            this.Samenvatting.Name = "Samenvatting";
            this.Samenvatting.Size = new System.Drawing.Size(195, 90);
            this.Samenvatting.TabIndex = 26;
            this.Samenvatting.Text = "";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(423, 77);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(72, 13);
            this.label2.TabIndex = 27;
            this.label2.Text = "Samenvatting";
            // 
            // servicedesk
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(704, 296);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.Samenvatting);
            this.Controls.Add(this.bedrijfNaam);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.y);
            this.Controls.Add(this.x);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.iPath);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "servicedesk";
            this.Text = "s";
            ((System.ComponentModel.ISupportInitialize)(this.x)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.y)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.NumericUpDown x;
        private System.Windows.Forms.NumericUpDown y;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.TextBox iPath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.ListBox bedrijfNaam;
        private System.Windows.Forms.RichTextBox Samenvatting;
        private System.Windows.Forms.Label label2;
    }
}

