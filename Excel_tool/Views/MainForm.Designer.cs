namespace Excel_tool
{
    partial class MainForm
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
            this.btnGenerate = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtDate = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtInspector = new System.Windows.Forms.TextBox();
            this.txt = new System.Windows.Forms.Label();
            this.txt2 = new System.Windows.Forms.Label();
            this.txtPiezoV = new System.Windows.Forms.TextBox();
            this.txtDriverV = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.btnFileBrowse = new System.Windows.Forms.Button();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnOpen = new System.Windows.Forms.Button();
            this.txtResult = new System.Windows.Forms.TextBox();
            this.btnClose = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.cbStartIdx = new System.Windows.Forms.ComboBox();
            this.cbEndIdx = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnGenerate
            // 
            this.btnGenerate.Font = new System.Drawing.Font("PMingLiU", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnGenerate.Location = new System.Drawing.Point(318, 442);
            this.btnGenerate.Margin = new System.Windows.Forms.Padding(2);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(152, 54);
            this.btnGenerate.TabIndex = 0;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(203, 274);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 16);
            this.label1.TabIndex = 1;
            this.label1.Text = "Date:";
            // 
            // txtDate
            // 
            this.txtDate.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtDate.Location = new System.Drawing.Point(248, 270);
            this.txtDate.Margin = new System.Windows.Forms.Padding(2);
            this.txtDate.Name = "txtDate";
            this.txtDate.Size = new System.Drawing.Size(114, 27);
            this.txtDate.TabIndex = 2;
            this.txtDate.Text = "2024/8/20";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(178, 314);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(79, 16);
            this.label2.TabIndex = 3;
            this.label2.Text = "Inspector:";
            // 
            // txtInspector
            // 
            this.txtInspector.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtInspector.Location = new System.Drawing.Point(248, 311);
            this.txtInspector.Margin = new System.Windows.Forms.Padding(2);
            this.txtInspector.Name = "txtInspector";
            this.txtInspector.Size = new System.Drawing.Size(114, 27);
            this.txtInspector.TabIndex = 4;
            this.txtInspector.Text = "tester";
            // 
            // txt
            // 
            this.txt.AutoSize = true;
            this.txt.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt.Location = new System.Drawing.Point(382, 274);
            this.txt.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.txt.Name = "txt";
            this.txt.Size = new System.Drawing.Size(97, 16);
            this.txt.TabIndex = 1;
            this.txt.Text = "Piezo board:";
            // 
            // txt2
            // 
            this.txt2.AutoSize = true;
            this.txt2.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txt2.Location = new System.Drawing.Point(380, 314);
            this.txt2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.txt2.Name = "txt2";
            this.txt2.Size = new System.Drawing.Size(104, 16);
            this.txt2.TabIndex = 1;
            this.txt2.Text = "Driver board:";
            // 
            // txtPiezoV
            // 
            this.txtPiezoV.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtPiezoV.Location = new System.Drawing.Point(485, 269);
            this.txtPiezoV.Margin = new System.Windows.Forms.Padding(2);
            this.txtPiezoV.Name = "txtPiezoV";
            this.txtPiezoV.Size = new System.Drawing.Size(102, 27);
            this.txtPiezoV.TabIndex = 2;
            this.txtPiezoV.Text = "V1.0";
            // 
            // txtDriverV
            // 
            this.txtDriverV.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtDriverV.Location = new System.Drawing.Point(486, 311);
            this.txtDriverV.Margin = new System.Windows.Forms.Padding(2);
            this.txtDriverV.Name = "txtDriverV";
            this.txtDriverV.Size = new System.Drawing.Size(102, 27);
            this.txtDriverV.TabIndex = 2;
            this.txtDriverV.Text = "V1.0";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(203, 361);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(44, 16);
            this.label3.TabIndex = 3;
            this.label3.Text = "Sign:";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.White;
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBox1.Location = new System.Drawing.Point(248, 347);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(67, 38);
            this.pictureBox1.TabIndex = 5;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.DragDrop += new System.Windows.Forms.DragEventHandler(this.pictureBox1_DragDrop);
            this.pictureBox1.DragEnter += new System.Windows.Forms.DragEventHandler(this.pictureBox1_DragEnter);
            // 
            // btnBrowse
            // 
            this.btnBrowse.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnBrowse.Location = new System.Drawing.Point(328, 353);
            this.btnBrowse.Margin = new System.Windows.Forms.Padding(2);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(80, 32);
            this.btnBrowse.TabIndex = 6;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // btnFileBrowse
            // 
            this.btnFileBrowse.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnFileBrowse.Location = new System.Drawing.Point(143, 44);
            this.btnFileBrowse.Margin = new System.Windows.Forms.Padding(2);
            this.btnFileBrowse.Name = "btnFileBrowse";
            this.btnFileBrowse.Size = new System.Drawing.Size(64, 25);
            this.btnFileBrowse.TabIndex = 7;
            this.btnFileBrowse.Text = "Browse";
            this.btnFileBrowse.UseVisualStyleBackColor = true;
            this.btnFileBrowse.Click += new System.EventHandler(this.btnFileBrowse_Click);
            // 
            // txtFilePath
            // 
            this.txtFilePath.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtFilePath.Location = new System.Drawing.Point(222, 44);
            this.txtFilePath.Margin = new System.Windows.Forms.Padding(2);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.ReadOnly = true;
            this.txtFilePath.Size = new System.Drawing.Size(521, 27);
            this.txtFilePath.TabIndex = 8;
            // 
            // btnOpen
            // 
            this.btnOpen.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnOpen.Location = new System.Drawing.Point(145, 88);
            this.btnOpen.Margin = new System.Windows.Forms.Padding(2);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(62, 25);
            this.btnOpen.TabIndex = 9;
            this.btnOpen.Text = "Open";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // txtResult
            // 
            this.txtResult.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.txtResult.Location = new System.Drawing.Point(145, 126);
            this.txtResult.Margin = new System.Windows.Forms.Padding(2);
            this.txtResult.Name = "txtResult";
            this.txtResult.ReadOnly = true;
            this.txtResult.Size = new System.Drawing.Size(314, 27);
            this.txtResult.TabIndex = 10;
            this.txtResult.Text = "No file is opened";
            // 
            // btnClose
            // 
            this.btnClose.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.btnClose.Location = new System.Drawing.Point(222, 88);
            this.btnClose.Margin = new System.Windows.Forms.Padding(2);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(56, 25);
            this.btnClose.TabIndex = 11;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(149, 175);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(62, 16);
            this.label4.TabIndex = 12;
            this.label4.Text = "Start No.";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.Location = new System.Drawing.Point(149, 205);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 16);
            this.label5.TabIndex = 12;
            this.label5.Text = "End No.";
            // 
            // cbStartIdx
            // 
            this.cbStartIdx.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cbStartIdx.FormattingEnabled = true;
            this.cbStartIdx.Location = new System.Drawing.Point(211, 171);
            this.cbStartIdx.Margin = new System.Windows.Forms.Padding(2);
            this.cbStartIdx.Name = "cbStartIdx";
            this.cbStartIdx.Size = new System.Drawing.Size(121, 24);
            this.cbStartIdx.TabIndex = 13;
            // 
            // cbEndIdx
            // 
            this.cbEndIdx.Font = new System.Drawing.Font("PMingLiU", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.cbEndIdx.FormattingEnabled = true;
            this.cbEndIdx.Location = new System.Drawing.Point(211, 200);
            this.cbEndIdx.Margin = new System.Windows.Forms.Padding(2);
            this.cbEndIdx.Name = "cbEndIdx";
            this.cbEndIdx.Size = new System.Drawing.Size(121, 24);
            this.cbEndIdx.TabIndex = 13;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(836, 518);
            this.Controls.Add(this.cbEndIdx);
            this.Controls.Add(this.cbStartIdx);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.txtResult);
            this.Controls.Add(this.btnOpen);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.btnFileBrowse);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.txtInspector);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtDriverV);
            this.Controls.Add(this.txtPiezoV);
            this.Controls.Add(this.txtDate);
            this.Controls.Add(this.txt2);
            this.Controls.Add(this.txt);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnGenerate);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "MainForm";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.MainForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtDate;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtInspector;
        private System.Windows.Forms.Label txt;
        private System.Windows.Forms.Label txt2;
        private System.Windows.Forms.TextBox txtPiezoV;
        private System.Windows.Forms.TextBox txtDriverV;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Button btnFileBrowse;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.TextBox txtResult;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cbStartIdx;
        private System.Windows.Forms.ComboBox cbEndIdx;
    }
}

