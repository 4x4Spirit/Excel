namespace CreateExcelWorksheet
{
    partial class FormExcel
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
            this.createFile = new System.Windows.Forms.Button();
            this.createRange = new System.Windows.Forms.Button();
            this.readExcel = new System.Windows.Forms.Button();
            this.readAll = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.RESET = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // createFile
            // 
            this.createFile.Location = new System.Drawing.Point(140, 103);
            this.createFile.Name = "createFile";
            this.createFile.Size = new System.Drawing.Size(75, 23);
            this.createFile.TabIndex = 0;
            this.createFile.Text = "Create file";
            this.createFile.UseVisualStyleBackColor = true;
            this.createFile.Click += new System.EventHandler(this.createFile_Click);
            // 
            // createRange
            // 
            this.createRange.Location = new System.Drawing.Point(139, 192);
            this.createRange.Name = "createRange";
            this.createRange.Size = new System.Drawing.Size(75, 35);
            this.createRange.TabIndex = 1;
            this.createRange.Text = "Create Range ";
            this.createRange.UseVisualStyleBackColor = true;
            this.createRange.Click += new System.EventHandler(this.createRange_Click);
            // 
            // readExcel
            // 
            this.readExcel.Location = new System.Drawing.Point(295, 103);
            this.readExcel.Name = "readExcel";
            this.readExcel.Size = new System.Drawing.Size(99, 23);
            this.readExcel.TabIndex = 2;
            this.readExcel.Text = "Read excel file";
            this.readExcel.UseVisualStyleBackColor = true;
            this.readExcel.Click += new System.EventHandler(this.readExcel_Click);
            // 
            // readAll
            // 
            this.readAll.Location = new System.Drawing.Point(295, 191);
            this.readAll.Name = "readAll";
            this.readAll.Size = new System.Drawing.Size(75, 23);
            this.readAll.TabIndex = 3;
            this.readAll.Text = "Read all";
            this.readAll.UseVisualStyleBackColor = true;
            this.readAll.Click += new System.EventHandler(this.readAll_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(140, 275);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(75, 23);
            this.button5.TabIndex = 4;
            this.button5.Text = "Write ";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // RESET
            // 
            this.RESET.AutoSize = true;
            this.RESET.Location = new System.Drawing.Point(220, 381);
            this.RESET.Name = "RESET";
            this.RESET.Size = new System.Drawing.Size(43, 13);
            this.RESET.TabIndex = 5;
            this.RESET.Text = "RESET";
            this.RESET.Click += new System.EventHandler(this.RESET_Click);
            // 
            // FormExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(488, 450);
            this.Controls.Add(this.RESET);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.readAll);
            this.Controls.Add(this.readExcel);
            this.Controls.Add(this.createRange);
            this.Controls.Add(this.createFile);
            this.Name = "FormExcel";
            this.Text = "Excel";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button createFile;
        private System.Windows.Forms.Button createRange;
        private System.Windows.Forms.Button readExcel;
        private System.Windows.Forms.Button readAll;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Label RESET;
    }
}

