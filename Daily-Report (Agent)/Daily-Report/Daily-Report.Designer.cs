namespace Daily_Report
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.OutputPath = new System.Windows.Forms.Button();
            this.txtOutputPath = new System.Windows.Forms.TextBox();
            this.OpenPath1 = new System.Windows.Forms.Button();
            this.TxtSelectExcelFile1 = new System.Windows.Forms.TextBox();
            this.Run_Reports = new System.Windows.Forms.Button();
            this.OpenPath2 = new System.Windows.Forms.Button();
            this.TxtSelectExcelFile2 = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.txtStartDate = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtbatchNum = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // OutputPath
            // 
            this.OutputPath.Location = new System.Drawing.Point(439, 161);
            this.OutputPath.Margin = new System.Windows.Forms.Padding(4);
            this.OutputPath.Name = "OutputPath";
            this.OutputPath.Size = new System.Drawing.Size(115, 28);
            this.OutputPath.TabIndex = 47;
            this.OutputPath.Text = "Output File...";
            this.OutputPath.UseVisualStyleBackColor = true;
            this.OutputPath.Click += new System.EventHandler(this.OutputPath_Click);
            // 
            // txtOutputPath
            // 
            this.txtOutputPath.Enabled = false;
            this.txtOutputPath.HideSelection = false;
            this.txtOutputPath.Location = new System.Drawing.Point(11, 162);
            this.txtOutputPath.Margin = new System.Windows.Forms.Padding(4);
            this.txtOutputPath.Name = "txtOutputPath";
            this.txtOutputPath.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.txtOutputPath.Size = new System.Drawing.Size(401, 22);
            this.txtOutputPath.TabIndex = 48;
            this.txtOutputPath.TabStop = false;
            this.txtOutputPath.Text = "Output path file...";
            // 
            // OpenPath1
            // 
            this.OpenPath1.Location = new System.Drawing.Point(439, 12);
            this.OpenPath1.Margin = new System.Windows.Forms.Padding(4);
            this.OpenPath1.Name = "OpenPath1";
            this.OpenPath1.Size = new System.Drawing.Size(115, 28);
            this.OpenPath1.TabIndex = 44;
            this.OpenPath1.Text = "Open File...";
            this.OpenPath1.UseVisualStyleBackColor = true;
            this.OpenPath1.Click += new System.EventHandler(this.OpenPath1_Click);
            // 
            // TxtSelectExcelFile1
            // 
            this.TxtSelectExcelFile1.Enabled = false;
            this.TxtSelectExcelFile1.HideSelection = false;
            this.TxtSelectExcelFile1.Location = new System.Drawing.Point(11, 15);
            this.TxtSelectExcelFile1.Margin = new System.Windows.Forms.Padding(4);
            this.TxtSelectExcelFile1.Name = "TxtSelectExcelFile1";
            this.TxtSelectExcelFile1.ReadOnly = true;
            this.TxtSelectExcelFile1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.TxtSelectExcelFile1.Size = new System.Drawing.Size(401, 22);
            this.TxtSelectExcelFile1.TabIndex = 46;
            this.TxtSelectExcelFile1.TabStop = false;
            this.TxtSelectExcelFile1.Text = "Select From List Client...";
            // 
            // Run_Reports
            // 
            this.Run_Reports.Location = new System.Drawing.Point(436, 218);
            this.Run_Reports.Margin = new System.Windows.Forms.Padding(4);
            this.Run_Reports.Name = "Run_Reports";
            this.Run_Reports.Size = new System.Drawing.Size(115, 28);
            this.Run_Reports.TabIndex = 45;
            this.Run_Reports.Text = "Run";
            this.Run_Reports.UseVisualStyleBackColor = true;
            this.Run_Reports.Click += new System.EventHandler(this.Run_Reports_Click);
            // 
            // OpenPath2
            // 
            this.OpenPath2.Location = new System.Drawing.Point(436, 57);
            this.OpenPath2.Margin = new System.Windows.Forms.Padding(4);
            this.OpenPath2.Name = "OpenPath2";
            this.OpenPath2.Size = new System.Drawing.Size(115, 28);
            this.OpenPath2.TabIndex = 52;
            this.OpenPath2.Text = "Open File...";
            this.OpenPath2.UseVisualStyleBackColor = true;
            this.OpenPath2.Click += new System.EventHandler(this.OpenPath2_Click);
            // 
            // TxtSelectExcelFile2
            // 
            this.TxtSelectExcelFile2.Enabled = false;
            this.TxtSelectExcelFile2.HideSelection = false;
            this.TxtSelectExcelFile2.Location = new System.Drawing.Point(8, 59);
            this.TxtSelectExcelFile2.Margin = new System.Windows.Forms.Padding(4);
            this.TxtSelectExcelFile2.Name = "TxtSelectExcelFile2";
            this.TxtSelectExcelFile2.ReadOnly = true;
            this.TxtSelectExcelFile2.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.TxtSelectExcelFile2.Size = new System.Drawing.Size(401, 22);
            this.TxtSelectExcelFile2.TabIndex = 53;
            this.TxtSelectExcelFile2.TabStop = false;
            this.TxtSelectExcelFile2.Text = "Select Database...";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // txtStartDate
            // 
            this.txtStartDate.HideSelection = false;
            this.txtStartDate.Location = new System.Drawing.Point(8, 218);
            this.txtStartDate.Margin = new System.Windows.Forms.Padding(4);
            this.txtStartDate.Name = "txtStartDate";
            this.txtStartDate.Size = new System.Drawing.Size(401, 22);
            this.txtStartDate.TabIndex = 54;
            this.txtStartDate.TabStop = false;
            this.txtStartDate.Text = "3/27/2018";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Location = new System.Drawing.Point(7, 199);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 17);
            this.label1.TabIndex = 55;
            this.label1.Text = "M/D/YYYY";
            // 
            // txtbatchNum
            // 
            this.txtbatchNum.Enabled = false;
            this.txtbatchNum.HideSelection = false;
            this.txtbatchNum.Location = new System.Drawing.Point(136, 130);
            this.txtbatchNum.Margin = new System.Windows.Forms.Padding(4);
            this.txtbatchNum.Name = "txtbatchNum";
            this.txtbatchNum.Size = new System.Drawing.Size(47, 22);
            this.txtbatchNum.TabIndex = 56;
            this.txtbatchNum.TabStop = false;
            this.txtbatchNum.Text = "1";
            this.txtbatchNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Location = new System.Drawing.Point(9, 134);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(117, 17);
            this.label2.TabIndex = 57;
            this.label2.Text = "Batching Number";
            // 
            // btnRefresh
            // 
            this.btnRefresh.Location = new System.Drawing.Point(334, 130);
            this.btnRefresh.Margin = new System.Windows.Forms.Padding(4);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(75, 24);
            this.btnRefresh.TabIndex = 58;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::Daily_Report.Properties.Resources.LOG4122017122431AM;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(561, 258);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtbatchNum);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtStartDate);
            this.Controls.Add(this.OpenPath2);
            this.Controls.Add(this.TxtSelectExcelFile2);
            this.Controls.Add(this.OutputPath);
            this.Controls.Add(this.txtOutputPath);
            this.Controls.Add(this.OpenPath1);
            this.Controls.Add(this.TxtSelectExcelFile1);
            this.Controls.Add(this.Run_Reports);
            this.DoubleBuffered = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Daily-Report (Agent)";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button OutputPath;
        private System.Windows.Forms.TextBox txtOutputPath;
        private System.Windows.Forms.Button OpenPath1;
        private System.Windows.Forms.TextBox TxtSelectExcelFile1;
        private System.Windows.Forms.Button Run_Reports;
        private System.Windows.Forms.Button OpenPath2;
        private System.Windows.Forms.TextBox TxtSelectExcelFile2;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.TextBox txtStartDate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtbatchNum;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnRefresh;
    }
}

