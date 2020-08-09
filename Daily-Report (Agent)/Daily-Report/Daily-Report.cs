using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Globalization;

namespace Daily_Report
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        private void OpenPath1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Execl file 2013 |*.xlsx";
            openFileDialog1.Title = "Import of the file to be exported";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (openFileDialog1.FileName == "")
                {
                    MessageBox.Show("Please select path file....");
                }
                else
                {
                    TxtSelectExcelFile1.Text = openFileDialog1.FileName.ToString();
                    string Batch = System.IO.Path.GetFileNameWithoutExtension(openFileDialog1.FileName);
                    string[] getBatch = Batch.Split(' ');
                    txtbatchNum.Text = getBatch[1];
                    if (File.ReadAllText("../../PathFile.txt") != TxtSelectExcelFile1.Text) 
                    {
                        File.WriteAllText("../../PathFile.txt", TxtSelectExcelFile1.Text);
                    }
                }
            }
            Run_Reports.Select();
        }

        private void OpenPath2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Execl file 2013 |*.xlsx";
            openFileDialog1.Title = "Import of the file to be exported";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (openFileDialog1.FileName == "")
                {
                    MessageBox.Show("Please select path file....");
                }
                else
                {
                    TxtSelectExcelFile2.Text = openFileDialog1.FileName.ToString();
                }
            }
            Run_Reports.Select();
        }

        private void OutputPath_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Execl file |*.xls";
            saveFileDialog1.Title = "Save path of the file to be exported";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (saveFileDialog1.FileName == "")
                {
                    MessageBox.Show("Please select output path file....");
                }
                else
                {
                    txtOutputPath.Text = saveFileDialog1.FileName.ToString();
                }
            }
            Run_Reports.Select();
        }
 
        private void Run_Reports_Click(object sender, EventArgs e)
        {
            if (txtOutputPath.Text == "")
            {
                MessageBox.Show("Please select output path file....");
            }
            else
            { 
                ClsRpt.HeaderRpt1(txtOutputPath.Text.ToString(), TxtSelectExcelFile1.Text.ToString(), TxtSelectExcelFile2.Text.ToString(),txtStartDate.Text.Trim().ToString(),txtbatchNum.Text.Trim().ToString());
                if (MessageBox.Show("Daily-Report has been successful!!!." + "\n" + "Do you want to open file?", "Open File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Process.Start(ClsRpt.openfilePath);
                    Application.Exit();
                }
                else { Application.Exit(); }
                
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DateTime dt = DateTime.Now;
            string filldate = String.Format("{0:M/d/yyyy}", dt);
            string[] splitDate = filldate.Split('/');
            int day = Convert.ToInt32(splitDate[1]) - 1;
            txtStartDate.Text = splitDate[0] + "/" + day.ToString() + "/" + splitDate[2];

            if (File.ReadAllText("../../PathFile.txt") != "") 
            {
                TxtSelectExcelFile1.Text = File.ReadAllText("../../PathFile.txt");
                string Batch = System.IO.Path.GetFileNameWithoutExtension(File.ReadAllText("../../PathFile.txt"));
                string[] getBatch = Batch.Split(' ');
                txtbatchNum.Text = getBatch[1];

                string getPathOutput = Path.GetDirectoryName(File.ReadAllText("../../PathFile.txt"));
                if (day < 10)
                {
                    txtOutputPath.Text = getPathOutput + "\\Rpt-0" + day.ToString() + "-" + dt.ToString("MMM") + "-" + dt.ToString("yyyy") + " (Agent).xls";
                }
                else 
                {
                    txtOutputPath.Text = getPathOutput + "\\Rpt-" + day.ToString() + "-" + dt.ToString("MMM") + "-" + dt.ToString("yyyy") + " (Agent).xls";
                }
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            DateTime dt = DateTime.ParseExact(txtStartDate.Text, "M/d/yyyy", CultureInfo.InvariantCulture);
            string filldate = String.Format("{0:M/d/yyyy}", dt);
            string[] splitDate = filldate.Split('/');
            int day = Convert.ToInt32(splitDate[1]);
            string getPathOutput = Path.GetDirectoryName(File.ReadAllText("../../PathFile.txt"));
            if (day < 10)
            {
                txtOutputPath.Text = getPathOutput + "\\Rpt-0" + day.ToString() + "-" + dt.ToString("MMM") + "-" + dt.ToString("yyyy") + " (Agent).xls";
            }
            else
            {
                txtOutputPath.Text = getPathOutput + "\\Rpt-" + day.ToString() + "-" + dt.ToString("MMM") + "-" + dt.ToString("yyyy") + " (Agent).xls";
            }
        }
    }
}
