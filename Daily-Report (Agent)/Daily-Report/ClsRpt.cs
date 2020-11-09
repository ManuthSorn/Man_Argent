using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Reflection;
using System.IO;

namespace Daily_Report
{
    class ClsRpt
    {
        public static string openfilePath = "";
        public static void HeaderRpt1(string path, string Selectpath1, string Selectpath2, string txtStartDate, string batchNum)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Sheets.Add(xlWorkBook.Sheets[1], Type.Missing, Type.Missing, Type.Missing);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            string[] date = txtStartDate.Split('/');
            xlWorkSheet.Name = "Daily Reports (" + date[1].ToString() + "-" + date[0].ToString() + "-" + date[2].ToString() + ")";
            //Daily Reports (28-03-2018)
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;
            //xlApp.Windows.Application.ActiveWindow.DisplayGridlines = false;

            // read excel file 1
            Excel.Application xlApp1 = new Excel.Application();
            Excel.Workbook xlWorkbook1 = xlApp1.Workbooks.Open(Selectpath1);
            Excel._Worksheet xlWorksheet1 = xlWorkbook1.Sheets[1];
            Excel.Range xlRange1 = xlWorksheet1.UsedRange;

            int rowCount1 = xlRange1.Rows.Count;
            //int colCount1 = xlRange1.Columns.Count;

            // read excel file 2
            Excel.Application xlApp2 = new Excel.Application();
            Excel.Workbook xlWorkbook2 = xlApp2.Workbooks.Open(Selectpath2);
            Excel._Worksheet xlWorksheet2 = xlWorkbook2.Sheets[1];
            Excel.Range xlRange2 = xlWorksheet2.UsedRange;

            int rowCount2 = xlRange2.Rows.Count;
            //int colCount2 = xlRange2.Columns.Count;
            //Add Header
            //String[] ABC = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W" };
            String[] Header_Name = { "BU", "Touchpoint", "Agent List Batch Number", "Dummy ID", "Agent Name", "Agent's Number", "Agent's Gender", "Channel", "Owner Name", "Owner's Gender", "Policy Number", "Product Name", "Interview Date", "Interviewer ID", "Call Outcome", "S1.Policy document status", "Q1 Recommend Advisor career", "Q2 Recommend Advisor Career Verbatim", "Q2_Code_01", "Q2_Code_02", "Q2_Code_03", "Q2_Code_04", "Q2_Code_05", "Q2_Code_06", "Q2_Code_07", "Q2_Code_08", "Q2_Code_09", "Q2_Code_10", "Q3 Satisfaction towards NB Process", "Q4 Satisfaction towards NB Process verbatim", "Q4_Code_01", "Q4_Code_02", "Q4_Code_03", "Q4_Code_04", "Q4_Code_05", "Q4_Code_06", "Q4_Code_07", "Q4_Code_08", "Q4_Code_09", "Q4_Code_10", "Q5 Submit via Epp", "Q6 Resubmit any document", "Q7  Document Resubmit-code 01", "Q7  Document Resubmit-code 02", "Q7  Document Resubmit-code 03", "Q7  Document Resubmit-code 04", "Q7  Document Resubmit-code 05", "Q7  Document Resubmit-code 06", "Q7  Document Resubmit-code 07", "Q7  Document Resubmit-code 08", "Q7  Document Resubmit-code 09", "Q7  Document Resubmit-code 10", "Q8 Additional Support Verbatim", "Q8_Code_01", "Q8_Code_02", "Q8_Code_03", "Q8_Code_04", "Q8_Code_05", "Q8_Code_06", "Q8_Code_07", "Q8_Code_08", "Q8_Code_09", "Q8_Code_10", "Q9_Permit to Follow Up", "Q10_Request Manulife to call back", "Daily Flag Report" };

            for (int i = 0; i <= Header_Name.Length - 1; i++)
            {
                xlWorkSheet.Cells[1, i + 1] = Header_Name[i];
                xlWorkSheet.Cells[1, i + 1].HorizontalAlignment = 3;
                xlWorkSheet.Cells[1, i + 1].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[1, i + 1].WrapText = true;
                xlWorkSheet.Cells[1, i + 1].VerticalAlignment = 2;
                if (i >= 3 && i <= 11)
                {
                    xlWorkSheet.Cells[1, i + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                }
                if (i >= 12 && i <= 14)
                {
                    xlWorkSheet.Cells[1, i + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
                }
                if (i == 15 || i == 17 || i == 29 || i == 52)
                {
                    xlWorkSheet.Cells[1, i + 1].Font.Bold = true;
                }

                if (i >= 19 && i <= 28)
                {
                    xlWorkSheet.Cells[1, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                }
                if (i >= 31 && i <= 40)
                {
                    xlWorkSheet.Cells[1, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                }
                if (i >= 54 && i <= 63)
                {
                    xlWorkSheet.Cells[1, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                }
                //if (i == 35)
                //{
                //    xlWorkSheet.Cells[1, i + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 176, 80));
                //    xlWorkSheet.Cells[1, i + 1].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                //}
            }
            xlWorkSheet.Cells[2, 68] = "Red";
            xlWorkSheet.Cells[3, 68] = "Green";
            xlWorkSheet.Cells[4, 68] = "Black";
            xlWorkSheet.Cells[2, 69] = "Q10 = No,Q1 Code 0-4";
            xlWorkSheet.Cells[3, 69] = "Q10 = No,Q1 Code 9 or 10";
            xlWorkSheet.Cells[4, 69] = "Q10= Yes";

            xlWorkSheet.get_Range("BP2", "BP4").Cells.Font.Size = 11;
            xlWorkSheet.get_Range("BP2", "BP4").Cells.Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White);
            xlWorkSheet.get_Range("BP2", "BP2").Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);
            xlWorkSheet.get_Range("BP3", "BP3").Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(0,176,80));
            xlWorkSheet.get_Range("BP4", "BP4").Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
            xlWorkSheet.get_Range("BP2", "BP4").Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            //===========================================
            //Read Data
            int CountData = 0;
            List<int> IDcode = new List<int>();
            List<string> ListDate = new List<string>();
            for (int rowidex = rowCount2; rowidex <= rowCount2; rowidex--)
            {
                if (rowidex == 1) { break; }
                string datevalue = xlRange2.Cells[rowidex, 29].Value2.ToString().Trim();
                double ddate = double.Parse(datevalue);
                DateTime d = DateTime.FromOADate(ddate);
                string getdate = d.ToString("M/d/yyyy");
                if (txtStartDate.ToString() == getdate)
                {
                    CountData += 1;
                    IDcode.Add(rowidex);
                    ListDate.Add(getdate);
                }
            }
            //Column in DB
            //int[] colidex = { 28, 29, 37, 38, 40, 41, 43, 44, 45, 46, 47, 48, 50, 4 };
            int[] colidex = { 28, 29, 4, 37, 38, 39, 41, 42, 44, 45, 46, 47, 48, 49, 50, 51, 53, 55, 57, 59, 60, 61, 63 };
            //Column in Daily-Report
            //int[] reportidex = { 4, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 21 };
            //int[] reportidex = { 4, 13, 14, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35 };
            int[] reportidex = { 4, 13, 14, 16, 17, 18, 29, 30, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 64, 65 };
            int rowCnt = IDcode.Count - 1;
            for (int rowidex = IDcode.Count - 1; rowidex <= IDcode.Count - 1; rowidex--)
            {
                if (rowidex < 0)
                { break; }

                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, 1] = "KH";
                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, 2] = "Agent at NB";
                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, 3] = batchNum;
                if (xlRange2.Cells[IDcode[rowidex], 32].Value2.ToString().Trim() == "99. Other (ផ្សេងៗ​ ទៀត)")
                {
                    xlWorkSheet.Cells[(rowCnt - rowidex) + 2, 15] = xlRange2.Cells[IDcode[rowidex], 33].Value2.ToString().Trim();
                }

                if (xlRange2.Cells[IDcode[rowidex], 32].Value2.ToString().Trim() == "Completed Interview" || xlRange2.Cells[IDcode[rowidex], 32].Value2.ToString().Trim() == "Completed")
                {
                    xlWorkSheet.Cells[(rowCnt - rowidex) + 2, 15] = "Completed"; // Set text: " Completed"  to report.
                }

                else
                {
                    xlWorkSheet.Cells[(rowCnt - rowidex) + 2, 15] = xlRange2.Cells[IDcode[rowidex], 32].Value2.ToString().Trim();
                }
                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, 1].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, 2].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, 3].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, 15].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, 35].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, 36].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                //Get Data

                for (int j = 1; j <= Header_Name.Length; j++) // Set line to report
                {
                    xlWorkSheet.Cells[(rowCnt - rowidex) + 2, j].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    // Original Set Line" in Report"
                    // But not full 
                    //xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                }
  
                for (int i = 0; i <= colidex.Length - 1; i++)
                {
                    if (i == 0)
                    {
                        xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                        for (int getidx = 1; getidx <= rowCount1; getidx++)
                        {
                            if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == xlRange1.Cells[getidx, 1].Value2.ToString().Trim())
                            {

                                for (int colidx = 1; colidx <= 8; colidx++)
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i] + colidx].HorizontalAlignment = 2;
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 2, 6].NumberFormat = "@";
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 2, 11].NumberFormat = "@";
                                    if (colidx >= 1)
                                    {
                                        xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i] + colidx] = xlRange1.Cells[getidx, colidx + 1].Value2.ToString().Trim();
                                    }
                                    else
                                    {
                                        try
                                        {
                                            xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i] + colidx] = xlRange1.Cells[getidx, colidx + 2].Value2.ToString().Trim();
                                        }
                                        catch { }
                                    }
                                    
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i] + colidx].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                                }
                                //xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i] + 1] = xlRange1.Cells[getidx, 2].Value2.ToString().Trim();
                                //xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i] + 2] = xlRange1.Cells[getidx, 3].Value2.ToString().Trim();
                                //xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i] + 1].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                                //xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i] + 2].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                                break;
                            }
                        }
                    }
                    else if (i == 4)
                    {
                        if (xlRange2.Cells[IDcode[rowidex], colidex[3]] != null && xlRange2.Cells[IDcode[rowidex], colidex[3]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[3]].Value2.ToString().Trim() != "")
                        {
                            if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "10 : extremely likely")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = "10";
                            }
                            else if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "0 : not at all likely")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = "0";
                            }
                            else
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                            }
                            xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]].NumberFormat = "0";
                        }
                    }
                    else if (i == 6)
                    {
                        if (xlRange2.Cells[IDcode[rowidex], colidex[3]] != null && xlRange2.Cells[IDcode[rowidex], colidex[3]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[3]].Value2.ToString().Trim() != "")
                        {
                            if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "10 : Very satisfied")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = "10";
                            }
                            else if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "0 : Not satisfied at all")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = "0";
                            }
                            else
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                            }
                            xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]].NumberFormat = "0";
                        }
                    }
                    else if (i >= 10 && i <= 15)
                    {
                        if (xlRange2.Cells[IDcode[rowidex], colidex[3]] != null && xlRange2.Cells[IDcode[rowidex], colidex[3]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[3]].Value2.ToString().Trim() != "")
                        {
                            if (xlRange2.Cells[IDcode[rowidex], colidex[9]].Value2.ToString().Trim() == "Yes")
                            {

                                if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() != "0")
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                                }
                            }
                        }
                    }
                    else if (i >= 16 && i <= 19)
                    {
                        if (xlRange2.Cells[IDcode[rowidex], colidex[3]] != null && xlRange2.Cells[IDcode[rowidex], colidex[3]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[3]].Value2.ToString().Trim() != "")
                        {
                            if (xlRange2.Cells[IDcode[rowidex], colidex[9]].Value2.ToString().Trim() == "Yes")
                            {

                                if (xlRange2.Cells[IDcode[rowidex], colidex[i] - 1].Value2.ToString().Trim() != "0")
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                                }
                            }
                        }
                    }
                    else if (i == 2)
                    {
                        xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                        xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]].HorizontalAlignment = 3;
                        xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    }
                    else
                    {
                        if (xlRange2.Cells[IDcode[rowidex], colidex[3]] != null && xlRange2.Cells[IDcode[rowidex], colidex[3]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[3]].Value2.ToString().Trim() != "")
                        {
                            if (xlRange2.Cells[IDcode[rowidex], colidex[i]] != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() != "")
                            {
                                if (i == 1)
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = ListDate[rowidex];
                                }
                                else
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                                }
                            }
                            else
                            {
                                if (i == 5)
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i] + 1].Value2.ToString().Trim();
                                }
                                else if (i == 7)
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i] + 1].Value2.ToString().Trim();
                                }
                                else
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = "";
                                }
                            }
                        }
                        else
                        {
                            if (i == 1)
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 2, reportidex[i]] = ListDate[rowidex];
                            }
                        }
                    }
                }
            }
            xlWorkSheet.Range["B:B"].ColumnWidth = 13.00;
            xlWorkSheet.Range["E:E"].ColumnWidth = 17.00;
            xlWorkSheet.Range["F:F"].ColumnWidth = 31.00;
            xlWorkSheet.Range["G:H"].ColumnWidth = 7.00;
            xlWorkSheet.Range["I:I"].ColumnWidth = 15.00;
            xlWorkSheet.Range["J:J"].ColumnWidth = 7.00;
            xlWorkSheet.Range["K:K"].ColumnWidth = 15.00;
            xlWorkSheet.Range["L:L"].ColumnWidth = 37.00;
            xlWorkSheet.Range["N:N"].ColumnWidth = 10.00;
            xlWorkSheet.Range["M:M"].ColumnWidth = 10.00;
            xlWorkSheet.Range["P:P"].Columns.AutoFit();
            xlWorkSheet.Range["R:R"].ColumnWidth = 10.00;
            xlWorkSheet.Range["T:T"].Columns.AutoFit();
            xlWorkSheet.Range["AG:AG"].Columns.AutoFit();
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange1);
            Marshal.ReleaseComObject(xlWorksheet1);
            Marshal.ReleaseComObject(xlRange2);
            Marshal.ReleaseComObject(xlWorksheet2);

            //close and release
            xlWorkbook1.Close();
            Marshal.ReleaseComObject(xlWorkbook1);
            xlWorkbook2.Close();
            Marshal.ReleaseComObject(xlWorkbook2);

            //quit and release
            xlApp1.Quit();
            Marshal.ReleaseComObject(xlApp1);
            xlApp2.Quit();
            Marshal.ReleaseComObject(xlApp2);
            try
            {
                //xlWorkBook.CheckCompatibility = false;
                xlApp.DisplayAlerts = false;
                //xlWorkBook.DoNotPromptForConvert = true;
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
                xlWorkBook.SaveAs(path, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            openfilePath = path;
            //MessageBox.Show("Daily-Report has been successful!!!.");
        }
    }
}
