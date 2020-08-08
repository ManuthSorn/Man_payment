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
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;
           

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
        
            String[] Header_Name = { "BU", "Touchpoint", "Customer List Batch Number", "Dummy ID", "Client Number", "Owner Name", "Product Name", "Name of Rider", "Channel", "Interview Date", "Interviewer No", "Call Outcome", "Q1 Renewal Payment tNPS", "Renewal Payment tNPS Group", "Q2/Q3 Renewal Payment tNPS Verbatim", "Q2_Code_01", "Q2_Code_02", "Q2_Code_03", "Q2_Code_04", "Q2_Code_05", "Q2_Code_06", "Q2_Code_07", "Q2_Code_08", "Q2_Code_09", "Q2_Code_10", "Q4 Satisfaction with payment process", "Q5 Renewal Payment Method - Raw", "Q5 Renewal Payment - Who", "Q5.1 Renewal Payment - How", "Q5.2 Renewal Payment - Where", "Q6. Receive any premium notication", "Q6. Receive Phone call from Manulife", "Q6. Receive Phone call from IA/IS", "Q6. Receive Email from Manulife", "Q6. Receive Email from IA/IS", "Q6. Receive SMS from Manulife", "Q6. Receive SMS from IA/IS", "Q6. Receive WahtsApp from IA/IS", "Q6. Receive FB Messenger from my IA/IS", "Q6. Receive Other Noticiation", "Q6. Receive Other Noticiation (specified)", "Q7. Last policy review with IA/IS", "Q8. Satisfaction towards Annual Portfolio Review", "Q9.Area for Improvement verbatim", "Q9_Code_01", "Q9_Code_02", "Q9_Code_03", "Q9_Code_04", "Q9_Code_05", "Q9_Code_06", "Q9_Code_07", "Q9_Code_08", "Q9_Code_09", "Q9_Code_10", "Q10. Permit to Follow Up", "Q11_Request Manulife to call back", "Daily Flag Report", "AGE", "Age Group", "CUSTOMER_CATE", "NUMBER_OF_POLICYS", "ANNUAL_INCOME", "Income Group ", "PRODUCT_CATEGORY", "APE Group", "PREMIUM_PAYMENT_MOD", "CUSTOMER_TENTURE_MANULIFE", "Tenure Group", "PAYMENT_METHOD", "AGENT_ID", "AM_NM", "AGENT_LOCATION", "ORPHAN_STATUS", "UNDERWRITER_ID", "PROCESSING_TIME", "SUBMISSION_DATE", "DATA_ENTRY_DATE", "CUSTOMER_CONFIRMED_DATE", "UW_REASULT", "TRXN_DT", "TRXN_AMT", "PAID_TO_DATE", "Re-PTD" };

            for (int i = 0; i <= Header_Name.Length - 1; i++)
            {
                
                xlWorkSheet.Cells[1, i + 1].HorizontalAlignment = 3;
                xlWorkSheet.Cells[1, i + 1].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                xlWorkSheet.Cells[3, i + 1] = Header_Name[i];
                xlWorkSheet.Cells[3, i + 1].HorizontalAlignment = 3;
                xlWorkSheet.Cells[3, i + 1].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[3, i + 1].WrapText = true;
                xlWorkSheet.Cells[3, i + 1].VerticalAlignment = 2;

                // set text bold in report 
                // 3 is row 
                // i is culom
                // +1 becouse if i set == the real culom is not work in the end culom, so i had to set real culom -1 and +1 when set culom bold
                if (i == 14 || i == 43 || i >= 57)
                {
                    xlWorkSheet.Cells[3, i + 1].Font.Bold = true;
                }
            }

            xlWorkSheet.Range[xlWorkSheet.Cells[2, 1], xlWorkSheet.Cells[2, 3]].Merge();
            xlWorkSheet.Cells[2, 1] = "IndoChina to create";
            xlWorkSheet.Cells[2, 1].HorizontalAlignment = 3;
            //xlWorkSheet.Cells[2, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
            xlWorkSheet.Cells[2, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(146, 208, 80));
            for (int i = 1; i <= 3; i++)
            {
                //xlWorkSheet.Cells[3, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                xlWorkSheet.Cells[2, i].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            }

            xlWorkSheet.Range[xlWorkSheet.Cells[2, 4], xlWorkSheet.Cells[2, 9]].Merge();
            xlWorkSheet.Cells[2, 4] = "From Customer Data Set";
            xlWorkSheet.Cells[2, 4].HorizontalAlignment = 3;
            xlWorkSheet.Cells[2, 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
            for (int i = 4; i <= 9; i++)
            {
                xlWorkSheet.Cells[3, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                xlWorkSheet.Cells[2, i].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            }

            xlWorkSheet.Range[xlWorkSheet.Cells[2, 10], xlWorkSheet.Cells[2, 12]].Merge();
            xlWorkSheet.Cells[2, 10] = "Official use";
            xlWorkSheet.Cells[2, 10].HorizontalAlignment = 3;
            xlWorkSheet.Cells[2, 10].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
            for (int i = 10; i <= 12; i++)
            {
                xlWorkSheet.Cells[3, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
                xlWorkSheet.Cells[2, i].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            }

            xlWorkSheet.Range[xlWorkSheet.Cells[2, 13], xlWorkSheet.Cells[2, 57]].Merge();
            xlWorkSheet.Cells[2, 13] = "tNPS Survey (response and coded Oes)";
            xlWorkSheet.Cells[2, 13].HorizontalAlignment = 3;
            //xlWorkSheet.Cells[2, 9].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
            xlWorkSheet.Cells[2, 13].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(146, 208, 80));
            for (int i = 16; i <= 25; i++)
            {
                xlWorkSheet.Cells[3, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 192, 0));
            }
            for (int i = 45; i <= 54; i++)
            {
                xlWorkSheet.Cells[3, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 192, 0));
            }


            // line in report
            for (int i = 13; i <= 57; i++)
            {
                //xlWorkSheet.Cells[3, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                xlWorkSheet.Cells[2, i].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            }

            xlWorkSheet.Range[xlWorkSheet.Cells[2, 58], xlWorkSheet.Cells[2, 83]].Merge();
            xlWorkSheet.Cells[2, 58] = "Official use";
            xlWorkSheet.Cells[2, 58].HorizontalAlignment = 3;
            xlWorkSheet.Cells[2, 58].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
            for (int i = 58; i <= 83; i++)
            {
                //xlWorkSheet.Cells[3, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGray);
                xlWorkSheet.Cells[2, i].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            }

            xlWorkSheet.Cells[4, 85] = "Red";
            xlWorkSheet.Cells[5, 85] = "Green";
            xlWorkSheet.Cells[6, 85] = "Black";
            xlWorkSheet.Cells[4, 86] = "Q11= No, Q1 code 0-4";
            xlWorkSheet.Cells[5, 86] = "Q11= No, Q1 code 9 or 10";
            xlWorkSheet.Cells[6, 86] = "Q11= Yes";

            xlWorkSheet.get_Range("CG4", "CG6").Cells.Font.Size = 11;
            xlWorkSheet.get_Range("CG4", "CG6").Cells.Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White);
            xlWorkSheet.get_Range("CG4", "CG4").Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Red);
            xlWorkSheet.get_Range("CG5", "CG5").Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(0, 176, 80));
            xlWorkSheet.get_Range("CG6", "CG6").Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Black);
            xlWorkSheet.get_Range("CG4", "CG6").Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

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
            int[] colidex = { 28, 29, 4, 33, 35, 36, 38, 39, 40, 42, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 55, 56, 57, 58, 60 };
            //Column in Daily-Report
            //int[] reportidex = { 4, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 21 };
            //int[] reportidex = { 4, 8, 9, 10, 11, 13, 24, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 40, 41, 42, 53, 54 };
            int[] reportidex = { 4, 10, 11, 12, 13, 15, 26, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 42, 43, 44, 55, 56 };
            int rowCnt = IDcode.Count - 1;
            for (int rowidex = IDcode.Count - 1; rowidex <= IDcode.Count - 1; rowidex--)
            {
                if (rowidex < 0)
                { break; }

                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 1] = "KH";
                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 2] = "Payment";
                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 3] = batchNum;
                //xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 8] = "Completed";
                if (xlRange2.Cells[IDcode[rowidex], 32].Value2.ToString().Trim() == "99. Other (ផ្សេងៗ​ ទៀត)") //99. Other (ផ្សេងៗ​ ទៀត)
                {
                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 12] = xlRange2.Cells[IDcode[rowidex], 33].Value2.ToString().Trim();
                }
                else
                {
                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 12] = xlRange2.Cells[IDcode[rowidex], 32].Value2.ToString().Trim();
                }
                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 1].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 2].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 3].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 8].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                string[] IDout_come = xlRange2.Cells[IDcode[rowidex], 32].Value2.ToString().Trim().Split('.');
                //Get Data
                for (int i = 0; i <= colidex.Length - 1; i++)
                {
                    for (int colbd = 1; colbd <= 83; colbd++)
                    {
                        xlWorkSheet.Cells[(rowCnt - rowidex) + 4, colbd].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    }
                    //xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    //xlWorkSheet.Cells[(rowCnt - rowidex) + 4, 20].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                    //xlWorkSheet.get_Range("A" + (rowCnt - rowidex), "AN" + reportidex[i]).Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    if (i == 0)
                    {
                        xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                        for (int getidx = 1; getidx <= rowCount1; getidx++)
                        {
                            if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == xlRange1.Cells[getidx, 1].Value2.ToString().Trim())
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 1].NumberFormat = "@";
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 1] = xlRange1.Cells[getidx, 2].Value2.ToString().Trim();
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 2] = xlRange1.Cells[getidx, 3].Value2.ToString().Trim();
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 3] = xlRange1.Cells[getidx, 4].Value2.ToString().Trim();
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 4].NumberFormat = "@";
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 4] = xlRange1.Cells[getidx, 5].Value2.ToString().Trim();
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 5] = xlRange1.Cells[getidx, 6].Value2.ToString().Trim();
                                //try { xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 6] = xlRange1.Cells[getidx, 7].Value2.ToString().Trim(); }//=========error
                                //catch { }
                                //xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 7] = xlRange1.Cells[getidx, 8].Value2.ToString().Trim();
                                //xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 8] = xlRange1.Cells[getidx, 9].Value2.ToString().Trim();
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 1].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 2].BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                                break;
                            }
                        }



                    }
                    else if (i == 4)
                    {
                        if (xlRange2.Cells[IDcode[rowidex], colidex[i]] != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() != "")
                        {
                            int q1 = 0;
                            if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "10. Extremely likely")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "10";
                                q1 = 10;
                            }
                            else if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "0. Not at all likely")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "0"; q1 = 0;
                            }
                            else
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                                q1 = Convert.ToInt32(xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim());
                            }

                            // Set "Renewal Payment tNPS Group" in rpt (Dayily report)
                            xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]].NumberFormat = "0";
                            if (q1 >= 9)
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 1] = "Promoter";
                            }
                            else if (q1 <= 6)
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 1] = "Detractor";
                            }
                            else { xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 1] = "Passive"; }

                        }
                    }
                    else if (i == 5)
                    {
                        if (xlRange2.Cells[IDcode[rowidex], colidex[4]] != null && xlRange2.Cells[IDcode[rowidex], colidex[4]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[4]].Value2.ToString().Trim() != "")
                        {
                            if (xlRange2.Cells[IDcode[rowidex], colidex[i]] != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() != "")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                            }
                            else
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i] + 1].Value2.ToString().Trim();
                            }
                        }
                    }
                    else if (i == 6)
                    { // "Q4 Satifaction with payment process"
                        if (xlRange2.Cells[IDcode[rowidex], colidex[i]] != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() != "")
                        {
                            if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "10. Very satisfied")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "10";
                            }
                            else if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "0. Not satisfied at all")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "0";
                            }
                            else
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                            }
                            xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]].NumberFormat = "0";
                        }
                    }
                    else if (i == 21)
                    {
                        if (xlRange2.Cells[IDcode[rowidex], colidex[i]] != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() != "")
                        {
                            if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "10. Very satisfied")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "10";
                            }
                            else if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "0. Not satisfied at all")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "0";
                            }
                            else
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                            }
                            xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]].NumberFormat = "0";
                        }
                    }
                    else if (i == 8 || i == 9)
                    {
                        if (xlRange2.Cells[IDcode[rowidex], colidex[i]] != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() != "")
                        {
                            if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "Others (Specified)")
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i] + 1].Value2.ToString().Trim();
                            }
                            else
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim();
                            }
                        }
                    }
                    else if (i >= 10 && i <= 19)
                    {
                        if (xlRange2.Cells[IDcode[rowidex], colidex[i]] != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() != "")
                        {
                            if (i == 19)
                            {
                                if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "Others (Specified)")
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "Yes";
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + 1] = xlRange2.Cells[IDcode[rowidex], colidex[i] + 1].Value2.ToString().Trim();
                                }
                                else
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "No";
                                }
                            }
                            else
                            {
                                if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() != "0")
                                {
                                    if (i == 10)
                                    {
                                        if (xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() == "No")
                                        {
                                            for (int n = 0; n <= 9; n++)
                                            {
                                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i] + n] = "No";
                                            }
                                        }
                                        else
                                        {
                                            xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "Yes";
                                        }
                                    }
                                    else
                                    {
                                        xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "Yes";
                                    }
                                }
                                else
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = "No";
                                }
                            }
                        }
                    }
                    else
                    {
                        if (xlRange2.Cells[IDcode[rowidex], colidex[i]] != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() != "")
                        {
                            if (xlRange2.Cells[IDcode[rowidex], colidex[i]] != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2 != null && xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim() != "")
                            {
                                if (i == 1)
                                {
                                    xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = ListDate[rowidex];
                                }
                                else { xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = xlRange2.Cells[IDcode[rowidex], colidex[i]].Value2.ToString().Trim(); }
                            }
                        }
                        else
                        {
                            if (i == 1)
                            {
                                xlWorkSheet.Cells[(rowCnt - rowidex) + 4, reportidex[i]] = ListDate[rowidex];
                            }
                        }
                    }
                }
            }





            xlWorkSheet.Range["B:B"].ColumnWidth = 21.00;
            xlWorkSheet.Range["E:L"].ColumnWidth = 20.00;
            xlWorkSheet.Range["U:U"].ColumnWidth = 32.00;
            xlWorkSheet.Range["AI:AI"].ColumnWidth = 17.00;
            xlWorkSheet.Range["R:R"].Columns.AutoFit();
            xlWorkSheet.Range["U:W"].Columns.AutoFit();
            xlWorkSheet.Range["AH:AH"].Columns.AutoFit();
            xlWorkSheet.Range["AK:AK"].Columns.AutoFit();
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
