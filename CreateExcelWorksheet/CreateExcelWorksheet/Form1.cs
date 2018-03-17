using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection; 
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace CreateExcelWorksheet
{
    public partial class FormExcel : Form
    {
        public FormExcel()
        {
            InitializeComponent();
        }

        private void createFile_Click(object sender, EventArgs e)
        {
            RESET_Click(sender, e);
            createFile.BackColor = System.Drawing.Color.Chartreuse;
            MessageBox.Show("button1 , \nThis will create excel file  : d:\\csharp-Excel.xls");
            Microsoft.Office.Interop.Excel.Application xlApp = new
            Microsoft.Office.Interop.Excel.Application();
            if (xlApp == null) 
            { 
                MessageBox.Show("Excel is not properly installed!!"); 
                return;
            }
                MessageBox.Show("Excel is installed on this computer");


            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;

          




            object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue); //create new Workbook
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1); // Get the worksheet number 1
                xlWorkSheet.Name = "sheet1";

                xlWorkSheet.Cells[1, 1] = "cell 1 1";                        //write content to worksheet
                xlWorkSheet.Cells[1, 2] = "cell 1 2";
                xlWorkSheet.Cells[2, 1] = "cell 2 1";
                xlWorkSheet.Cells[2, 2] = "cell 2 2";


                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.Add(System.Reflection.Missing.Value,xlWorkBook.Worksheets[xlWorkBook.Worksheets.Count],System.Reflection.Missing.Value,System.Reflection.Missing.Value); // Create new worksheet
                xlWorkSheet.Name = "sheet2";

                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.Add(System.Reflection.Missing.Value, xlWorkBook.Worksheets[xlWorkBook.Worksheets.Count], System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                xlWorkSheet.Name = "sheet3";



                xlWorkBook.SaveAs("d:\\csharp-Excel.xls",
                    Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue); //save the excel file.
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                MessageBox.Show("Excel file created , you can find the file d:\\csharp-Excel.xls");
            
        }

        private void createRange_Click(object sender, EventArgs e)
        {
            //This will create a Excel app and write on this 

            RESET_Click(sender, e);
            createRange.BackColor = System.Drawing.Color.Chartreuse;
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    MessageBox.Show("EXCEL could not be started. Check that your office installation and project references are correct.");
                    return;
                }
                xlApp.Visible = true;

                Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet ws = (Worksheet)wb.Worksheets[1];
                ws.Name = "page1";



            if (ws == null)
            {
                MessageBox.Show("Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            else
            {
                // Select the Excel cells, in the range c1 to c7 in the worksheet.
                Range aRange = ws.get_Range("C1", "C7");

                if (aRange == null)
                {
                    MessageBox.Show("Could not get a range. Check to be sure you have the correct versions of the office DLLs.");
                }

                // Fill the cells in the C1 to C7 range of the worksheet with the number 6.
                Object[] args = new Object[1];
                args[0] = 6;

                aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);

                // Change the cells in the C1 to C7 range of the worksheet to the number 8.
                aRange.Value2 = 8;



                aRange = ws.get_Range("A1", "B7");
                args[0] = 6;
                aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);
                aRange.Value2 = 5;

                // Create workshit page2
                Excel.Worksheet addedSheet = wb.Worksheets.Add(Type.Missing,wb.Worksheets[1], Type.Missing, Type.Missing);
                addedSheet.Name = "page2";

                addedSheet.Cells[2, 3] = "here is rox2 - column3";

                // Create workshit page3
                int count = wb.Worksheets.Count;
                addedSheet = wb.Worksheets.Add(Type.Missing, wb.Worksheets[count], Type.Missing, Type.Missing);
                addedSheet.Name = "page3";
                ws = (Worksheet)wb.Worksheets[3];
                aRange = ws.get_Range("B1", "B7");
                args[0] = 6;
                aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);
                aRange.Value2 = "page3";

                xlApp.ActiveWorkbook.Sheets[1].Select(); //activates first sheet
            }

            //wb.Close(true, null, null);
            //xlApp.Quit();
            releaseObject(ws);
            releaseObject(wb);
            releaseObject(xlApp);

        }



        private void readExcel_Click(object sender, EventArgs e)
        {
            RESET_Click(sender, e);
            readExcel.BackColor = System.Drawing.Color.Chartreuse;
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open("d:\\csharp-Excel.xls", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Range aRange = xlWorkSheet.get_Range("A1", "A2");

            if (aRange == null)
            {
                MessageBox.Show("Could not get a range. Check to be sure you have the correct versions of the office DLLs.");
            }
            else
            {
                MessageBox.Show(xlWorkSheet.get_Range("A1", "A1").Value2.ToString());
                MessageBox.Show(xlWorkSheet.get_Range("A2", "A2").Value2.ToString());

            }

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void readAll_Click(object sender, EventArgs e)
        {
            RESET_Click(sender, e);
            readAll.BackColor = System.Drawing.Color.Chartreuse;

            Microsoft.Office.Interop.Excel.Application xlRead;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range range;

            string str;
            int rCnt = 0;
            int cCnt = 0;

            xlRead = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlRead.Workbooks.Open("d:\\csharp-Excel.xls", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0); //WorkBooks.open(string Filename, object UpdateLinks, object ReadOnly, object Format, object Password, object WriteResPassword, object ReadOnlyRecommend, object Origin, object Delimiter, object Editable, object Notify, object Converter, object AddToMru, object Local, object CorruptLoad )
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {
                for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                    MessageBox.Show("Worksheet name : "  + xlWorkSheet.Name  + "\n" +  str);
                }
            }

            xlWorkBook.Close(true, null, null);
            xlRead.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlRead);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }

            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }

            finally
            {
                GC.Collect();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

            RESET_Click(sender, e);

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook g_Workbook = excelApp.Workbooks.Open(
                @"d:\\csharp-Excel.xls",
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            for (int i = 0; i < 5; i++)
            {
                int count = g_Workbook.Worksheets.Count;
                MessageBox.Show("" + count);
                Excel.Worksheet addedSheet = g_Workbook.Worksheets.Add(Type.Missing,g_Workbook.Worksheets[count], Type.Missing, Type.Missing);
                addedSheet.Name = i.ToString();
            }
        }

        private void RESET_Click(object sender, EventArgs e)
        {
            createFile.BackColor = System.Drawing.SystemColors.Control;
            createRange.BackColor = System.Drawing.SystemColors.Control;
            readExcel.BackColor = System.Drawing.SystemColors.Control;
            readAll.BackColor = System.Drawing.SystemColors.Control;
        }
    }
}
