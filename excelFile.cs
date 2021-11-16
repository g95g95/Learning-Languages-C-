using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace NewWords
{
    public class fileExcel <typeofColumn>
    {
        private string filepath;
        private bool Visible = false;
        private string sheetName = "Sheet1";
        private static int excelCounter = 0;


        // Getters & Setters section
        public int getCounter()
        {
            return excelCounter;
        }

        public string getFilepath()
        {
            return filepath;
        }

        public bool getVisible()
        {
            return Visible;
        }

        public string getSheetname()
        {
            return sheetName;
        }


        public string setFilepath(string newfilepath)
        {
            return filepath = newfilepath;
        }

        public bool setVisible(bool newVisible)
        {
            return Visible = newVisible;
        }

        public string setSheetname(string newsheetName)
        {
            return sheetName = newsheetName;
        }

        // Default Constructor
        public fileExcel()
        {
            excelCounter++;
        }

        //Constructor
        public fileExcel(string filepath, bool Visible, string sheetName = "Sheet1")
        {
            excelCounter++;
            this.Visible = Visible;
            this.filepath = filepath;
            this.sheetName = sheetName;
        }
        // This function returns a Workbook
        public Excel.Workbook returnWorkbook()
        {
            try
            {
                Excel.Application xlsApp = new Excel.Application();
                xlsApp.Visible = this.Visible;
                Excel.Workbook xlsWorkbook = xlsApp.Workbooks.Open(this.filepath);
                //xlsApp.Quit();

                return xlsWorkbook;
 

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
                var excel = new Excel.Application();
                var workbooks = excel.Workbooks;
                var workbook = workbooks.Add();
                var worksheet = (Excel.Worksheet)excel.ActiveSheet;

                workbook.SaveAs2(filepath);
                //excel.Quit();

                return workbook;



            }

        }
        // This function returns a Sheet!
        public Excel.Worksheet returnWorksheet()
        {
            try
            {
                Excel.Application xlsApp = new Excel.Application();
                xlsApp.Visible = this.Visible;
                Excel.Workbook xlsWorkbook = xlsApp.Workbooks.Open(this.filepath);
                Excel.Sheets xlsSheets = xlsWorkbook.Worksheets;
                Excel.Worksheet xlsSheet = (Excel.Worksheet)xlsSheets.get_Item("Sheet1");
                //xlsApp.Quit();

                return xlsSheet;


            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
                var excel = new Excel.Application();
                var workbooks = excel.Workbooks;
                var workbook = workbooks.Add();
                var worksheet = (Excel.Worksheet)excel.ActiveSheet;

                workbook.SaveAs2(filepath);
                //excel.Quit();
                return worksheet;
 

            }
        }
        static public typeofColumn[] getColumnsbyHeader(string Header, Excel.Worksheet xlsSheet)
        {
            int rowsTot = (int)xlsSheet.UsedRange.Rows.Count;
            int colsTot = (int)xlsSheet.UsedRange.Columns.Count;
            int indexHeader = 0;
            typeofColumn[] column = new typeofColumn[rowsTot-1];


            for (int i = 1; i <= colsTot; ++i)
            {
                if (Convert.ToString(((Excel.Range)xlsSheet.Cells[1, i]).Value2) == Header)
                {
                    indexHeader = i;
                }
            }

            for (int i = 0; i < (rowsTot -1); ++i)
            {
                column[i] = (typeofColumn)(((Excel.Range)xlsSheet.Cells[i + 2, indexHeader]).Value2);
            }


            return column;
        }

        static public typeofColumn[] getColumnbyExcelIndex(string excelIndex, Excel.Worksheet xlsSheet)
        {
            int rowsTot = (int)xlsSheet.UsedRange.Rows.Count;
            int colsTot = (int)xlsSheet.UsedRange.Columns.Count;
            typeofColumn[] column = new typeofColumn[rowsTot - 1];

            for (int i = 0; i < (rowsTot -1); ++i)
            {
                column[i] = (typeofColumn)((xlsSheet.Range[excelIndex + Convert.ToString(i + 2)]).Value2);
            }


            return column;

        }
    }

}
