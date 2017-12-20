using System;
using System.Data;
using MyExcel = Microsoft.Office.Interop.Excel;

namespace ExportHelper
{
    public static class ExcelHelper
    {
        public static void Export(DataTable dt, string filePath)
        {
            MyExcel.Application oexcel = new MyExcel.Application();
            try
            {
                string path = AppDomain.CurrentDomain.BaseDirectory;
                object misValue = System.Reflection.Missing.Value;
                MyExcel.Workbook obook = oexcel.Workbooks.Add(misValue);
                MyExcel.Worksheet osheet = new MyExcel.Worksheet();

                osheet = (MyExcel.Worksheet)obook.Sheets["Sheet1"];
                int colIndex = 0;
                int rowIndex = 1;

                foreach (DataColumn dc in dt.Columns)
                {
                    colIndex++;
                    osheet.Cells[1, colIndex] = dc.ColumnName;
                }
                foreach (DataRow dr in dt.Rows)
                {
                    rowIndex++;
                    colIndex = 0;

                    foreach (DataColumn dc in dt.Columns)
                    {
                        colIndex++;
                        osheet.Cells[rowIndex, colIndex] = dr[dc.ColumnName];
                    }
                }

                osheet.Columns.AutoFit();

                obook.SaveAs(filePath);
                obook.Close();
                oexcel.Quit();

                GC.Collect();
            }
            catch (Exception ex)
            {
                oexcel.Quit();
            }
        }
    }
}
