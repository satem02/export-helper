using ExportHelper.Helpers;
using ExportHelper.Interfaces;
using System.Collections.Generic;
using System.Data;

namespace ExportHelper.ProcessReport
{
    public class ExcelReport : IReport
    {
        string _filePath;
        public string FilePath
        {
            set { _filePath = value; }
        }

        public void Export<T>(List<T> items)
        {
            var dt = ListtoDataTableConverter.ToDataTable(items);
            ExcelHelper.Export(dt, _filePath);
        }

        public void Export(DataTable items)
        {
            ExcelHelper.Export(items, _filePath);
        }
    }
}
