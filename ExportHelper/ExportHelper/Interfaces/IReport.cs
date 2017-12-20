using System.Collections.Generic;
using System.Data;

namespace ExportHelper.Interfaces
{
    public interface IReport
    {
        string FilePath { set; }
        void Export<T>(List<T> items);
        void Export(DataTable items);
    }
}
