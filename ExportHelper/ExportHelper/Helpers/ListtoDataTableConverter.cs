using ExportHelper.Attributes;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;

namespace ExportHelper.Helpers
{
    public static class ListtoDataTableConverter
    {

        public static DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            //Get all the properties
            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            PropertyInfo[] fixProps = FixProbs(props);


            foreach (PropertyInfo prop in fixProps)
            {
                dataTable.Columns.Add(prop.Name);
            }

            foreach (T item in items)
            {
                var values = new object[fixProps.Length];
                for (int i = 0; i < fixProps.Length; i++)
                {
                    values[i] = fixProps[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }

            return dataTable;
        }

        private static PropertyInfo[] FixProbs(PropertyInfo[] props)
        {
            var response = props
                        .Where(s =>
                            s.CustomAttributes
                                .Where(a => a.AttributeType.Equals(typeof(IgnoreAttribute)))
                                .ToList().Count == 0)
                        .ToArray();

            return response;
        }
    }
}
