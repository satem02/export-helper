using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExportHelper.ProcessReport;
using ExportHelper.Interfaces;
using System.Data;
using System.Collections.Generic;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        IReport excelReport = new ExcelReport();
        string filePath = "C:\\ExampleExcelFileName";

        [TestMethod]
        public void TestMethod1()
        {
            List<Customer> input = GetDummyList();

            excelReport.FilePath = filePath;
            excelReport.Export(input);
        }

        [TestMethod]
        public void TestMethod2()
        {
            DataTable input = GetDummyDataTable();

            excelReport.FilePath = filePath;
            excelReport.Export(input);
        }

        private DataTable GetDummyDataTable()
        {
            var dt = new DataTable();
            // Fill dt
            return dt;
        }

        private List<Customer> GetDummyList()
        {
            var list = new List<Customer>();


            // Fill list
            Customer cs = new Customer 
            {
                Id = 1,
                Name = "safak",
                Surname = "temel"
            };
            list.Add(cs);

            return list;
        }
    }

    public class Customer
    {
        [ExportHelper.Attributes.IgnoreAttribute]
        public int Id { get; set; }

        public string Name { get; set; }

        public string Surname { get; set; }
    }
}
