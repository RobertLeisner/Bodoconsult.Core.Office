using System;
using System.Diagnostics;
using System.Runtime.Versioning;
using Bodoconsult.Core.Database;
using NUnit.Framework;

namespace Bodoconsult.Core.Office.Tests
{
    [TestFixture]
    [SupportedOSPlatform("windows")]
    public class UnitTestExcelLateBinding
    {
        [Test]
        public void TestFillDataTable()
        {
            var db = SqlClientConnManager.GetConnManager("Data Source=.\\SQLEXPRESS;Initial Catalog=MediaDb;Integrated Security=True");
            var dt = db.GetDataTable("select top 1000 * from settings");

            var excel = new ExcelLateBinding();
            excel.Status += ExcelStatus;
            excel.NewWorkbook();
            //if (e.ErrorCode != 0) return;

            //excel.NewSheet("Daten");
            excel.SelectSheetFirst("TransactionData");
            excel.Header("Test");
            excel.NumberFormat = "#,##0.000000";
            excel.FillDataTable(dt, 4, 1);

            excel.NewSheet("Daten2");
            excel.Header("Test2");
            excel.NumberFormat = "#,##0.00";
            excel.FillDataTable(dt, 4, 1);

            excel.Quit();

        }



        private static void ExcelError(Exception ex, string message)
        {
            var s = $"Error:{ex.Message}:{message}";
            Debug.Print(s);
        }

        private static void ExcelStatus(string message)
        {
            Debug.Print(message);
        }
    }
}
