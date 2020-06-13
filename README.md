# What does the library

Bodoconsult.Core.Office library simplifies creating OpenXml spredsheets (xlsx) for database data in form of System.Data.DataTable objects.

It was developed with the intention to easily export database data to Excel spreadsheets.

# How to use the library

The following code samples make usage of repository <https://github.com/RobertLeisner/Bodoconsult.Core.Database> for accessing Microsoft SqlServer database.
The method GetDataTable used below returns a plain old System.Data.DataTable object.

## Use ExcelLateBinding class

The ExcelLateBinding class uses COM late binding to export a DataTable (in the sample code the variable dt) to an Excel spreadsheet.


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

            execl.Dispose();

## Use XlsxOpenXml class

The XlsxOpenXml class writes the content of a DataTable (in the sample code the variable dt) directly to an OpenXml spreadsheet file.

            var db = SqlClientConnManager.GetConnManager("Data Source=.\\SQLEXPRESS;Initial Catalog=MediaDb;Integrated Security=True");
            var dt = db.GetDataTable("select top 1000 * from settings");

            var oe = new XlsxOpenXml();
            oe.Status += ExcelStatus;
            oe.Error += ExcelError;
            oe.NumberFormatDouble = "#,##0.000000";
            oe.NewWorkbook(path);

            oe.NewSheet("Daten");

            oe.SelectRange(1, 1);
            oe.Style = XlsxStyles.Header;
            oe.SetValue("Hallo Welt1");
			
            oe.FillDataTable(dt, 4, 1);

            oe.Quit();


# About us

Bodoconsult (<http://www.bodoconsult.de>) is a Munich based software development company.

Robert Leisner is senior software developer at Bodoconsult. See his profile on <http://www.bodoconsult.de/Curriculum_vitae_Robert_Leisner.pdf>.

