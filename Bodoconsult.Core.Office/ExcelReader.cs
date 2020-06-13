//using System;
//using System.Data;
//using System.Data.OleDb;

//namespace Bodoconsult.Office
//{
//    /// <summary>
//    /// Class to read excel sheets in a <see cref="System.Data.DataTable"/>
//    /// </summary>
//    public class ExcelReader
//    {

//        private OleDbConnection _conn;


//        /// <summary>
//        /// Source filename of the Excel file for data import
//        /// </summary>
//        public string FileName { get; set; }

//        /// <summary>
//        /// All sheet names in the Excel file
//        /// </summary>
//        public string[] SheetNames { get; private set; }

//        /// <summary>
//        /// Open the connection to the Excel file and get sheet names
//        /// </summary>
//        public void OpenFile()
//        {

//            var connString = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=YES"";", FileName);

//            _conn = new OleDbConnection(connString);
//            try
//            {
//                _conn.Open();

//                var schemaTable = _conn.GetOleDbSchemaTable(
//                        OleDbSchemaGuid.Tables,
//                        new object[] { null, null, null, "TABLE" });

//                var v = new DataView(schemaTable) { Sort = "Table_Name", RowFilter = "Table_Name is not null" };
//                schemaTable = v.ToTable();

//                SheetNames = new string[schemaTable.Rows.Count];
//                var i = 0;
//                foreach (DataRow r in schemaTable.Rows)
//                {
//                    SheetNames[i] = r["Table_name"].ToString();
//                    i++;
//                }

//            }
//            catch (Exception ex)
//            {
//                var msg = string.Format("Error:ExcelFile:{0}", FileName);
//                throw new Exception(msg, ex);
//            }
//        }

//        /// <summary>
//        /// Get the datatable for a sheet name
//        /// </summary>
//        /// <param name="sheetName"></param>
//        /// <returns></returns>
//        public DataTable GetSheet(string sheetName)
//        {
//            try
//            {

//                if (!sheetName.EndsWith("$")) sheetName += "$";

//                var dt = new DataTable();

//                var sql = string.Format("select * from [{0}]", sheetName);

//                var cmd = new OleDbCommand { CommandText = sql, Connection = _conn };

//                var da = new OleDbDataAdapter { SelectCommand = cmd };
//                dt.Clear();
//                da.Fill(dt);

//                return dt;
//            }
//            catch (Exception ex)
//            {
//                var msg = string.Format("Error:ExcelFile:{0}:GetSheetByName:{1}", FileName, sheetName);
//                throw new Exception(msg, ex);
//            }
//        }

//        /// <summary>
//        ///  Get the datatable for a sheet by the index of the sheet (Attention: index is based on indexing of <see cref="SheetNames"/>)
//        /// </summary>
//        /// <param name="index"></param>
//        /// <returns></returns>
//        public DataTable GetSheet(int index)
//        {
//            try
//            {
//                return GetSheet(SheetNames[index]);
//            }
//            catch (Exception ex)
//            {
//                var msg = string.Format("Error:ExcelFile:{0}:GetSheetByIndex:{1}", FileName, index);
//                throw new Exception(msg, ex);
//            }
//        }
//    }
//}
