using System.Data;
using System.IO;
using System.Text;

namespace Bodoconsult.Core.Office
{
    public class Csv
    {

        private readonly StringBuilder _erg = new StringBuilder();

        private int _columnCount;

        public bool Header; // { get; set; }

        public string LineSeparator; // { get; set; }

        public string FieldSeparator; // { get; set; }


        public DataTable Data; // { get; set; }

        public string FileName; // { get; set; }

        public Csv()
        {
            Header = true;
            LineSeparator = "\r\n";
            FieldSeparator = ";";
        }

        public void Export()
        {

            _columnCount = Data.Columns.Count - 1;

            if (Header)
            {
                var i = 0;

                foreach (DataColumn f in Data.Columns)
                {
                    _erg.Append(f.ColumnName + ((i < _columnCount) ? FieldSeparator : ""));
                    i++;
                }

                _erg.Append(LineSeparator);
            }


            foreach (DataRow r in Data.Rows)
            {

                var i = 0;

                foreach (DataColumn f in Data.Columns)
                {
                    _erg.Append(r[f.ColumnName] + ((i < _columnCount) ? FieldSeparator : ""));
                    i++;
                }

                _erg.Append(LineSeparator);
            }


            var sw = new StreamWriter(FileName, false, Encoding.GetEncoding("utf-8"));
            sw.Write(_erg.ToString());
            sw.Close();
            sw.Dispose();

        }


    }
}
