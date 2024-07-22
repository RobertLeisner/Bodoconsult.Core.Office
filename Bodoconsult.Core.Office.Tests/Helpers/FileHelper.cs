using System.Diagnostics;

namespace Bodoconsult.Core.Office.Tests.Helpers
{
    public static class FileHelper
    {

        public static string TempPath { get; set; } = @"D:\tmp";

        public static void StartExcel(string path)
        {
            Process.Start("\"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE\"", path);
        }
    }
}
