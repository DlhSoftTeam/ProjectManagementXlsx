using System;
using System.IO;
using System.Text;

namespace ProjectManagementXlsxConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("This tool converts Microsoft Project XML files to Excel, or Excel (specific format) to Microsoft Project XML.");
                Console.WriteLine();
                Console.WriteLine("Usage:");
                Console.WriteLine("GanttChartExcelConverter source.xml [target[.xslx]]");
                Console.WriteLine("GanttChartExcelConverter source.xslx [target[.xml]]");
                return;
            }

            var source = args[0];
            var sourcePath = Path.GetDirectoryName(source);
            var sourceName = Path.GetFileNameWithoutExtension(source);
            var sourceExt = Path.GetExtension(source).ToLowerInvariant();
            var sourceIsExcel = sourceExt == ".xlsx";
            var targetIsExcel = !sourceIsExcel;
            var target = args.Length > 1 ? args[1] : Path.Combine(sourcePath, sourceName + (targetIsExcel ? ".xlsx" : ".xml"));

            Console.WriteLine("Converting " + source + " to " + target + " (from " + (sourceIsExcel ? "Excel" : "Project XML") + " to " + (targetIsExcel ? "Excel" : "Project XML") + ")...");

            if (sourceIsExcel)
            {
                var excelBytes = File.ReadAllBytes(source);
                var projectXml = ProjectManagementXlsx.Adapter.GetProjectXml(excelBytes);
                File.WriteAllBytes(target, Encoding.UTF8.GetBytes(projectXml));
            }
            else
            {
                var projectXml = Encoding.UTF8.GetString(File.ReadAllBytes(source));
                var excelBytes = ProjectManagementXlsx.Adapter.GetExcelBytes(projectXml);
                File.WriteAllBytes(target, excelBytes);
            }

            Console.WriteLine("Done.");
        }
    }
}
