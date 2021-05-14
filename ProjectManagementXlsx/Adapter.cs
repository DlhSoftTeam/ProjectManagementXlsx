using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Globalization;
using System.Xml.Linq;
using DlhSoft.Windows.Data;
using System.Collections.Generic;

namespace ProjectManagementXlsx
{
    public static class Adapter
    {
        public static byte[] GetExcelBytes(string projectXml, string projectManagementFrameworkLicense = null)
        {
            if (projectManagementFrameworkLicense != null)
                Licensing.Serialization.SetLicense(projectManagementFrameworkLicense);

            var taskManager = new TaskManager();
            taskManager.LoadProjectXml(projectXml);

            var assembly = typeof(Adapter).Assembly;
            var x = assembly.GetManifestResourceNames();
            using (var excelTemplateStream = assembly.GetManifestResourceStream("ProjectManagementXlsx.ProjectTemplate.xlsx"))
            {
                using (var excelStream = new MemoryStream())
                {
                    excelTemplateStream.CopyTo(excelStream);
                    excelStream.Seek(0, SeekOrigin.Begin);
                    using (var zip = new ZipArchive(excelStream, ZipArchiveMode.Update))
                    {
                        var sheetEntry = zip.Entries.Single(e => e.FullName == "xl/worksheets/sheet1.xml");
                        using (var sheetStream = sheetEntry.Open())
                        {
                            var sheetXml = XDocument.Load(sheetStream);
                            var stringsEntry = zip.Entries.Single(e => e.FullName == "xl/sharedStrings.xml");
                            using (var stringsStream = stringsEntry.Open())
                            {
                                var stringsXml = XDocument.Load(stringsStream);
                                AddTasksToExcel(taskManager, sheetXml, stringsXml);
                                stringsStream.Seek(0, SeekOrigin.Begin);
                                stringsXml.Save(stringsStream);
                            }
                            sheetStream.Seek(0, SeekOrigin.Begin);
                            sheetXml.Save(sheetStream);
                        }
                    }
                    return excelStream.ToArray();
                }
            }
        }

        private static void AddTasksToExcel(TaskManager taskManager, XDocument sheet, XDocument strings)
        {
            var sheetData = sheet.Element(XName.Get("worksheet", xmlns)).Element(XName.Get("sheetData", xmlns));
            var r = 1;
            foreach (var task in taskManager.Items)
            {
                r++;

                var row = new XElement(XName.Get("row", xmlns));
                row.SetAttributeValue("r", r);
                row.SetAttributeValue("spans", "1:13");
                sheetData.Add(row);

                AddValueToExcel(row, r, "A", taskManager.GetIndexString(task));
                AddValueToExcel(row, r, "B", task.Indentation);
                AddStringValueToExcel(strings, row, r, "C", task.Tag as string);
                AddDateValueToExcel(row, r, "D", task.Start);
                AddDecimalValueToExcel(row, r, "E", taskManager.GetEffort(task).TotalHours);
                AddDecimalValueToExcel(row, r, "F", taskManager.GetDuration(task).TotalDays);
                AddDateValueToExcel(row, r, "G", task.Finish);
                AddBoolValueToExcel(row, r, "H", task.IsMilestone);
                AddBoolValueToExcel(row, r, "I", taskManager.IsCompleted(task));
                AddPercentValueToExcel(row, r, "J", taskManager.GetCompletion(task));
                AddStringValueToExcel(strings, row, r, "K", taskManager.GetPredecessorsString(task));
                AddStringValueToExcel(strings, row, r, "L", taskManager.GetAssignmentsString(task));
                AddDecimalValueToExcel(row, r, "M", taskManager.GetCost(task));
            }
        }

        private static void AddValueToExcel(XElement row, int reference, string column, object value, string t = null, string s = null)
        {
            var cell = new XElement(XName.Get("c", xmlns));
            cell.SetAttributeValue("r", column + reference);
            if (t != null)
                cell.SetAttributeValue("t", t);
            if (s != null)
                cell.SetAttributeValue("s", s);
            row.Add(cell);

            var valueElement = new XElement(XName.Get("v", xmlns));
            valueElement.SetValue(value);
            cell.Add(valueElement);
        }

        private static void AddStringValueToExcel(XDocument strings, XElement row, int reference, string column, string value)
        {
            if (string.IsNullOrEmpty(value))
                return;
            var stringReference = AddStringToExcel(strings, value);
            AddValueToExcel(row, reference, column, stringReference, t: "s");
        }
        private static string AddStringToExcel(XDocument strings, string value)
        {
            var stringsData = strings.Element(XName.Get("sst", xmlns));
            var count = int.Parse(stringsData.Attribute("count").Value, CultureInfo.InvariantCulture);
            var uniqueCount = int.Parse(stringsData.Attribute("uniqueCount").Value, CultureInfo.InvariantCulture);
            var stringElements = stringsData.Elements(XName.Get("si", xmlns)).ToArray();
            var index = -1;
            for (var i = 0; i < stringElements.Length; i++)
            {
                var stringElement = stringElements[i];
                if (stringElement.Element(XName.Get("t", xmlns)).Value == value)
                {
                    index = i;
                    break;
                }
            }
            if (index < 0)
            {
                var stringElement = new XElement(XName.Get("si", xmlns));
                stringsData.Add(stringElement);
                var valueElement = new XElement(XName.Get("t", xmlns));
                valueElement.SetValue(value);
                stringElement.Add(valueElement);
                index = uniqueCount;
                uniqueCount++;
            }
            count++;
            stringsData.SetAttributeValue("count", count.ToString(CultureInfo.InvariantCulture));
            stringsData.SetAttributeValue("uniqueCount", uniqueCount.ToString(CultureInfo.InvariantCulture));
            return index.ToString(CultureInfo.InvariantCulture);
        }

        private static void AddDecimalValueToExcel(XElement row, int reference, string column, double value)
        {
            AddValueToExcel(row, reference, column, value, s: "3");
        }
        private static void AddPercentValueToExcel(XElement row, int reference, string column, double value)
        {
            AddValueToExcel(row, reference, column, value, s: "4");
        }
        private static void AddBoolValueToExcel(XElement row, int reference, string column, bool value)
        {
            AddValueToExcel(row, reference, column, value, t: "b");
        }

        private static void AddDateValueToExcel(XElement row, int reference, string column, DateTime value)
        {
            AddValueToExcel(row, reference, column, GetStringFromDate(value), s: "2");
        }
        public static string GetStringFromDate(DateTime value)
        {
            var diff = (value - originDate).TotalDays;
            return diff.ToString(CultureInfo.InvariantCulture);
        }

        public static string GetProjectXml(byte[] excelBytes, string projectManagementFrameworkLicense = null)
        {
            if (projectManagementFrameworkLicense != null)
                Licensing.Serialization.SetLicense(projectManagementFrameworkLicense);

            using (var excelStream = new MemoryStream(excelBytes))
            {
                using (var zip = new ZipArchive(excelStream, ZipArchiveMode.Read))
                {
                    var sheetEntry = zip.Entries.Single(e => e.FullName == "xl/worksheets/sheet1.xml");
                    using (var sheetStream = sheetEntry.Open())
                    {
                        var sheetXml = XDocument.Load(sheetStream);
                        var stringsEntry = zip.Entries.Single(e => e.FullName == "xl/sharedStrings.xml");
                        using (var stringsStream = stringsEntry.Open())
                        {
                            var stringsXml = XDocument.Load(stringsStream);

                            var taskManager = new TaskManager();
                            ReadTasksFromExcel(taskManager, sheetXml, stringsXml);

                            var projectXml = taskManager.GetProjectXml();
                            return projectXml;
                        }
                    }
                }
            }
        }

        private static void ReadTasksFromExcel(TaskManager taskManager, XDocument sheet, XDocument strings)
        {
            var sheetData = sheet.Element(XName.Get("worksheet", xmlns)).Element(XName.Get("sheetData", xmlns));
            var tasks = new Dictionary<int, TaskItem>();
            foreach (var row in sheetData.Elements(XName.Get("row", xmlns)))
            {
                var r = int.Parse(row.Attribute("r").Value, CultureInfo.InvariantCulture);
                if (r <= 1)
                    continue;

                var task = new TaskItem();
                taskManager.Items.Add(task);
                tasks.Add(r, task);

                task.Indentation = int.Parse(GetValueFromExcel(row, r, "B"), CultureInfo.InvariantCulture);
                task.Tag = GetValueFromExcel(row, r, "C", strings);
                task.Start = GetDateFromString(GetValueFromExcel(row, r, "D"));
                task.Finish = GetDateFromString(GetValueFromExcel(row, r, "G"));
                task.IsMilestone = new[] { "true", "1" }.Contains(GetValueFromExcel(row, r, "H").ToLowerInvariant());
                taskManager.SetCompletion(task, double.Parse(GetValueFromExcel(row, r, "J")));
                taskManager.UpdateAssignments(task, GetValueFromExcel(row, r, "L", strings));
                task.ExecutionCost = double.Parse(GetValueFromExcel(row, r, "M"));
            }
            foreach (var row in sheetData.Elements(XName.Get("row", xmlns)))
            {
                var r = int.Parse(row.Attribute("r").Value, CultureInfo.InvariantCulture);
                if (r <= 1)
                    continue;

                var taskItem = tasks[r];
                taskManager.UpdatePredecessors(taskItem, GetValueFromExcel(row, r, "K", strings));
            }
        }

        private static string GetValueFromExcel(XElement row, int reference, string column, XDocument strings = null)
        {
            var cell = row.Elements(XName.Get("c", xmlns)).SingleOrDefault(c => c.Attribute("r")?.Value == column + reference);
            var value = cell?.Element(XName.Get("v", xmlns))?.Value;
            if (strings != null && cell?.Attribute("t")?.Value == "s")
            {
                value = GetStringFromExcel(strings, value);
            }
            return value ?? string.Empty;
        }
        private static string GetStringFromExcel(XDocument strings, string indexString)
        {
            var index = int.Parse(indexString, CultureInfo.InvariantCulture);
            var stringsData = strings.Element(XName.Get("sst", xmlns));
            var stringElement = stringsData.Elements(XName.Get("si", xmlns)).Skip(index).First();
            return stringElement.Element(XName.Get("t", xmlns)).Value;
        }

        private static DateTime GetDateFromString(string value)
        {
            var numericValue = double.Parse(value, CultureInfo.InvariantCulture);
            return originDate.AddDays(numericValue);
        }

        private const string xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private static readonly DateTime originDate = new DateTime(1899, 12, 30);
    }

    namespace Licensing
    {
        public static class Serialization
        {
            public static void SetLicense(string projectManagementFrameworkLicense)
            {
                DlhSoft.Windows.Data.Licensing.TaskManager.SetLicense(projectManagementFrameworkLicense);
            }
        }
    }
}
