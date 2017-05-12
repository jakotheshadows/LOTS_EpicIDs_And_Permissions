using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using Microsoft.Office.Interop.Excel;

namespace LOTS_EpicIDs_And_Permissions
{
    class Program
    {
        static void Main(string[] args)
        {
            var data = new List<string[]>();
            var sb = new StringBuilder();
            var app = new Application();
            var interopWb = app.Workbooks.Open(Path.Combine(Environment.CurrentDirectory, "Epics.xls"));
            interopWb.SaveAs(Filename: Path.Combine(Environment.CurrentDirectory, "Epics.xlsx"), FileFormat: XlFileFormat.xlOpenXMLWorkbook);
            interopWb.Close();
            app.Quit();

            using (var pkg = new ExcelPackage(new FileInfo("Epics.xlsx")))
            {
                var ws = pkg.Workbook.Worksheets.FirstOrDefault();
                const string PERM_MARKER = "Permission(s)";
                if (ws != null)
                {
                    var header = new Dictionary<string, int>();
                    for (int rowIndex = 4; rowIndex <= ws.Dimension.End.Row; rowIndex += 1)
                    {
                        if (rowIndex == 4)
                        {
                            header = ExcelHelper.GetExcelHeader(ws, rowIndex);
                        }
                        else
                        {
                            string id = ExcelHelper.ParseWorksheetValue(ws, header, rowIndex, "Project ID");
                            string description = ExcelHelper.ParseWorksheetValue(ws, header, rowIndex, "Description");
                            string itemType = ExcelHelper.ParseWorksheetValue(ws, header, rowIndex, "Item Type");
                            if (itemType != "Folder")
                            {
                                if (!String.IsNullOrWhiteSpace(description))
                                {
                                    if (description.Contains(PERM_MARKER))
                                    { 
                                        description =
                                            description.Substring(description.LastIndexOf(PERM_MARKER) +
                                                                  PERM_MARKER.Length);
                                        var tokens = description.Split(null);
                                        foreach (string token in tokens)
                                        {
                                            if (sb.Length + token.Length + 3 >= 32767)
                                            {
                                                data.Add(new string[2] { id, sb.ToString() });
                                                sb.Clear();
                                            }

                                            if (token.Contains("_"))
                                            {
                                                sb.Append(Environment.NewLine + token + " ");
                                            }
                                            else
                                            {
                                                sb.Append(token + " ");
                                            }
                                        }
                                        data.Add(new string[2] {id, sb.ToString()});
                                    }
                                    else
                                    {
                                        data.Add(new string[2] {id, "DEBUG: Permission(s) marker not found"});
                                    }
                                }
                                else
                                {
                                    data.Add(new string[2] {id, "DEBUG: Description was empty"});
                                }
                            }
                        }
                    }
                }
            }
            using (var pkg = new ExcelPackage())
            {
                var wb = pkg.Workbook;
                var ws = wb.Worksheets.Add("EPIC Permissions");
                ws.Cells["A1"].Value = "Project ID";
                ws.Cells["B1"].Value = "Permissions";

                for (int i = 0; i < data.Count; i += 1)
                {
                    ws.Cells["A" + (i + 2)].Value = data.ElementAt(i)[0];
                    ws.Cells["B" + (i + 2)].Value = data.ElementAt(i)[1];
                }

                pkg.SaveAs(new FileInfo("EpicPermissions.xlsx"));
            }
        }
    }

    public static class ExcelHelper
    {
        public static Dictionary<string, int> GetExcelHeader(ExcelWorksheet workSheet, int rowIndex)
        {
            var header = new Dictionary<string, int>();

            if (workSheet != null)
            {
                for (int columnIndex = workSheet.Dimension.Start.Column;
                    columnIndex <= workSheet.Dimension.End.Column;
                    columnIndex++)
                {
                    if (workSheet.Cells[rowIndex, columnIndex].Value != null)
                    {
                        string columnName = workSheet.Cells[rowIndex, columnIndex].Value.ToString();

                        if (!header.ContainsKey(columnName) && !string.IsNullOrEmpty(columnName))
                        {
                            header.Add(columnName, columnIndex);
                        }
                    }
                }
            }

            return header;
        }

        public static string ParseWorksheetValue(ExcelWorksheet workSheet, Dictionary<string, int> header, int rowIndex, string columnName)
        {
            string value = string.Empty;
            int? columnIndex = header.ContainsKey(columnName) ? header[columnName] : (int?)null;

            if (workSheet != null && columnIndex != null && workSheet.Cells[rowIndex, columnIndex.Value].Value != null)
            {
                value = workSheet.Cells[rowIndex, columnIndex.Value].Value.ToString();
            }

            return value;
        }
    }
}
