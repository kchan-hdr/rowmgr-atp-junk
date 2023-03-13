using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Drawing;
using System.IO;

namespace ROWM.Reports
{
    public class ReportingMethods
    {

        public MemoryStream StandardReport(string reportName, int numCols, string logoPath, DbDataReader dt, bool doubleheader)
        {
            var results = CreatePackageAddSheets(new List<string>() { "Report Info", reportName });
            var stream = results.Item1;
            var package = results.Item2;
            var dict_worksheets = results.Item3;

            // add data to report info
            Dictionary<string, string> infodict = new Dictionary<string, string>();
            infodict.Add(reportName, "");
            infodict.Add("Date of Report Generation", DateTime.UtcNow.ToString());
            dict_worksheets["Report Info"].Cells.LoadFromCollection(infodict, false);

            // add image
            //AddPicture(dict_worksheets["Report Info"], logoPath, 3, 0, 79, 213);

            // add formatting
            ChangeColumnWidth(dict_worksheets[reportName], 1, numCols, 25);

            // send data to report data
            dict_worksheets[reportName].Cells.LoadFromDataReader(dt, true);
            if (doubleheader)
            {
                // format headers
                dict_worksheets[reportName].InsertRow(1, 1);
                FormatFont(dict_worksheets[reportName], "1:1", "Calibri", fontsize: 13, bold: true, wraptext: true);
                FormatFont(dict_worksheets[reportName], "2:2", "Calibri", fontsize: 11, bold: true, wraptext: true);

                dict_worksheets[reportName].Cells["C1"].Value = "Parcel";
                dict_worksheets[reportName].Cells["C1:J1"].Merge = true;
                dict_worksheets[reportName].Cells["C1:J1"].Style.Font.Bold = true;
                dict_worksheets[reportName].Cells["C1:J1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                foreach (var x in dict_worksheets[reportName].Cells["A2:J2"])
                {
                    x.Value = x.Value.ToString().Replace("_", " ");
                }

                // break field names into two rows
                int colCount = dict_worksheets[reportName].Dimension.Columns;
                foreach (var x in dict_worksheets[reportName].Cells[$"{ExcelCellBase.TranslateFromR1C1("R2C11", 0, 0)}:{ExcelCellBase.TranslateFromR1C1($"R2C{colCount}", 0, 0)}"])
                {
                    int index = x.Value.ToString().IndexOf("__");
                    x.Offset(-1, 0).Value = x.Value.ToString().Substring(0, index).Replace("_", " ");
                    x.Value = x.Value.ToString().Substring(index + 2).Replace("__", " - ").Replace("_", " ");
                }

                // delete repeating values in top row, merge, and horizontal alignment
                var mergeStart = $"{ExcelCellBase.TranslateFromR1C1("R1C11", 0, 0)}";
                var stringCheck = "";
                foreach (var x in dict_worksheets[reportName].Cells[$"{ExcelCellBase.TranslateFromR1C1("R1C11", 0, 0)}:{ExcelCellBase.TranslateFromR1C1($"R1C{colCount}", 0, 0)}"])
                {
                    if (x.Address == $"{ExcelCellBase.TranslateFromR1C1("R2C11", 0, 0)}")
                    {
                        stringCheck = x.Value.ToString();
                    }

                    else
                    {
                        int index = x.FullAddressAbsolute.ToString().IndexOf("!");
                        var absAddress = x.FullAddressAbsolute.ToString().Substring(index + 1);
                        if (x.Value.ToString() == stringCheck)
                        {
                            x.Value = "";
                            dict_worksheets[reportName].Cells[$"{mergeStart}:{absAddress}"].Merge = true;
                            dict_worksheets[reportName].Cells[$"{mergeStart}:{absAddress}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        }
                        else
                        {
                            stringCheck = x.Value.ToString();
                            mergeStart = absAddress;
                        }
                    }
                }
            }

            else
            {
                // format headers
                FormatFont(dict_worksheets[reportName], "1:1", "Calibri", fontsize: 11, bold: true, wraptext: true);

                // break field names into two rows
                int colCount = dict_worksheets[reportName].Dimension.Columns;
                foreach (var x in dict_worksheets[reportName].Cells[$"{ExcelCellBase.TranslateFromR1C1("R1C1", 0, 0)}:{ExcelCellBase.TranslateFromR1C1($"R1C{colCount}", 0, 0)}"])
                {
                    x.Value = x.Value.ToString().Replace("_", " ");
                }
            }
            dict_worksheets[reportName].DeleteColumn(1);
            dict_worksheets[reportName].DeleteColumn(1);


            // export file
            package.Save();
            stream.Position = 0;
            return stream;
        }

        public (MemoryStream, ExcelPackage, Dictionary<string, ExcelWorksheet>) CreatePackageAddSheets(List<string> sheets)
        {
            MemoryStream stream = new MemoryStream();
            ExcelPackage package = new ExcelPackage(stream);
            Dictionary<string, ExcelWorksheet> dict_worksheets = new Dictionary<string, ExcelWorksheet>();
            foreach (string sheet in sheets) { dict_worksheets.Add(sheet, package.Workbook.Worksheets.Add(sheet)); }

            return (stream, package, dict_worksheets);
        }

        public void AddPicture(ExcelWorksheet worksheet, string filepath, int rowindex = 0, int colindex = 0, int height = 1, int width = 1)
        {
            // try to use resource
            ExcelPicture pic = worksheet.Drawings.AddPicture("", new FileInfo(filepath));
            pic.SetPosition(rowindex, 0, colindex, 0);
            pic.SetSize(width, height);
        }


        public void FormatFont(ExcelWorksheet worksheet, string range, string fontname = "", int fontsize = 0, bool bold = false, bool italic = false, bool underline = false, bool wraptext = false,
            int[] fill = null, bool merge = false)
        {
            if (!string.IsNullOrEmpty(fontname) && fontsize != 0) { worksheet.Cells[range].Style.Font.SetFromFont(new Font(fontname, fontsize)); }
            worksheet.Cells[range].Style.Font.Bold = bold;
            worksheet.Cells[range].Style.Font.Italic = italic;
            worksheet.Cells[range].Style.Font.UnderLine = underline;
            worksheet.Cells[range].Style.WrapText = wraptext;
            worksheet.Cells[range].Merge = merge;
            var fills = fill ?? new int[0];
            if (fills.Length > 0)
            {
                worksheet.Cells[range].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[range].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(fills[0], fills[1], fills[2]));
            }
        }

        public void ChangeRowHeight(ExcelWorksheet worksheet, int rowStartIndex, int rowEndIndex, double rowHeight)
        {
            for (int j = rowStartIndex; j < rowEndIndex; j++) { worksheet.Row(j).Height = rowHeight; }
        }

        public void ChangeColumnWidth(ExcelWorksheet worksheet, int columnStartIndex, int columnEndIndex, double columnWidth)
        {
            for (int j = columnStartIndex; j <= columnEndIndex; j++) { worksheet.Column(j).Width = columnWidth; }
        }


    }
}