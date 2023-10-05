using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport
{
    public class PreAcquisitionExport : Exporter<string>    // placeholder type. Not used
    {
        public DataTable MyDatatable { get; set; }

        public PreAcquisitionExport(IEnumerable<string> d, string logo) : base(d) { }

        public override byte[] Export()
        {
            reportname = "Pre-Acquisition Report";
            return base.Export();
        }

        protected override void Write(uint pageId)
        {
            var p = bookPart.AddNewPart<WorksheetPart>($"uId{pageId}");
            var d = new SheetData();
            p.Worksheet = new Worksheet(d);

            uint row = 1;

            //row = WriteLogo(row, p, d, reportname);

            var hr = InsertRow(row++, d);
            var c = 0;

            foreach(DataColumn column in MyDatatable.Columns)
            {
                WriteText(hr, GetColumnCode(c++), column.ColumnName, 1);
            }

            foreach(DataRow datarow in MyDatatable.Rows)
            {
                var r = InsertRow(row++, d);
                c = 0;

                foreach(var v in datarow.ItemArray)
                {
                    WriteText(r, GetColumnCode(c++), v?.ToString() ?? "");
                }
            }

            var format = SetColumnWidth((uint) MyDatatable.Columns.Count);
            p.Worksheet.InsertBefore(format, d);
            sheets.Append(new Sheet { Id = bookPart.GetIdOfPart(p), SheetId = pageId, Name = "Pre-Acquisition" });
            bookPart.Workbook.Save();
        }

        Columns SetColumnWidth(uint numCols)
        {
            var columns = new Columns();
            var col = new Column
            {
                CustomWidth = true,
                Min = 1U,
                Max = (DocumentFormat.OpenXml.UInt32Value)numCols,
                Width = (DocumentFormat.OpenXml.DoubleValue)20
            };

            columns.Append(col);
            return columns;
        }
    }
}
