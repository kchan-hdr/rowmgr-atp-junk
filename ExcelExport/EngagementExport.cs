using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelExport
{
    public class EngagementExport : Exporter<ROWM.Dal.OwnerRepository.EngagementDto>
    {
        public EngagementExport(IEnumerable<ROWM.Dal.OwnerRepository.EngagementDto> d) : base(d) { }

        public override byte[] Export()
        {
            reportname = "Community Engagement Report";

            using (var memory = new MemoryStream())
            {
                var doc = MakeDoc(memory);
                bookPart = doc.AddWorkbookPart();
                bookPart.Workbook = new Workbook();
                sheets = doc.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                MakeStyles();

                uint pagenum = 1;
                var lines = items.SelectMany(ix => ix.Project).Distinct();
                foreach (var line in lines)
                {
                    WriteEngagement(++pagenum, line);
                    WriteAction(++pagenum, line);
                }

                doc.Close();

                return memory.ToArray();
            }
        }

        void WriteEngagement(uint pageId, string line)
        {
            var p = bookPart.AddNewPart<WorksheetPart>($"uId{pageId}");
            var d = new SheetData();
            p.Worksheet = new Worksheet(d);

            uint row = 1;

            var rx = InsertRow(row++, d);
            var titleCell = WriteText(rx, "A", $"{line} {reportname}", 1);
            
            rx = InsertRow(row++, d);
            var dateCell = WriteText(rx, "A", $"as of {DateTime.Now.ToLongDateString()}");

            var hr = InsertRow(row++, d);
            var c = 0;
            WriteText(hr, GetColumnCode(c++), "Parcel ID", 1);
            WriteText(hr, GetColumnCode(c++), "Impacted", 1);
            WriteText(hr, GetColumnCode(c++), "Owner Name", 1);
            WriteText(hr, GetColumnCode(c++), "Contact Name", 1);
            WriteText(hr, GetColumnCode(c++), "Date of Contact", 1);
            WriteText(hr, GetColumnCode(c++), "Channel", 1);
            WriteText(hr, GetColumnCode(c++), "Type", 1);
            WriteText(hr, GetColumnCode(c++), "Title", 1);
            WriteText(hr, GetColumnCode(c++), "Contact Summary", 1);
            WriteText(hr, GetColumnCode(c++), "Agent Name", 1);


            var eng2 = from px in items.Where(px => px.Project.Contains(line))
                       from lx in px.Logs.Where(ix => ix.ProjectPhase.EndsWith("Engagement"))
                       select new { px.Apn, px.IsImpacted, px.OwnerName, cot= lx };

            foreach(var par in eng2.OrderByDescending(cdt => cdt.cot.DateAdded))
            //foreach (var par in items.Where(ix => ix.Project.Contains(line)).OrderBy(px => px.TrackingNumber))
            //{
            //    if (par.Logs.Any())
            //    {
            //        foreach (var cot in par.Logs.Where(px => px.ProjectPhase.EndsWith("Engagement")).OrderBy(pdt => pdt.DateAdded))
                    {
                        var r = InsertRow(row++, d);
                        c = 0;
                        WriteText(r, GetColumnCode(c++), par.Apn);
                        WriteText(r, GetColumnCode(c++), par.IsImpacted ? "Impacted Parcel" : "Parcel Not Impacted");
                        WriteText(r, GetColumnCode(c++), string.Join(" | ", par.OwnerName));
                        WriteText(r, GetColumnCode(c++), par.cot.ContactNames);
                        WriteDate(r, GetColumnCode(c++), par.cot.DateAdded.LocalDateTime);

                        WriteText(r, GetColumnCode(c++), par.cot.ContactChannel);
                        WriteText(r, GetColumnCode(c++), par.cot.ProjectPhase);

                        WriteText(r, GetColumnCode(c++), par.cot.Title);
                        WriteText(r, GetColumnCode(c++), par.cot.Notes);
                        WriteText(r, GetColumnCode(c++), par.cot.AgentName);

            }

            sheets.Append(new Sheet { Id = bookPart.GetIdOfPart(p), SheetId = pageId, Name = $"{line} - Outreach" });
            bookPart.Workbook.Save();
        }

        void WriteAction(uint pageId, string line)
        {
            var p = bookPart.AddNewPart<WorksheetPart>($"uId{pageId}");
            var d = new SheetData();
            p.Worksheet = new Worksheet(d);

            uint row = 1;

            var rx = InsertRow(row++, d);
            WriteText(rx, "A", "Community Engagement Actions", 1);
            rx = InsertRow(row++, d);
            WriteText(rx, "A", DateTime.Now.ToLongDateString());
            
            var hr = InsertRow(row++, d);
            var c = 0;
            WriteText(hr, GetColumnCode(c++), "Parcel ID", 1);
            WriteText(hr, GetColumnCode(c++), "Owner Name", 1);
            WriteText(hr, GetColumnCode(c++), "Contacts", 1);
            WriteText(hr, GetColumnCode(c++), "Action Item", 1);
            WriteText(hr, GetColumnCode(c++), "Action Item Owner", 1);
            WriteText(hr, GetColumnCode(c++), "Due Date", 1);
            WriteText(hr, GetColumnCode(c++), "Status", 1);

            foreach (var par in items.Where(px => px.Project.Contains(line)).OrderBy(px => px.TrackingNumber))
            {
                foreach (var cot in par.Actions.OrderBy(pdt => pdt.Due))
                {
                    var st = Enum.GetName(typeof(ROWM.Dal.ActionStatus), cot.Status);

                    var r = InsertRow(row++, d);
                    c = 0;
                    WriteText(r, GetColumnCode(c++), par.Apn);
                    WriteText(r, GetColumnCode(c++), string.Join(" | ", par.OwnerName));
                    WriteText(r, GetColumnCode(c++), par.ContactNames);
                    WriteText(r, GetColumnCode(c++), cot.Action);
                    WriteText(r, GetColumnCode(c++), cot.Assigned);
                    WriteDate(r, GetColumnCode(c++), cot.Due.HasValue ? cot.Due.Value.LocalDateTime : default);
                    WriteText(r, GetColumnCode(c++), st);

                }
            }

            sheets.Append(new Sheet { Id = bookPart.GetIdOfPart(p), SheetId = pageId, Name = $"{line} - Action Items" });

            bookPart.Workbook.Save();
        }
    }
}
