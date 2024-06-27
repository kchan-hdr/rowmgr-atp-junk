
using Humanizer;
using ROWM.Dal;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ROWM.Models
{
    public class SyncfusionContactLogHelper
    {
        readonly ROWM_Context _context;
        public SyncfusionContactLogHelper(ROWM_Context c) => (_context) = c;
        public async Task<byte[]> Generate(Parcel parcel)
        {
#if DEBUG
            var names = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceNames();
#endif
            string projectName = await GetProjectNameForParcel(parcel.ParcelId);
            ContactLogDto dto = new ContactLogDto(parcel, projectName);
            var logs = parcel.ContactLog.Where(cx => cx.IsDeleted == false).OrderBy(cx => cx.DateAdded);

            using (var working = new MemoryStream())
            {
                var s = GetTemplateStream();
                if (s is null)
                    return default;
                await s.CopyToAsync(working);
                WordDocument dox = new WordDocument(s, Syncfusion.DocIO.FormatType.Docx);
                dox.MailMerge.Execute(dto.AsDataTable());
                AppendLogs(dox, logs);
                dox.Save(working, Syncfusion.DocIO.FormatType.Docx);
                dox.Close();

                return working.ToArray();
                //return new ReportPayload(
                //    $"{parcel.Assessor_Parcel_Number} Contact Log.docx",
                //    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                //    working.ToArray()
                //);
            }
        }

        private async Task<string> GetProjectNameForParcel(Guid parcelId)
        {
            var projectName = (from p in _context.Parcel.AsNoTracking()
                               join pa in _context.Allocations.AsNoTracking() on p.ParcelId equals pa.ParcelId
                               join pp in _context.ProjectParts.AsNoTracking() on pa.ProjectPartId equals pp.ProjectPartId
                               where p.ParcelId == parcelId && p.IsActive && pa.IsActive
                               select pp.Caption).FirstOrDefault();

            return projectName ?? "ATP";
        }

        #region file helper
        static readonly string _TEMPLATE = "ROWM.ReportTemplates.ATP_contact_log_template_sycfusion.docx";
        static Stream GetTemplateStream() =>
            System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(_TEMPLATE);
        #endregion

        #region report content
        public class ContactLogDto
        {
            public string ProjectName { get; set; } = "XXX LINE";
            public string ParcelNumber { get; set; } = "";
            public string VestedOwnerName { get; set; } = "";
            public string SiteAddress { get; set; } = "";
            public string VestedOwnerAddress { get; set; } = "";
            public string ContactInfo { get; set; } = "";

            internal ContactLogDto(Parcel p, string projectName)
            {
                ProjectName = projectName;
                ParcelNumber = p.Assessor_Parcel_Number;
                VestedOwnerName = p.Ownership?.FirstOrDefault()?.Owner?.PartyName ?? string.Empty;
                SiteAddress = p.SitusAddress;
                VestedOwnerAddress = p.Ownership?.FirstOrDefault()?.Owner?.OwnerAddress ?? string.Empty;
                ContactInfo = PrettyPrintContact(p);
            }

            static string PrettyPrintContact(Parcel p)
            {
                if (!p.ParcelContacts.Any(cx => cx.IsDeleted == false))
                    return "";

                var info = p.ParcelContacts
                    .FirstOrDefault(cx => cx.IsPrimaryContact && cx.IsDeleted == false);

                if (info == null)
                    info = p.ParcelContacts.FirstOrDefault(cx => cx.IsDeleted == false);

                if (info == null)
                {
                    System.Diagnostics.Trace.TraceWarning($"corrupted contact info {p.Assessor_Parcel_Number}");
                    return "";
                }

                var list = new List<string>
                {
                    info.FirstName
                };

                if (!string.IsNullOrWhiteSpace(info.HomePhone))
                    list.Add($"H {PrettyPrintPhoneNumber(info.HomePhone)}");

                if (!string.IsNullOrWhiteSpace(info.CellPhone))
                    list.Add($"M {PrettyPrintPhoneNumber(info.CellPhone)}");

                if (!string.IsNullOrWhiteSpace(info.WorkPhone))
                    list.Add($"W {PrettyPrintPhoneNumber(info.WorkPhone)}");

                if (!string.IsNullOrWhiteSpace(info.Email))
                    list.Add($"email {info.Email}");

                return list.Humanize(",");
            }

            static string PrettyPrintPhoneNumber(string p)
            {
                var util = PhoneNumbers.PhoneNumberUtil.GetInstance();
                try
                {
                    var ph = util.Parse(p, "US");
                    return util.Format(ph, PhoneNumbers.PhoneNumberFormat.RFC3966);
                }
                catch (Exception)
                {
                    return p;
                }
            }

            internal DataTable AsDataTable()
            {
                DataTable table = new DataTable();
                var properties = GetType().GetProperties();

                var columns = properties.Select(p => new DataColumn { ColumnName = p.Name, DataType = p.PropertyType }).ToArray();
                table.Columns.AddRange(columns);

                DataRow r = table.NewRow();
                foreach (var p in properties)
                    r[p.Name] = p.GetValue(this);

                table.Rows.Add(r);
                return table;
            }
        }
        #endregion

        internal static void AppendLogs(WordDocument dox, IEnumerable<ContactLog> logs)
        {
            var sections = dox.Sections;
            if (sections.Count == 0)
                return;

            var tables = sections[0].Body.Tables;
            if (tables.Count == 0)
                return;

            var table = tables[0];

            var rows = table.Rows;
            if (rows.Count < 8)
                return;

            var rt = rows[7];
            var cells = rt.Cells;
            if (cells.Count < 2)
                return;

            foreach (ContactLog log in logs)
            {
                var r = table.AddRow(true);

                var dateCell = r.Cells[0];
                var d = dateCell.AddParagraph();
                d.AppendText($"Date: {log.DateAdded.LocalDateTime.ToString("MM/dd/yy")}\n");

                if (!string.IsNullOrWhiteSpace(log.Agent?.AgentName ?? ""))
                    d.AppendText($"Agent: {log.Agent?.AgentName}\n");

                if (!string.IsNullOrWhiteSpace(log.ContactChannel))
                    d.AppendText($"Contact Type: {log.ContactChannel}\n");

                //Append Contact Person and Purpose of Contact here as necessary

                var logCell = r.Cells[1];
                var nr = logCell.AddParagraph();

                if (!string.IsNullOrWhiteSpace(log.Title))
                    nr.AppendText($"Title: {log.Title}\n");


                // Check if log.Notes contains HTML content
                if (ContainsHtml(log.Notes))
                {
                    nr.AppendHTML(log.Notes);
                }
                else
                {
                    nr.AppendText(log.Notes);
                }
            }
        }

        private static bool ContainsHtml(string input)
        {
            return !string.IsNullOrEmpty(input) && (input.Contains("<p>") && input.Contains("</p>"));
        }
    }

}