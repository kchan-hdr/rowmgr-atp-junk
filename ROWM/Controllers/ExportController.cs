using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Routing;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using Microsoft.AspNetCore.Routing;
using Microsoft.Extensions.FileProviders;
using ROWM.Dal;
using ROWM.Reports;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ROWM.Controllers
{
    [Produces("application/json")]
    public class ExportController : Controller
    {
        OwnerRepository _repo;
        IFileProvider _file;
        IRowmReports _reports;
        Lazy<string> LogoPath;

        readonly LinkGenerator _links;

        public ExportController(OwnerRepository repo, LinkGenerator g, IFileProvider fileProvider = null, IRowmReports reports = null)
        {
            _repo = repo;
            _file = fileProvider;
            _reports = reports;
            _links = g;

            LogoPath = new Lazy<string>( () => GetLogo());
        }

        #region new reporting engine
        [HttpGet("api/export2")]
        public IEnumerable<ReportDef> GetReportsList()
        {
            var myList = _reports.GetReports();
            return myList.Select( rx => { rx.ReportUrl = $"//{HttpContext.Request.Host.Value}/export2/{rx.ReportCode}"; return rx; });
        }

        [HttpGet("export2/{reportCode}")]
        public async Task<ActionResult> GetReport(string reportCode)
        {
            var m = _reports.GetReports();
            var r = m.FirstOrDefault(x => x.ReportCode == reportCode);
            if (r == null)
                return BadRequest();

            var payload = await _reports.GenerateReport(r);
            return File(payload.Content, payload.Mime, payload.Filename);
        }

        [HttpGet("export/acq")]
        public async Task<IActionResult> DummyReport()
        {
            var m = _reports.GetReports();
            var r = m.FirstOrDefault(x => x.ReportCode == "internal" );
            var payload = await _reports.GenerateReport(r);
            return File(payload.Content, payload.Mime, payload.Filename);
        }
        #endregion


        [Route("parcels/{pid}/logs"), HttpGet]
        public async Task<IActionResult> ExportContactLogs(string pid, [FromServices] SiteDecoration _decoration)
        {
            if (string.IsNullOrWhiteSpace(pid))
                return BadRequest();

            var p = await _repo.GetParcel(pid);
            var o = await _repo.GetOwner(p.Ownership.FirstOrDefault()?.OwnerId ?? default);
            p.ParcelContacts = o?.ContactInfo;

            var h = new Models.ContactLogHelper(_decoration.SiteTitle());
            var docx = await h.GeterateImpl(p);

            return File(docx, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"{pid} Agent Log.docx");
        }

        [Route("parcels/{pid}/relocations/{rcid}/logs"), HttpGet]
        public async Task<IActionResult> ExportDisplaceeContactLogs(string pid, Guid rcid, [FromServices] SiteDecoration _decoration, [FromServices] RelocationRepository _relorepo)
        {
            if (string.IsNullOrWhiteSpace(pid) || rcid == Guid.Empty)
                return BadRequest();

            var p = await _repo.GetParcel(pid);
            var rc = await _relorepo.GetRelocationCaseWithLogs(rcid);
            var o = await _repo.GetOwner(p.Ownership.FirstOrDefault()?.OwnerId ?? default);
            p.ParcelContacts = o?.ContactInfo;

            var h = new Models.ContactLogHelper(_decoration.SiteTitle());
            var docx = await h.GeterateImpl(p, rc);

            return File(docx, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"{rcid} Agent Log.docx");
        }

        [Route("parcels/{pid}/offerPackage"), HttpGet]
        public async Task<IActionResult> ExportOfferPackage(string pid, [FromServices] SiteDecoration _decoration)
        {
            if (string.IsNullOrWhiteSpace(pid))
                return BadRequest();

            var p = await _repo.GetParcel(pid);

            var h = new Models.EasementOfferHelper();
            var docx = await h.GeneratePackage(p);

            return File(docx, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", $"{pid} Easement Offer Package.docx");
        }

        [HttpGet("export/contactlog/{parcelId}")]
        [ProducesDefaultResponseType(typeof(File))]
        public async Task<IActionResult> ContactLogByParcel(string parcelId)
        {
            var p = await _repo.GetParcel(parcelId);
            if (p == null)
                return NoContent();

            var h = new Models.AtpRoeReportHelper();
            var payload = await h.Generate(p);
            return File(payload.Content, payload.Mime, payload.Filename);
        }

        [HttpGet("export/preacq")]
        public async Task<IActionResult> ExportAtpPreAcquisition(string f, [FromServices] ROWM_Context context)
        {
            using (var cmd = context.Database.Connection.CreateCommand())
            {
                cmd.CommandText = "SELECT '1' 'a','1' 'b', * FROM ROWM.rpt_pre_acquisition ORDER BY 3";     // the enginer drops the first 2 columns for other reports.
                cmd.CommandType = System.Data.CommandType.Text;
                await cmd.Connection.OpenAsync();

                var rd = await cmd.ExecuteReaderAsync();

                var rm = new ReportingMethods();
                var stream = rm.StandardReport("Pre-Acquisition Report", rd.FieldCount, LogoPath.Value, rd, false);
                return File(stream
                        , "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        , $"Pre-Acquisiton Report.xlsx");

                //var dt = new System.Data.DataTable();
                //dt.Load(rd);

                //var eng = new ExcelExport.PreAcquisitionExport(Enumerable.Empty<string>(), LogoPath.Value);
                //eng.MyDatatable = dt;

                //var bytes = eng.Export();
                //return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "pre acquisition.xlsx");
            }
        }


        /// <summary>
        /// support excel only
        /// </summary>
        /// <param name="f"></param>
        /// <returns></returns>
        [HttpGet("export/contactlogs")]
        public IActionResult ExportContactLog(string f)
        {
            if ("excel" != f)
                return BadRequest($"not supported export '{f}'");

            var logs = this._repo.GetLogs();
            if (logs.Count() <= 0)
                return NoContent();

            // to do. inject export engine
            try
            {
                var d = logs.SelectMany(lx => lx.Parcel.Select(p =>
                {
                    var rgi = lx.Landowner_Score?.ToString() ?? "";
                    var l = new ExcelExport.AgentLogExport.AgentLog
                    {
                        agentname = lx.Agent.AgentName,
                        contactchannel = lx.ContactChannel,
                        dateadded = lx.DateAdded,
                        notes = lx.Notes?.TrimEnd(',') ?? "",
                        ownerfirstname = p.Ownership.FirstOrDefault()?.Owner.PartyName?.TrimEnd(',') ?? "",
                        ownerlastname = p.Ownership.FirstOrDefault()?.Owner.PartyName?.TrimEnd(',') ?? "",
                        parcelid = p.Assessor_Parcel_Number,
                        parcelstatus = p.Parcel_Status.Description,
                        parcelstatuscode = p.ParcelStatusCode,
                        projectphase = lx.ProjectPhase,
                        roestatus = rgi, // p.Roe_Status.Description,
                        roestatuscode = p.RoeStatusCode,
                        title = lx.Title?.TrimEnd(',') ?? ""
                    };
                    return l;
                }));

                var e = new ExcelExport.AgentLogExport(d, LogoPath.Value);
                var bytes = e.Export();
                return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "logs.xlsx");
            }
            catch (Exception)
            {
                using (var s = new MemoryStream())
                {
                    using (var writer = new StreamWriter(s))
                    {
                        writer.WriteLine(LogExport.Header());

                    foreach (var l in logs.SelectMany(l => LogExport.Export(l)).OrderBy(l => l.Tracking))
                        writer.WriteLine(l);

                        writer.Close();
                    }

                    return File(s.GetBuffer(), "text/csv", "logs.csv");
                }
            }
        }

        /// <summary>
        /// support excel only
        /// </summary>
        /// <param name="f"></param>
        /// <returns></returns>
        [HttpGet("export/documents")]
        public async Task<IActionResult> ExportDocumentg(string f)
        {
            const string DOCUMENT_HEADER = "Parcel CAD,Title,Content Type,Date Sent,Date Delivered,Client Tracking Number,Date Received,Date Signed,Check No,Date Recorded,Document ID";

            if ("excel" != f)
                return BadRequest($"not supported export '{f}'");

            var d = await this._repo.GetDocs();
            if (d.Count() <= 0)
                return NoContent();

            var lines = d.OrderBy(dh => dh.Parcel_ParcelId)
                .Select(dh => $"=\"{dh.Parcel_ParcelId}\",\"{dh.Title}\",{dh.ContentType},{dh.SentDate?.Date.ToShortDateString() ?? ""},{dh.DeliveredDate?.Date.ToShortDateString() ?? ""},{dh.ClientTrackingNumber},{dh.ReceivedDate?.Date.ToShortDateString() ?? ""},{dh.SignedDate?.Date.ToShortDateString() ?? ""},=\"{dh.CheckNo}\",{dh.DateRecorded?.Date.ToShortDateString() ?? ""},=\"{dh.DocumentId}\"");

            using (var s = new MemoryStream())
            {
                using (var writer = new StreamWriter(s))
                {
                    writer.WriteLine(DOCUMENT_HEADER);

                    foreach (var l in lines)
                        writer.WriteLine(l);

                    writer.Close();
                }

                return File(s.GetBuffer(), "text/csv", "documents.csv");
            }
        }

        [HttpGet("export/roe")]
        public IActionResult ExportRoe(string f)
        {
            if ("excel" != f)
                return BadRequest($"not supported export '{f}'");

            var parcels = this._repo.GetParcels2();

            if (parcels.Count() <= 0)
                return NoContent();

            using (var s = new MemoryStream())
            {
                using (var writer = new StreamWriter(s))
                {
                    writer.WriteLine("Parcel CAD,Owner,ROE Status,Conditions,Condition Expires,Date Status Last Modified");

                    foreach (var p in parcels.OrderBy(px => px.Assessor_Parcel_Number))
                    {
                        var os = p.Ownership.OrderBy(ox => ox.IsPrimary() ? 1 : 2).FirstOrDefault();
                        var oname = os?.Owner.PartyName?.TrimEnd(',') ?? "";

                        var conditions = "";
                        var conditionsExp = "";
                        if (p.Conditions != null && p.Conditions.Any())
                        {
                            var cond = p.Conditions?.OrderByDescending(px => px.EffectiveStartDate.HasValue ? px.EffectiveStartDate.Value : DateTime.MinValue).FirstOrDefault();
                            conditions = cond?.Condition ?? "";
                            conditionsExp = cond?.EffectiveStartDate?.DateTime.ToShortDateString() ?? "";
                        }

                        var row = $"{p.Assessor_Parcel_Number},\"{oname}\",{p.Roe_Status.Description},\"{conditions}\",{conditionsExp},{p.LastModified.Date.ToShortDateString()}";
                        
                        writer.WriteLine(row);
                    }

                    writer.Close();
                }

                return File(s.GetBuffer(), "text/csv", "roe.csv");
            }
        }

        [HttpGet("export/roeowner")]
        public IActionResult ExportRoeOwner(string f)
        {
            if ("excel" != f)
                return BadRequest($"not supported export '{f}'");

            var parcels = this._repo.GetRoeOwner();

            if (parcels.Count() <= 0)
                return NoContent();

            var printDate = $"as of {DateTime.Today.ToLongDateString()}";

            var excel = new ExcelExport.RoeOwnerReport();
            var buffer = excel.CreatePackage(printDate, parcels);

            return File(buffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "roe-owner report.xlsx");
        }


        /// <summary>
        /// support excel only
        /// </summary>
        /// <param name="f"></param>
        /// <returns></returns>
        [HttpGet("export/contacts")]
        public async Task< IActionResult> ExportContact(string f)
        {
            if ("excel" != f)
                return BadRequest($"not supported export '{f}'");

            var contacts = await this._repo.GetContacts_AtpReport();
            //var contacts = this._repo.GetContacts();

            //var q = from o in contacts
            //        group o by o.OwnerId into og
            //        select og;

            //var cc = q.SelectMany(og => ContactExport2.Export(og));

            if (contacts.Count() <= 0)
                return NoContent();

            // to do. inject export engine
            try
            {
                var data = contacts
                    .Where(cx => !string.IsNullOrWhiteSpace(cx.ownerfirstname) || !string.IsNullOrWhiteSpace(cx.letter))
                    .OrderBy(cx => cx.partyname)
                    .Select(ccx => new ExcelExport.ContactListExport.ContactList
                    {
                        partyname = ccx.partyname,
                        ownerfirstname = ccx.ownerfirstname,
                        owneremail = ccx.owneremail,
                        ownercellphone = ccx.ownercellphone,
                        ownerhomephone = ccx.ownerhomephone,
                        ownerstreetaddress = ccx.ownerstreetaddress,
                        ownercity = ccx.ownercity,
                        ownerstate = ccx.ownerstate,
                        ownerzip = ccx.ownerzip,
                        representation = ccx.representation,
                        parcelid = string.Join(",", ccx.relatedParcels.OrderBy(r => r)),
                        LetterDate = ccx.letter
                    });

                var e = new ExcelExport.ContactListExport(data, LogoPath.Value);
                var bytes = e.Export();
                return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "contacts.xlsx");
            }
            catch (Exception)
            {
                return BadRequest();
            }
        }

        /// <summary>
        /// support excel only
        /// </summary>
        /// <param name="f"></param>
        /// <returns></returns>
        [HttpGet("export/contacts_i")]
        public IActionResult ExportContactByParcel(string f)
        {
            if ("excel" != f)
                return BadRequest($"not supported export '{f}'");

            var contacts = this._repo.GetContacts();

            var cc = contacts.SelectMany(cx => ContactExport.Export(cx));
            if (cc.Count() <= 0)
                return NoContent();

            using (var s = new MemoryStream())
            {
                using (var writer = new StreamWriter(s))
                {
                    writer.WriteLine(ContactExport.Header());

                    foreach (var l in cc.OrderBy(cx => cx.ParcelId)
                                        .ThenByDescending(cx => cx.IsPrimary)
                                        .ThenBy(cx => cx.LastName)
                                        .Select(ccx => ccx.ToString()))
                        writer.WriteLine(l);

                    writer.Close();
                }

                return File(s.GetBuffer(), "text/csv", "contacts.csv");
            }
        }

        [HttpGet("export/engagement")]
        public async Task<IActionResult> ExportEngagement(string f)
        {
            if ("excel" != f)
                return BadRequest($"not supported export '{f}'");

            var data = await _repo.GetEngagement();
            var e = new ExcelExport.EngagementExport(data);
            var bytes = e.Export();
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "outreach.xlsx");
        }

        #region atp reports
        [HttpGet("export/atp_roe")]
        public async Task<IActionResult> ExportAtpRoe([FromServices] ROWM_Context _ctx, string f)
        {
            using (var command = _ctx.Database.Connection.CreateCommand())
            {
                command.CommandText = "exec dbo.ROEStatusReport";
                await _ctx.Database.Connection.OpenAsync();
                using (var result = command.ExecuteReader())
                {
                    var rm = new ReportingMethods();
                    var stream = rm.StandardReport("ROE Status Report", 15, LogoPath.Value, result, false);
                    return File(stream
                            , "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            , $"ROE Status Report.xlsx");
                }
            }
        }

        [HttpGet("export/atp_all_fields")]
        public async Task<IActionResult> ExportAtpAllMeta([FromServices] ROWM_Context _ctx, string f)
        {
            using (var command = _ctx.Database.Connection.CreateCommand())
            {
                command.CommandText = "exec dbo.AllFieldsReport";
                await _ctx.Database.Connection.OpenAsync();
                using (var result = command.ExecuteReader())
                {
                    var rm = new ReportingMethods();
                    var stream = rm.StandardReport("All Fields Report", 100, LogoPath.Value, result, true);
                    return File(stream
                            , "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            , $"All Fields Report.xlsx");
                }
            }
        }
        #endregion

        #region logo image
        string GetLogo()
        {
            if (_file == null)
                return string.Empty;

            var fileInfo = _file.GetFileInfo("wwwroot/assets/IDP-logo-color.png");
            return  fileInfo.PhysicalPath;
        }
        #endregion
        #region helpers
        public class LogExport
        {
            public string Tracking { get; set; }
            public string ParcelId { get; set; }
            public string ParcelStatusCode { get; set; }
            public string RoeStatusCode { get; set; }
            public string ContactName { get; set; }

            public DateTimeOffset DateAdded { get; set; }
            public string ContactChannel { get; set; }
            public string ProjectPhase { get; set; }
            public string Title { get; set; }
            public string Notes { get; set; }

            public string AgentName { get; set; }

            public static IEnumerable<LogExport> Export(ContactLog log)
            {
                return log.Parcel.Where(p => p.IsActive).Select(p => new LogExport
                {
                    Tracking = p.Tracking_Number,
                    ParcelId = p.Assessor_Parcel_Number,
                    ParcelStatusCode = p.ParcelStatusCode,
                    RoeStatusCode = log.Landowner_Score?.ToString() ?? "", // p.RoeStatusCode,
                    ContactName = p.Ownership.FirstOrDefault()?.Owner.PartyName?.TrimEnd(',') ?? "",
                    DateAdded = log.DateAdded,
                    ContactChannel = log.ContactChannel,
                    ProjectPhase = log.ProjectPhase,
                    Title = log.Title?.TrimEnd(',') ?? "",
                    Notes = log.Notes?.TrimEnd(',') ?? "",
                    AgentName = log.Agent.AgentName
                });
            }

            public static string Header() =>
                "Parcel ID,Parcel Status,Landowner Score,Contact Name,Date,Channel,Type,Title,Notes,Agent Name";

            public override string ToString()
            {
                var n = Notes.Replace('"', '\'');
                return $"=\"{ParcelId}\",{ParcelStatusCode},{RoeStatusCode},\"{ContactName}\",{DateAdded.Date.ToShortDateString()},{ContactChannel},{ProjectPhase},\"{Title}\",\"{n}\",\"{AgentName}\"";
            }
        }

        public class ContactExport2
        {
            public string[] ParcelId { get; set; }


            public string PartyName { get; set; }
            public bool IsPrimary { get; set; }
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string Email { get; set; }
            public string CellPhone { get; set; }
            public string HomePhone { get; set; }
            public string StreetAddress { get; set; }
            public string City { get; set; }
            public string State { get; set; }
            public string ZIP { get; set; }
            public string Representation { get; set; }

            public static IEnumerable<ContactExport2> Export(IGrouping<Guid, Ownership> og)
            {
                var relatedParcels = og.Select(p => $"{p.Parcel.Tracking_Number} ({p.Parcel.Assessor_Parcel_Number})").OrderBy(p => p).ToArray<string>();

                var ox = og.First();
                return ox.Owner.ContactInfo.Select(cx =>  new ContactExport2
                {
                    PartyName = ox.Owner.PartyName?.TrimEnd(',') ?? "",
                    IsPrimary = cx.IsPrimaryContact,
                    FirstName = cx.FirstName?.TrimEnd(',') ?? "",
                    LastName = cx.LastName?.TrimEnd(',') ?? "",
                    Email = cx.Email?.TrimEnd(',') ?? "",
                    CellPhone = cx.CellPhone?.TrimEnd(',') ?? "",
                    HomePhone = cx.HomePhone?.TrimEnd(',') ?? "",
                    StreetAddress = cx.StreetAddress?.TrimEnd(',') ?? "",
                    City = cx.City?.TrimEnd(',') ?? "",
                    State = cx.State?.TrimEnd(',') ?? "",
                    ZIP = cx.ZIP?.TrimEnd(',') ?? "",
                    Representation = cx.Representation,
                    ParcelId = relatedParcels
                });
            }

            string RelatedParcels =>
                string.Join(",", this.ParcelId.Select(p=> $"=\"{p}\""));

            public static string Header() =>
                "Owner,Is Primary Contact,First Name,Last Name,Email,Cell Phone,Phone,Street Address,City,State,ZIP,Representation";

            public override string ToString() =>
                $"\"{PartyName}\",{IsPrimary},\"{FirstName}\",\"{LastName}\",{Email},{CellPhone},{HomePhone},\"{StreetAddress}\",{City},{State},{ZIP},{Representation},{RelatedParcels}";
        }

        public class ContactExport
        {
            public string ParcelId { get; set; }
            public string PartyName { get; set; }
            public bool IsPrimary { get; set; }
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string Email { get; set; }
            public string CellPhone { get; set; }
            public string HomePhone { get; set; }
            public string StreetAddress { get; set; }
            public string City { get; set; }
            public string State { get; set; }
            public string ZIP { get; set; }
            public string Representation { get; set; }

            public static IEnumerable<ContactExport> Export(Ownership op)
            {
                return op.Owner.ContactInfo.Select(cx => new ContactExport
                {
                    ParcelId = op.Parcel.Assessor_Parcel_Number,
                    PartyName = op.Owner.PartyName?.TrimEnd(',') ?? "",
                    IsPrimary = cx.IsPrimaryContact,
                    FirstName = cx.FirstName?.TrimEnd(',') ?? "",
                    LastName = cx.LastName?.TrimEnd(',') ?? "",
                    Email = cx.Email?.TrimEnd(',') ?? "",
                    CellPhone = cx.CellPhone?.TrimEnd(',') ?? "",
                    HomePhone = cx.HomePhone?.TrimEnd(',') ?? "",
                    StreetAddress = cx.StreetAddress?.TrimEnd(',') ?? "",
                    City = cx.City?.TrimEnd(',') ?? "",
                    State = cx.State,
                    ZIP = cx.ZIP,
                    Representation = cx.Representation
                });
            }

            public static string Header() =>
                "Parcel ID,Owner,Is Primary Contact,First Name,Last Name,Email,Cell Phone,Phone,Street Address,City,State,ZIP,Representation";

            public override string ToString() =>
                $"=\"{ParcelId}\",\"{PartyName}\",{IsPrimary},\"{FirstName}\",\"{LastName}\",{Email},{CellPhone},{HomePhone},\"{StreetAddress}\",{City},{State},{ZIP},{Representation}";
        }
        #endregion
    }
}