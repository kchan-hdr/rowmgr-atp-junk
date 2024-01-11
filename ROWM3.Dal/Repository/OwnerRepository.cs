using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace ROWM.Dal
{
    public class OwnerRepository
    {
        #region ctor
        private readonly ROWM_Context _ctx;

        public OwnerRepository(ROWM_Context c) => _ctx = c;
        #endregion

        IQueryable<Parcel> ActiveParcels() => _ctx.Parcel.Where(px => px.IsActive);
        IQueryable<Owner> ActiveOwners() => _ctx.Owner.Where(ox => !ox.IsDeleted);
        IQueryable<ContactInfo> ActiveContacts() => _ctx.ContactInfo.Where(cx => !cx.IsDeleted);
        IQueryable<ContactLog> ActiveLogs() => _ctx.ContactLog.Where(lx => !lx.IsDeleted);
        IQueryable<Document> ActiveDocuments() => _ctx.Document.Where(dx => !dx.IsDeleted);


        public async Task<Owner> GetOwner(Guid uid)
        {
            return await ActiveOwners()
                .Include(ox => ox.Ownership.Select(o => o.Parcel))
                //.Include(ox => ox.ContactLogs)
                //.Include(ox => ox.ContactLogs.Select(ocx => ocx.ContactAgent))
                .Include(ox => ox.ContactInfo)
                .Include(ox => ox.ContactInfo.Select(ocx => ocx.ContactLog))
                .Include(ox => ox.Document)
                .FirstOrDefaultAsync(ox => ox.OwnerId == uid);
        }

        public async Task<IEnumerable<Owner>> FindOwner(string name)
        {
            var q = from owner in ActiveOwners()
                    where owner.PartyName.Contains(name)
                    select new
                    {
                        owner.OwnerId,
                        owner.OwnerAddress,
                        owner.PartyName,
                        owner.OwnerType,
                        Ownershp = owner.Ownership.Select(ox => new
                        {
                            ox.Ownership_t,
                            Parcel = new
                            {
                                ox.Parcel.ParcelId,
                                ox.Parcel.Assessor_Parcel_Number,
                                ox.Parcel.Tracking_Number
                            }
                        })
                    };
            
            return (await q.ToArrayAsync())
                    .Select(ox => new Owner
                    {
                        OwnerId = ox.OwnerId,
                        OwnerAddress = ox.OwnerAddress,
                        PartyName = ox.PartyName,
                        OwnerType = ox.OwnerType,
                        Ownership = ox.Ownershp.Select(osx => new Ownership
                        {
                            Ownership_t = osx.Ownership_t,
                            Parcel = new Parcel
                            {
                                ParcelId = osx.Parcel.ParcelId,
                                Assessor_Parcel_Number = osx.Parcel.Assessor_Parcel_Number,
                                Tracking_Number = osx.Parcel.Tracking_Number
                            }
                        }).ToArray()
                    });
        }

        public async Task<IEnumerable<Owner>> FindOwner_(string name)
        {
            return await ActiveOwners()
                .Include(ox => ox.ContactInfo)
                .Include(ox => ox.ContactInfo.Select(ocx => ocx.ContactLog))
                .Include(ox => ox.Document)
                .Where(ox => ox.PartyName.Contains(name)).ToArrayAsync();
        }

        public async Task<Parcel> GetParcel(string pid)
        {
            var p = await ActiveParcels()
                .Include(px => px.Ownership.Select( o=>o.Owner.ContactLog))
                .Include(px => px.ContactLog)
                .Include(px => px.ActionItems)
                .FirstOrDefaultAsync(px => px.Tracking_Number == pid);

            return p;
        }

        public Task<Parcel[]> GetParcels(IEnumerable<string> pids) =>
            ActiveParcels()
                .Include(px => px.Ownership.Select(o => o.Owner.ContactLog))
                .Include(px => px.ContactLog)
                .Include(px => px.ActionItems)
                .Where(px => pids.Contains(px.Tracking_Number))
                .ToArrayAsync();

        public async Task<List<Document>> GetDocumentsForParcel(string pid)
        {
            var p = await ActiveParcels().FirstOrDefaultAsync(px => px.Tracking_Number.Equals(pid));
            if ( p == null)
            {
                throw new ArgumentOutOfRangeException($"cannot find parcel <{pid}>");
            }

            var query = p.Document.Select(dx => new Document
            {
                DocumentId = dx.DocumentId,
                DocumentType = dx.DocumentType,
                Title = dx.Title,
                DateRecorded = dx.DateRecorded,
                Created = dx.Created,
                LastModified = dx.LastModified,
                IsDeleted = dx.IsDeleted,

                ContentType = dx.ContentType,
                CheckNo = dx.Content?.Length.ToString() ?? "0"
            });

            return query.ToList();

            //var q = _ctx.Database.SqlQuery<DocumentH>("SELECT d.DocumentId, d.DocumentType, d.title FROM rowm.ParcelDocuments pd INNER JOIN rowm.Document d on pd.document_documentid = d.documentid WHERE pd.parcel_parcelId = @pid and d.IsDeleted = 0", new System.Data.SqlClient.SqlParameter("@pid", p.ParcelId));
            //var ds = await q.ToListAsync();
            //return ds.Select(dx => new Document { Title = dx.Title, DocumentId = dx.DocumentId, DocumentType = dx.DocumentType }).ToList();
        }
        #region Db dto
        public class DocumentH
        {
            public Guid DocumentId { get; set; }
            public string DocumentType { get; set; }
            public string Title { get; set; }
        }
        #endregion

        public async Task<IEnumerable<StatusActivity>> GetStatusForParcel(string pid) => await GetStatusForParcel(pid, false);

        public async Task<IEnumerable<StatusActivity>> GetStatusForParcel(string pid, bool all)
        {
            var p = await ActiveParcels().AsNoTracking()
                .Include(px => px.Activities)
                .FirstOrDefaultAsync(px => px.Tracking_Number.Equals(pid));

            if ( p==null)
                throw new KeyNotFoundException($"cannot find parcel <{pid}>");

            if (all)
            {
                return p.Activities.ToArray();
            } 
            else           
            {
                var q = from a in p.Activities
                        group a by a.StatusCode into ag
                        select ag.OrderByDescending(ax => ax.ActivityDate).Take(1);

                return q.SelectMany(qx => qx);
            }
        }
        public IEnumerable<string> GetParcels() => ActiveParcels().AsNoTracking().Select(px => px.Assessor_Parcel_Number);
        public IEnumerable<Parcel> GetParcels2() => ActiveParcels()
            .Include(px => px.Ownership.Select( o => o.Owner ))
            .Include(px => px.Roe_Status)
            .Include(px => px.Conditions).AsNoTracking();

        public IEnumerable<RoeOwnerDto> GetRoeOwner()
        {
            var q = ActiveParcels()
                    .Include(px => px.Ownership.Select(o => o.Owner).Select(o => o.ContactInfo))
                    .Include(px => px.Roe_Status)
                    .Include(px => px.ParcelAllocations.Select(a => a.ProjectPart))
                    .Include(px => px.ContactLog)
                    .Include(px => px.Activities)
                    .Include(px => px.Document)
                    .Select(px => new ParcelDto
                    {
                        Assessor_Parcel_Number = px.Assessor_Parcel_Number,
                        SitusAddress = px.SitusAddress,
                        ParcelAllocations = px.ParcelAllocations,
                        Roe_Status = px.Roe_Status,
                        IsImpacted = px.IsImpacted,
                        Ownership = px.Ownership,
                        ContactLog = px.ContactLog.OrderByDescending(cx => cx.DateAdded).Take(1).ToList(),
                        Activities = px.Activities.ToList(),
                        //Document = px.Document.Where(dx => dx.DocumentType == "Engagement-Community-Engagement-Letter").OrderByDescending(dx => dx.SentDate).Take(1).ToList(),

                        Owner = px.Ownership.OrderBy(ox => ox.Ownership_t == 1 ? 1 : 2).FirstOrDefault().Owner,
                        PrimaryContact = px.Ownership.OrderBy(ox => ox.Ownership_t == 1 ? 1 : 2).FirstOrDefault().Owner.ContactInfo.FirstOrDefault(cx => cx.IsPrimaryContact),
                        Projects = px.ParcelAllocations.Select(pa => pa.ProjectPart.Caption).ToList(),
                        LetterSent = px.Document.Where(dx => dx.DocumentType == "Engagement-Community-Engagement-Letter").Max(dx => dx.SentDate)
                    })
                    .AsNoTracking()
                    .ToList();


            return q.Select(px => new RoeOwnerDto(px));
        }

        public class ParcelDto
        {
            public string Assessor_Parcel_Number { get; set; }
            public string SitusAddress { get; set; }
            public ICollection<ParcelAllocation> ParcelAllocations { get; set; }
            public Roe_Status Roe_Status { get; set; }
            public bool IsImpacted { get; set; }
            public ICollection<Ownership> Ownership { get; set; }
            public ICollection<ContactLog> ContactLog { get; set; }
            public ICollection<StatusActivity> Activities { get; set; }
            public ICollection<Document> Document { get; set; }

            public Owner Owner { get; set; }
            public List<string> Projects { get; set; }
            public ContactInfo PrimaryContact { get; set; }
            public DateTimeOffset? LetterSent { get; set; }
        }
        public class RoeOwnerDto
        {
            public string Apn { get; set; } = "-";
            public string PNum { get; set; } = "-";
            public string Impacted { get; set; }
            public string Projects { get; set; } = "-";
            public string OName { get; set; } = "-";
            public string Contacts { get; set; } = "";
            public string Situs { get; set; } = "";
            public string RoeReq { get; set; } = "";
            public string RoeRec { get; set; } = "";
            public string RoeStatus { get; set; } = "-";
            public string Ltr { get; set; } = "";
            public string LastContact { get; set; } = "-";

            static readonly string[] RECEIVED = new string[] { "ROE_Full_Access", "ROE_Partial_Access" };
            static readonly string REQUESTED = "ROE_Mailed";
            static readonly string SENT = "Owner_Letter_Sent";

            internal RoeOwnerDto(ParcelDto px)
            {
                //var os = px.Ownership.OrderBy(ox => ox.IsPrimary() ? 1 : 2).FirstOrDefault();
                OName = px.Owner?.PartyName?.TrimEnd(',') ?? ""; // os?.Owner.PartyName?.TrimEnd(',') ?? "";

                // Contacts = string.Join("|", os?.Owner.ContactInfo.Select(sx => MakeContactSummary(sx)) ?? new string[] { "-" } );
                //Contacts = string.Join("\n", os?.Owner.ContactInfo.Where(sx => sx.IsPrimaryContact).Select(sx => MakeContactSummary(sx)) ?? new string[] { "-" } );
                Contacts = MakeContactSummary(px.PrimaryContact);

                Projects = string.Join("\n", px.Projects); // px.ParcelAllocations.Select(pa => pa.ProjectPart.Caption));

                Apn = px.Assessor_Parcel_Number;
                Impacted = px.IsImpacted ? "Impacted Parcel" : "Parcel Not Impacted";
                Situs = px.SitusAddress;
                RoeStatus = px.Roe_Status.Description;

                RoeReq = px.Activities.Any(ax => ax.StatusCode.Equals(REQUESTED)) 
                    ? px.Activities.Where(ax => ax.StatusCode.Equals(REQUESTED)).Max(ax => ax.ActivityDate).DateTime.ToShortDateString() 
                    : "-";

                var qx = px.Activities.Where(ax => RECEIVED.Contains(ax.StatusCode));
                RoeRec = qx.Any() ? qx.Max(ax => ax.ActivityDate).DateTime.ToShortDateString() : "-";
                LastContact = px.ContactLog.Any() ? px.ContactLog.Max(cx => cx.DateAdded).Date.ToShortDateString() : "-";

                Ltr = px.LetterSent.HasValue ? px.LetterSent.Value.ToLocalTime().DateTime.ToShortDateString() : "-";
                //Ltr = px.Activities.Any(ax => ax.StatusCode.Equals(SENT))
                //    ? px.Activities.Where(ax => ax.StatusCode.Equals(SENT)).Max(ax => ax.ActivityDate).DateTime.ToShortDateString() : "-";
            }

            private string MakeContactSummary(ContactInfo c)
            {
                if (c == null) return "-";

                var fullname = $"{c.FirstName} {c.LastName}".Trim();
                return fullname;
                //var m = c.Email;
                //var ff = new string[] { c.WorkPhone, c.CellPhone, c.HomePhone };
                //var f = string.Join(",", ff.Where(fx => !string.IsNullOrWhiteSpace(fx)));

                //return string.Join("\n", new string[] { fullname, $"email:{m}", $"phone:{f}" });
            }
        }

        public async Task<Parcel> UpdateParcel (Parcel p)
        {
            if (_ctx.Entry<Parcel>(p).State == EntityState.Detached)
                _ctx.Entry<Parcel>(p).State = EntityState.Modified;

            await WriteDb();
            //if (await WriteDb() <= 0)
            //    throw new ApplicationException("update parcel failed");

            return p;
        }

        public async Task<Owner> AddOwner(string name, string first = "", string last = "", string address = "", string city = "", string state = "", string z = "", string email = "", string hfone = "", string wfone = "", string cfone = "",   bool primary = true )
        {
            var dt = DateTimeOffset.Now;

            var o = _ctx.Owner.Create();
            o.Created = dt;
            o.PartyName = name;
            o.OwnerAddress = MakeAddress(address, city, state, z);

            ///
            /// no longer automatically add a default contact
            /// 
            //var c = _ctx.ContactInfo.Create();
            //c.Created = dt;
            //c.IsPrimaryContact = primary;
            //c.FirstName = first;
            //c.LastName = last;
            //c.StreetAddress = address;
            //c.City = city;
            //c.State = state;
            //c.ZIP = z;
            //c.Email = email;
            //c.HomePhone = hfone;
            //c.CellPhone = cfone;
            //c.WorkPhone = wfone;
            
            //o.ContactInfo = new List<ContactInfo>();
            //o.ContactInfo.Add(c);

            _ctx.Owner.Add(o);

            if (await WriteDb() <= 0)
                throw new ApplicationException("Add owner failed");

            return o;
        }

        static string MakeAddress( string address, string city, string state, string zip)
        {
            char[] trimmer = { ',', ' ' };

            if (string.IsNullOrWhiteSpace(address) && string.IsNullOrWhiteSpace(city) && string.IsNullOrWhiteSpace(state) && string.IsNullOrWhiteSpace(zip))
                return string.Empty;

            return $"{address}, {city} {state} {zip}".Trim(trimmer);
        }

        public async Task<Owner> UpdateOwner(Owner o)
        {
            if (_ctx.Entry<Owner>(o).State == EntityState.Detached)
                _ctx.Entry<Owner>(o).State = EntityState.Modified;

            if (await WriteDb() <= 0)
                throw new ApplicationException("Update owner failed");

            return o;
        }

        public async Task<ContactInfo> UpdateContact(ContactInfo c)
        {
            if (_ctx.Entry<ContactInfo>(c).State == EntityState.Detached)
                _ctx.Entry<ContactInfo>(c).State = EntityState.Modified;

            if (await WriteDb() <= 0)
                throw new ApplicationException("Add owner failed");

            return c;
        }

        public IEnumerable<Ownership> GetContacts() => _ctx.Parcel.Where(p => p.IsActive).SelectMany(p => p.Ownership);

        public async Task<IEnumerable<ContactListDto>> GetContacts_AtpReport()
        {
            var q = from oo in _ctx.Owner.Include(ox => ox.ContactInfo)
                        .Include(ox => ox.Ownership.Select(os => os.Parcel.Document))
                    where oo.Ownership.Any(os => os.Parcel.IsActive)
                    select oo;

            var myOwners = await q.Select(oxx => new
            {
                oxx.PartyName, 
                oxx.ContactInfo,
                Related = oxx.Ownership.Select(p => p.Parcel.Assessor_Parcel_Number),
                Letter = oxx.Ownership.SelectMany(p => p.Parcel.Document.Where(dx => dx.DocumentType == "Engagement-Community-Engagement-Letter").Select(dx => dx.SentDate))
            }).ToArrayAsync();



            //var o = _ctx.Parcel.Where(p => p.IsActive).Include(p => p.Ownership.Select(os => os.Owner).Select(ox => ox.ContactInfo)).Include(p => p.Document).SelectMany(p => p.Ownership);
            //var og = o.GroupBy(ox => ox.OwnerId);
            //var myContacts = await og.Select(oxx => new
            //{
            //    Related = oxx.Select(p => p.Parcel.Assessor_Parcel_Number),
            //    Owner = oxx.FirstOrDefault().Owner,
            //    Letter = oxx.FirstOrDefault().Parcel.Document.Where(dx => dx.DocumentType == "Engagement-Community-Engagement-Letter").Select(dx => dx.SentDate)
            //}).ToArrayAsync();

            var x = myOwners.Select(oxx => {
                var ct = oxx.ContactInfo.OrderByDescending(cx => cx.IsPrimaryContact).FirstOrDefault();
                return new ContactListDto
                {
                    relatedParcels = oxx.Related,
                    letter = oxx.Letter?.Max()?.DateTime.ToShortDateString() ?? "",
                    partyname = oxx.PartyName,
                    ownerfirstname = ct?.FirstName,
                    owneremail = ct?.Email,
                    ownercellphone = ct?.CellPhone,
                    ownerhomephone = ct?.HomePhone,
                    ownerstreetaddress = ct?.StreetAddress,
                    ownercity = ct?.City,
                    ownerstate = ct?.State,
                    ownerzip = ct?.ZIP,
                    representation = ct?.Representation
                };
            });

            return x;
            //var 

            //var ox = og.First();
            //return ox.Owner.ContactInfo.Select(cx => new ContactExport2
            //{
            //    PartyName = ox.Owner.PartyName?.TrimEnd(',') ?? "",
            //    IsPrimary = cx.IsPrimaryContact,
            //    FirstName = cx.FirstName?.TrimEnd(',') ?? "",
            //    LastName = cx.LastName?.TrimEnd(',') ?? "",
            //    Email = cx.Email?.TrimEnd(',') ?? "",
            //    CellPhone = cx.CellPhone?.TrimEnd(',') ?? "",
            //    HomePhone = cx.HomePhone?.TrimEnd(',') ?? "",
            //    StreetAddress = cx.StreetAddress?.TrimEnd(',') ?? "",
            //    City = cx.City?.TrimEnd(',') ?? "",
            //    State = cx.State?.TrimEnd(',') ?? "",
            //    ZIP = cx.ZIP?.TrimEnd(',') ?? "",
            //    Representation = cx.Representation,
            //    ParcelId = relatedParcels
            //});
        }

        #region export dto
        public partial class ContactListDto
        {
            public string partyname { get; set; }
            public int ownership_t { get; set; }
            public bool isprimarycontact { get; set; }
            public string ownerfirstname { get; set; }
            public string ownerlastname { get; set; }
            public string owneremail { get; set; }
            public string ownercellphone { get; set; }
            public string ownerhomephone { get; set; }
            public string ownerstreetaddress { get; set; }
            public string ownercity { get; set; }
            public string ownerstate { get; set; }
            public string ownerzip { get; set; }
            public string representation { get; set; }

            public IEnumerable<string> relatedParcels { get; set; }
            public string letter { get; set; }
        }
        #endregion

        public IEnumerable<ContactLog> GetLogs() => ActiveLogs().Where(c => c.Parcel.Any(p => p.IsActive));
        public async Task<IEnumerable<DocHead>> GetDocs()
        {
            try
            {
                // for performance
                var qstr = "SELECT d.DocumentId, d.Title, d.ContentType, d.ReceivedDate, d.SentDate, d.DeliveredDate, d.SignedDate, d.DateRecorded, d.ClientTrackingNumber, d.CheckNo, p.Assessor_Parcel_Number as 'Parcel_ParcelId' FROM rowm.ParcelDocuments pd INNER JOIN Rowm.Document d on pd.Document_DocumentId = d.DocumentId INNER JOIN Rowm.Parcel p ON pd.Parcel_ParcelId = p.ParcelId where p.IsActive = 1 and d.IsDeleted = 0";
                var q = _ctx.Database.SqlQuery<DocHead>(qstr);

                return await q.ToListAsync();
            }
            catch ( Exception e)
            {
                throw;
            }
        }

        public class DocHead
        {
            public Guid DocumentId { get; set; }
            public string Title { get; set; }
            public string ContentType { get; set; }
            public DateTimeOffset? ReceivedDate { get; set; }
            public DateTimeOffset? SentDate { get; set; }
            public DateTimeOffset? DeliveredDate { get; set; }
            public DateTimeOffset? SignedDate { get; set; }
            public DateTimeOffset? DateRecorded { get; set; }
            public string ClientTrackingNumber { get; set; }
            public string CheckNo { get; set; }
            public string Parcel_ParcelId { get; set; }
        }
        public async Task<ContactLog> AddContactLog(IEnumerable<string> pids, IEnumerable<Guid> cids, ContactLog log)
        {
            var dt = DateTimeOffset.Now;

            _ctx.ContactLog.Add(log);

            if (pids != null && pids.Count() > 0)
            {
                foreach (var pid in pids)
                {
                    var px = await _ctx.Parcel.SingleOrDefaultAsync(pxid => pxid.Assessor_Parcel_Number.Equals(pid) && pxid.IsActive );
                    if (px == null)
                        Trace.TraceWarning($"invalid parcel {pid}");
                    log.Parcel.Add(px);
                }
            }

            if (cids != null && cids.Count() > 0)
            {
                foreach (var cid in cids)
                {
                    var cx = await _ctx.ContactInfo.SingleOrDefaultAsync(oxid => oxid.ContactId.Equals(cid));
                    if (cx == null)
                        Trace.TraceWarning($"invalid contact {cid}");
                    log.ContactInfo.Add(cx);
                }
            }

            if (await WriteDb() <= 0)
                throw new ApplicationException("Add log failed");

            return log;
        }

        public async Task<ContactLog> UpdateContactLog(IEnumerable<string> pids, IEnumerable<Guid> cids, ContactLog log)
        {
            if (_ctx.Entry<ContactLog>(log).State == EntityState.Detached)
            {
                _ctx.Entry<ContactLog>(log).State = EntityState.Modified;
            }

            var existingPids = log.Parcel.Select(p => p.Assessor_Parcel_Number).ToList();
            var existingCids = log.ContactInfo.Select(c => c.ContactId).ToList();

            // Find Deleted & added parcels & contacts
            var deletedPids = existingPids.Except(pids);
            var newPids = pids.Except(existingPids);
            var deletedCids = existingCids.Except(cids);
            var newCids = cids.Except(existingCids);

            // Remove deleted parcels & contacts
            if (deletedPids != null && deletedPids.Count() > 0)
            {
                foreach (var pid in deletedPids)
                {
                    var px = await _ctx.Parcel.SingleOrDefaultAsync(pxid => pxid.Tracking_Number.Equals(pid));
                    if (px == null)
                    {
                        Trace.TraceWarning($"invalid parcel {pid}");
                        continue;
                    }

                    log.Parcel.Remove(px);
                }
            }

            if (deletedCids != null && deletedCids.Count() > 0)
            {
                foreach (var cid in deletedCids)
                {
                    var cx = await _ctx.ContactInfo.SingleOrDefaultAsync(oxid => oxid.ContactId.Equals(cid));
                    if (cx == null)
                    {
                        Trace.TraceWarning($"invalid contact {cid}");
                        continue;
                    }

                    log.ContactInfo.Remove(cx);
                }
            }

            // Add new parcels & contacts
            if (newPids != null && newPids.Count() > 0)
            {
                foreach (var pid in newPids)
                {
                    var px = await _ctx.Parcel.SingleOrDefaultAsync(pxid => pxid.ParcelId.Equals(pid));
                    if (px == null)
                    {
                        Trace.TraceWarning($"invalid parcel {pid}");
                        continue;
                    }
                    log.Parcel.Add(px);
                }
            }

            if (newCids != null && newCids.Count() > 0)
            {
                foreach (var cid in newCids)
                {
                    var cx = await _ctx.ContactInfo.SingleOrDefaultAsync(oxid => oxid.ContactId.Equals(cid));
                    if (cx == null)
                    {
                        Trace.TraceWarning($"invalid contact {cid}");
                        continue;
                    }
                    log.ContactInfo.Add(cx);
                }
            }


            if (await WriteDb() <= 0)
                throw new ApplicationException("update contact log failed");

            return log;
        }

        [Obsolete("use add contactlog")]
        public async Task<Parcel> RecordContact(Parcel p, Agent a, string notes, DateTimeOffset date, string phase)
        {
            var dt = DateTimeOffset.Now;

            var log = _ctx.ContactLog.Create();
            log.Created = dt;
            log.Agent = a;
            log.Notes = notes;
            log.DateAdded = date;
            log.ProjectPhase = phase;

            p.ContactLog.Add(log);

            _ctx.ContactLog.Add(log);

            if (await WriteDb() <= 0)
                throw new ApplicationException("Record Contact failed");

            return p;
        }

        [Obsolete("use add contactlog")]
        public async Task<Owner> RecordOwnerContact(Owner o, Agent a, string notes, DateTimeOffset date, string phase)
        {
            var dt = DateTimeOffset.Now;

            var log = _ctx.ContactLog.Create();
            log.Created = dt;
            log.Agent = a;
            log.Notes = notes;
            log.DateAdded = date;
            log.ProjectPhase = phase;

            o.ContactLog.Add(log);

            _ctx.ContactLog.Add(log);

            if (await WriteDb() <= 0)
                throw new ApplicationException("Record Contact failed");

            return o;
        }

        #region engagement dto
        public async Task<IEnumerable<EngagementDto>> GetEngagement()
        {
            var ppp = await ActiveParcels()
                .Join(_ctx.Parcel_Status, px => px.OutreachStatusCode, sx => sx.Code, (px,sx) => new { px, sx } )
                .Select(p => new EngagementDto
                {
                    Apn = p.px.Assessor_Parcel_Number,
                    IsImpacted = p.px.IsImpacted,
                    TrackingNumber = p.px.Tracking_Number,
                    OwnerName = p.px.Ownership.Select(ox => ox.Owner.PartyName),
                    Contacts = p.px.Ownership.SelectMany(ox => ox.Owner.ContactInfo),
                    Project = p.px.ParcelAllocations.Select(ax => ax.ProjectPart.Caption),
                    OutreachStatus = p.sx.Description,
                    Actions = p.px.ActionItems.Select( ax => new ActionItemHdr
                    {
                        Action = ax.Action,
                        Assigned = ax.AssignedGroup.GroupNameCaption,
                        Due = ax.DueDate,
                        Status = ax.Status
                    }),
                    Logs = p.px.ContactLog.Select(cx => new ContactLogHdr
                    {
                       AgentName = cx.Agent.AgentName,
                       Contacts = cx.ContactInfo,
                       DateAdded = cx.DateAdded,
                       ContactChannel = cx.ContactChannel,
                        ProjectPhase = cx.ProjectPhase,
                       Title = cx.Title,
                       Notes = cx.Notes
                    })
                })
                .ToArrayAsync();

            return ppp;
        }

        public async Task<IEnumerable<EngagementDto>> GetEngagement(int part)
        {
            var ppp = ActiveParcels()
                .Join(_ctx.Parcel_Status, px => px.OutreachStatusCode, sx => sx.Code, (px, sx) => new { px, sx })
                .Select(p => new EngagementDto
                {
                    Apn = p.px.Assessor_Parcel_Number,
                    IsImpacted = p.px.IsImpacted,
                    TrackingNumber = p.px.Tracking_Number,
                    OwnerName = p.px.Ownership.Select(ox => ox.Owner.PartyName),
                    Contacts = p.px.Ownership.SelectMany(ox => ox.Owner.ContactInfo),
                    ProjectCode = p.px.ParcelAllocations.Select(ax => ax.ProjectPartId),
                    Project = p.px.ParcelAllocations.Select(ax => ax.ProjectPart.Caption),
                    OutreachStatus = p.sx.Description,
                    Actions = p.px.ActionItems.Select(ax => new ActionItemHdr
                    {
                        Action = ax.Action,
                        Assigned = ax.AssignedGroup.GroupNameCaption,
                        Due = ax.DueDate,
                        Status = ax.Status
                    }),
                    Logs = p.px.ContactLog.Select(cx => new ContactLogHdr
                    {
                        AgentName = cx.Agent.AgentName,
                        Contacts = cx.ContactInfo,
                        DateAdded = cx.DateAdded,
                        ContactChannel = cx.ContactChannel,
                        ProjectPhase = cx.ProjectPhase,
                        Title = cx.Title,
                        Notes = cx.Notes
                    })
                });
                //.ToArrayAsync();

            var ppx = from par in ppp
                      where par.ProjectCode.Contains(part)
                      select par;

            return await ppx.ToArrayAsync();
        }

        public class EngagementDto
        {
            public string Apn { get; set; }
            public bool IsImpacted { get; set; }
            public string TrackingNumber { get; set; }
            public string OutreachStatus { get; set; }

            public IEnumerable<int> ProjectCode { get; set; }
            public IEnumerable<string> Project { get; set; }
            public IEnumerable<string> OwnerName { get; set; }
            
            public IEnumerable<ContactInfo> Contacts { get; set; }

            public IEnumerable<ContactLogHdr> Logs { get; set; }
            public IEnumerable<ActionItemHdr> Actions { get; set; }

            public string ContactNames
            {
                get
                {
                    return string.Join(" | ", Contacts.Select(cx => $"{cx.FirstName} {cx.LastName}".Trim()));
                }
            }
        }

        public class ContactLogHdr
        { 
            public string AgentName { get; set; }
            public IEnumerable<ContactInfo> Contacts { get; set; }
            public DateTimeOffset DateAdded { get; set; }
            public string ContactChannel { get; set; }
            public string ProjectPhase { get; set; }
            public string Title { get; set; }
            public string Notes { get; set; }

            public string ContactNames {  get
                {
                    return string.Join(" | ", Contacts.Select(cx => $"{cx.FirstName} {cx.LastName}".Trim()));
                } 
            }
        }

        public class ActionItemHdr
        {
            public string Action { get; set; }
            public string Assigned { get; set; }
            public DateTimeOffset? Due { get; set; }
            public ActionStatus Status { get; set; }
        }        
        #endregion

        #region statics lookup
        public async Task<IEnumerable<Parcel_Status>> GetParcelStatus() => await _ctx.Parcel_Status.Where(s => s.IsActive).AsNoTracking().ToListAsync();
        public async Task<IEnumerable<Contact_Purpose>> GetPurpose() => await _ctx.Contact_Purpose.Include(p => p.Milestone).Where(p => p.IsActive).AsNoTracking().ToListAsync();
        #endregion

        #region documents
        public Document GetDocument(Guid id) => _ctx.Document.Find(id);

        public async Task<Document> UpdateDocument(Document d)
        {
            if (_ctx.Entry(d).State == System.Data.Entity.EntityState.Deleted)
                _ctx.Entry(d).State = System.Data.Entity.EntityState.Modified;

            var a = _ctx.DocumentActivity.Create();
            a.Document = d;
            a.ParentDocumentId = d.DocumentId;      //// model-first naming
            a.ActivityDate = DateTimeOffset.Now;
            a.Activity = "Updated Tracking";
            a.Agent = d.Agent.FirstOrDefault();        // TODO:

            _ctx.DocumentActivity.Add(a);

            if (await WriteDb() <= 0)
                throw new ApplicationException("document meta-data edit failed");

            return d;
        }

        public async Task<Document> Store(string title, string document_t, string content_t, string fname, Guid agentId, byte[] content)
        {
            var d = _ctx.Document.Create();
            d.Content = content;
            d.Title = title;
            d.DocumentType = document_t;
            d.SourceFilename = fname;
            d.ContentType = content_t;
            d.Created = DateTimeOffset.Now;

            var a = _ctx.DocumentActivity.Create();
            a.Document = d;
            a.ActivityDate = DateTimeOffset.Now;
            a.Activity = "Uploaded";
            a.AgentId = agentId;
            a.ParentDocumentId = d.DocumentId;

            _ctx.Document.Add(d);
            _ctx.DocumentActivity.Add(a);

            if (await WriteDb() <= 0)
                throw new ApplicationException("document upload failed");

            return d;
        }
        #endregion
        #region row agents
        public async Task<Agent> GetAgent(Guid id)
        {
            var a = await _ctx.Agent.FindAsync(id);
            if (a == null)
                a = await GetDefaultAgent();

            return a;
        }
        public async Task<Agent> GetAgent(string name) => await _ctx.Agent.FirstOrDefaultAsync(ax => ax.AgentName.Equals(name, StringComparison.CurrentCultureIgnoreCase));
        public async Task<Agent> GetDefaultAgent() => await _ctx.Agent.FirstOrDefaultAsync(ax => ax.AgentName.Equals("DEFAULT"));

        public async Task<IEnumerable<Agent>> GetAgents() =>
            await _ctx.Agent.AsNoTracking()
                .Include(ax => ax.ContactLog)
                .ToArrayAsync();
         
        #endregion
        #region helpers
        internal async Task<int> WriteDb()
        {
            if ( _ctx.ChangeTracker.HasChanges())
            {
                try
                {
                    return await _ctx.SaveChangesAsync();
                }
                catch ( Exception e )
                {
                    Trace.TraceError(e.Message);
                    throw;
                }
            }
            else
            {
                return 0;
            }
        }
        #endregion
    }
}
