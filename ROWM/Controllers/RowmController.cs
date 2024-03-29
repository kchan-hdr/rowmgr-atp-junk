using geographia.ags;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using ROWM.Dal;
using ROWM.Dal.Repository;
using ROWM.Models;
using SharePointInterface;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace ROWM.Controllers
{
    [Produces("application/json")]
    [Route("api")]
    public class RowmController : Controller
    {
        static readonly string _APP_NAME = "ROWM";
        static IEnumerable<Parcel_Status> MasterParcelStatus;

        #region ctor
        readonly ROWM_Context _ctx;
        readonly OwnerRepository _repo;
        readonly ParcelStatusRepository _parcelList;
        readonly ContactInfoRepository _contactRepo;
        readonly IStatisticsRepository _statistics;
        readonly DeleteHelper _delete;
        readonly IParcelStatusHelper _statusHelper;
        readonly IUpdateParcelStatus _statusUpdate;
        readonly UpdateParcelStatus2 _updater;
        readonly IFeatureUpdate _featureUpdate;
        readonly ISharePointCRUD _spDocument;
        readonly DocTypes _docTypes;

        public RowmController(ROWM_Context ctx, OwnerRepository r, ParcelStatusRepository l, ContactInfoRepository c, IStatisticsRepository sr, DeleteHelper del, UpdateParcelStatus2 u, IUpdateParcelStatus w, IParcelStatusHelper h, IFeatureUpdate f, ISharePointCRUD s, DocTypes dts)
        {
            _ctx = ctx;
            _repo = r;
            _parcelList = l;
            _contactRepo = c;
            _delete = del;
            _statistics = sr;
            _statusHelper = h;
            _featureUpdate = f;
            _spDocument = s;
            _statusUpdate = w;
            _updater = u;
            _docTypes = dts;
        }
        #endregion
        #region owner
        [Route("owners/{id:Guid}"), HttpGet]
        public async Task<OwnerDto> GetOwner(Guid id) => new OwnerDto(await _repo.GetOwner(id));

        [Route("owners"), HttpGet]
        public async Task<IEnumerable<OwnerDto>> FindOwner(string name)
        {
            return (await _repo.FindOwner(name))
                .Select(ox => new OwnerDto(ox));
        }
        [HttpPut("owners/{id:Guid}")]
        public async Task<ActionResult<OwnerDto>> UpdateOwner(Guid id, [FromBody]OwnerRequest o)
        {
            var ow = await _repo.GetOwner(id);
            if (ow == null)
                return BadRequest();

            ow.PartyName = o.PartyName;
            ow.OwnerType = o.OwnerType;

            ow = await _repo.UpdateOwner(ow);

            return new OwnerDto(ow);
        }
        [HttpPost("parcels/{pid}/owners")]
        public async Task<ActionResult<ParcelGraph>> SetOwner(string pid, [FromBody]OwnerRequest o)
        {
            var p = await _repo.GetParcel(pid);

            var update = new UpdateParcelOwner(this._ctx, p, o.PartyName, o.OwnerType);
            update.ModifiedBy = User?.Identity?.Name ?? _APP_NAME;
            p = await update.Apply();

            return new ParcelGraph(p, _docTypes, await _repo.GetDocumentsForParcel(pid));
        }


        #region contacts
        [Route("owners/{id:Guid}/contacts"), HttpPost]
        public async Task<IActionResult> AddContact(Guid id, [FromBody]ContactRequest info)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            var dt = DateTimeOffset.Now;

            var o = await _repo.GetOwner(id);
            var newc = new ContactInfo
            {
                FirstName = info.OwnerFirstName,
                LastName = info.OwnerLastName,

                StreetAddress = info.OwnerStreetAddress,
                City = info.OwnerCity,
                State = info.OwnerState,
                ZIP = info.OwnerZIP,

                Email = info.OwnerEmail,
                CellPhone = info.OwnerCellPhone,
                WorkPhone = info.OwnerWorkPhone,
                HomePhone = info.OwnerHomePhone,

                IsPrimaryContact = info.IsPrimaryContact,
                Representation = info.Relations,

                Created = dt,
                LastModified = dt,
                ModifiedBy = _APP_NAME
            };
            o.ContactInfo.Add(newc);

            await CheckBusiness(newc, info);

            var newo = await _repo.UpdateOwner(o);

            var sites = _featureUpdate as ReservoirParcel;
            if (sites != null)
            {
                // var apns = o.Ownership.Select(ox => ox.Parcel.Assessor_Parcel_Number);
                var payload = Convert(newc);
                // payload.APN = apns;          no need to denormalize. mobile fixed. geodatabase still broken
                await sites.Update(payload);
            }

            return Json(new OwnerDto(newo));
        }

        [Route("owners/{id:Guid}/contacts/{cinfo}"), HttpPut]
        public async Task<IActionResult> UpdateContact(Guid id, Guid cinfo, [FromBody]ContactRequest info)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            var dt = DateTimeOffset.Now;

            var o = await _repo.GetOwner(id);
            var c = o.ContactInfo.Single(cx => cx.ContactId == cinfo);

            c.FirstName = info.OwnerFirstName;
            c.LastName = info.OwnerLastName;

            c.StreetAddress = info.OwnerStreetAddress;
            c.City = info.OwnerCity;
            c.State = info.OwnerState;
            c.ZIP = info.OwnerZIP;

            c.Email = info.OwnerEmail;
            c.CellPhone = info.OwnerCellPhone;
            c.WorkPhone = info.OwnerWorkPhone;
            c.HomePhone = info.OwnerHomePhone;

            c.IsPrimaryContact = info.IsPrimaryContact;
            c.Representation = info.Relations;

            c.LastModified = dt;
            c.ModifiedBy = _APP_NAME;

            await CheckBusiness(c, info);

            var newc = await _repo.UpdateContact(c);

            var sites = _featureUpdate as ReservoirParcel;
            if (sites != null)
                await sites.Update(Convert(newc));

            return Json(new ContactInfoDto(newc));
        }

        [HttpDelete("contacts/{cid:Guid}")]
        public async Task<ActionResult> DeleteContact(Guid cid)
        {
            if (await _delete.DeleteContact(cid, User.Identity.Name))
                return Ok(cid);
            else
                return BadRequest();
        }

        private async Task<ContactInfo> CheckBusiness(ContactInfo c, ContactRequest r)
        {
            if (string.IsNullOrWhiteSpace(r.BusinessName))
            {
                if (c.OrganizationId.HasValue)
                {
                    c.OrganizationId = null;
                    c.Affiliation = null;
                }

                return c;
            }

            var org = await _contactRepo.FindOrganization(r.BusinessName);
            if (org == null)
            {
                org = new Organization { Name = r.BusinessName };
            }

            if (c.OrganizationId.HasValue)
            {
                if (c.OrganizationId == org.OrganizationId) // no ops. don't do business edits here
                    return c;

                Trace.TraceWarning($"changing affilitation for {c.FirstName} to {org.Name}");
            }

            c.Affiliation = org;
            return c;
        }

        static geographia.ags.ReservoirParcel.ContactInfo_dto Convert(ContactInfo c) =>
            new ReservoirParcel.ContactInfo_dto
            {
                CellPhone = c.CellPhone,
                City = c.City,
                ContactId = c.ContactId.ToString("B"),
                ContactOwnerId = c.ContactOwnerId.ToString("B"),
                Created = c.Created,
                Email = c.Email,
                FirstName = c.FirstName,
                IsPrimaryContact = c.IsPrimaryContact,
                HomePhone = c.HomePhone,
                LastModified = c.LastModified,
                LastName = c.LastName,
                ModifiedBy = c.ModifiedBy,
                Representation = c.Representation,
                State = c.State,
                StreetAddress = c.StreetAddress,
                WorkPhone = c.WorkPhone,
                ZIP = c.ZIP
            };
        #endregion
        #endregion
        #region parcel
        [Route("parcels"), HttpGet]
        public IEnumerable<string> GetAllParcels() => _repo.GetParcels();


        [Route("parcels/{pid}"), HttpGet]
        [ActionName("GetParcel")]
        public async Task<ActionResult<ParcelGraph>> GetParcel(string pid)
        {
            var p = await _repo.GetParcel(pid);
            if (p == null)
                return BadRequest();

            var xxx = _docTypes.Types;

            return Json(new ParcelGraph(p, _docTypes, await _repo.GetDocumentsForParcel(pid)));
        }

        [Route("parcels/{pid}/status"), HttpGet]
        public async Task<ActionResult<ParcelStatusHistoryDto>> GetParcelStatus(string pid)
        {
            var st = await _repo.GetStatusForParcel(pid);
            if (st == null)
                return BadRequest();

            if ( MasterParcelStatus == null )
            {
                MasterParcelStatus = _ctx.Parcel_Status.AsNoTracking().Where(s=> s.IsActive).ToArray();
            }

            var q = from s in MasterParcelStatus
                          join sx in st on s.Code equals sx.StatusCode into h
                          from evt in h.DefaultIfEmpty()
                          orderby s.DisplayOrder
                          select new StatusDto { 
                              Code = s.Code, 
                              Label = s.Description,
                              Category = s.Category,
                              ParentCode = s.ParentStatusCode,
                              DisplayOrder = s.DisplayOrder ?? 0, 
                              IsSet = evt != null,
                              Stage = CodeStage(evt, s),
                              ActivityDate = evt?.ActivityDate.UtcDateTime ?? null,
                              StatusAgent = evt?.AgentId ?? null,
                              StatusHistory = evt?.ActivityId ?? null
                          };

            var history = q.Where(x => x.Code != "No_Activity" ).ToArray();
            foreach( var stage in history.Where(s => !string.IsNullOrEmpty(s.ParentCode) && s.IsSet) )
            {
                var mum = history.FirstOrDefault(sx => sx.Code == stage.ParentCode);
                if (mum == null)
                    throw new IndexOutOfRangeException();

                if (mum.IsSet) continue;    // 
                mum.Stage = StatusDto.StageCode.InProgress.ToString();
            }

            foreach ( var milestone in MasterParcelStatus.Where(s => (s.IsAbort == true) || (s.IsComplete == true) ))
            {
                var child = history.FirstOrDefault(sx => sx.Code == milestone.Code);
                if (child != null && child.Stage != StatusDto.StageCode.Pending.ToString())
                {
                    var mum = history.FirstOrDefault(sx => sx.Code == child.ParentCode);
                    if (mum == null)
                        throw new IndexOutOfRangeException();

                    if (mum.IsSet) continue;    // 

                    mum.Stage = (milestone.IsComplete == true) ? StatusDto.StageCode.Completed.ToString()
                        : (milestone.IsAbort == true) ? StatusDto.StageCode.Aborted.ToString()
                        : StatusDto.StageCode.InProgress.ToString();
                }
            }

            var current = st.OrderByDescending(sx => sx.ActivityDate).FirstOrDefault()?.StatusCode ?? "";
            var status = MasterParcelStatus.SingleOrDefault(sx => sx.Code == current) ?? MasterParcelStatus.First();
            var statusCode = string.IsNullOrEmpty(status.ParentStatusCode) ? status.Code : status.ParentStatusCode;

            return new ParcelStatusHistoryDto
            {
                ParcelUrl = Url.Action("GetParcel", new { pid }),
                Status = statusCode,
                History = history
            };
        }

        string CodeStage (StatusActivity evt, Parcel_Status s)
        {
            if (evt?.ActivityDate == null)
                return StatusDto.StageCode.Pending.ToString();


            if (true == s.IsComplete)
                return StatusDto.StageCode.Completed.ToString();
            if (true == s.IsAbort)
                return StatusDto.StageCode.Aborted.ToString();

            return StatusDto.StageCode.InProgress.ToString();
        }
        
        [HttpGet("status/{milestone}/parcels")]
        public async Task<ActionResult<IEnumerable<ParcelHistory>>> GetParcelsByStatus(string milestone)
        {
            if (_parcelList == null)
                return NoContent();

            var keys = await _parcelList.GetStages(milestone);
            var parcels = await _parcelList.GetParcelList(milestone);

            foreach( var p in parcels)
            {
                p.ParcelUrl = Url.Action("GetParcel", new { pid = p.Tracking_Number });
            }

            return parcels.ToArray();
        }

        #region offer
        [Route("parcels/{pid}/initialOffer"), HttpPut]
        [ProducesResponseType(typeof(ParcelGraph), 202)]
        public async Task<IActionResult> UpdateInitialOffer(string pid, [FromBody]OfferRequest offer)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            var offer_t = offer.OfferType?.Trim().ToLower() ?? "";

            var p = await _repo.GetParcel(pid);
            switch (offer_t)
            {
                case "roe":
                    p.InitialROEOffer_OfferDate = offer.OfferDate;
                    p.InitialROEOffer_OfferAmount = offer.Amount;
                    p.InitialROEOffer_OfferNotes = offer.Notes;
                    break;
                case "option":
                    p.InitialOptionOffer_OfferDate = offer.OfferDate;
                    p.InitialOptionOffer_OfferAmount = offer.Amount;
                    p.InitialOptionOffer_OfferNotes = offer.Notes;
                    break;
                case "easement":
                    p.InitialEasementOffer_OfferDate = offer.OfferDate;
                    p.InitialEasementOffer_OfferAmount = offer.Amount;
                    p.InitialEasementOffer_OfferNotes = offer.Notes;
                    break;
                default:
                    return BadRequest($"Unknown offer type '{offer_t}'");
            }

            p.LastModified = DateTimeOffset.Now;
            p.ModifiedBy = _APP_NAME;

            return Ok(new ParcelGraph(await _repo.UpdateParcel(p), _docTypes, await _repo.GetDocumentsForParcel(pid)));
        }

        [Route("parcels/{pid}/finalOffer"), HttpPut]
        [ProducesResponseType(typeof(ParcelGraph), 202)]
        public async Task<IActionResult> UpdateFinalOffer(string pid, [FromBody]OfferRequest offer)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            var offer_t = offer.OfferType?.Trim().ToLower() ?? "";

            var p = await _repo.GetParcel(pid);

            switch (offer_t)
            {
                case "roe":
                    p.FinalROEOffer_OfferDate = offer.OfferDate;
                    p.FinalROEOffer_OfferAmount = offer.Amount;
                    p.FinalROEOffer_OfferNotes = offer.Notes;
                    break;
                case "option":
                    p.FinalOptionOffer_OfferDate = offer.OfferDate;
                    p.FinalOptionOffer_OfferAmount = offer.Amount;
                    p.FinalOptionOffer_OfferNotes = offer.Notes;
                    break;
                case "easement":
                    p.FinalEasementOffer_OfferDate = offer.OfferDate;
                    p.FinalEasementOffer_OfferAmount = offer.Amount;
                    p.FinalEasementOffer_OfferNotes = offer.Notes;
                    break;
                default:
                    return BadRequest($"Unknown offer type '{offer_t}'");
            }

            p.LastModified = DateTimeOffset.Now;
            p.ModifiedBy = _APP_NAME;

            return Json(new ParcelGraph(await _repo.UpdateParcel(p), _docTypes, await _repo.GetDocumentsForParcel(pid)));
        }
        #endregion
        #region parcel status
        [HttpPut("parcels/{pid}/status/{statusCode}")]
        public async Task<ActionResult<ParcelGraph>> UpdateStatus(string pid, string statusCode)
        {
            var p = await _repo.GetParcel(pid);
            if (p == null)
                return BadRequest();

            var a = await _repo.GetDefaultAgent();
            var ud = new UpdateParcelStatus(new Parcel[] { p }, a, context: _ctx, _repo, _featureUpdate, _statusHelper)
            {
                AcquisitionStatus = statusCode,
                ModifiedBy = User?.Identity?.Name ?? _APP_NAME
            };

            await ud.Apply();

            return new ParcelGraph(p, _docTypes, await _repo.GetDocumentsForParcel(pid));
        }
        [HttpPut("parcels/{pid}/status")]
        public async Task<ActionResult<ParcelGraph>> UpdateStatus(string pid, [FromBody] AcqRequest request)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            var p = await _repo.GetParcel(pid);
            if (p == null)
                return BadRequest();

            var a = await _repo.GetAgent(request.AgentId);

            this._statusUpdate.myParcels = new[] { p };
            this._statusUpdate.myAgent = a;
            this._statusUpdate.AcquisitionStatus = request.StatusCode;
            this._statusUpdate.StatusChangeDate = request.ChangeDate;
            this._statusUpdate.Notes = request.Notes;
            this._statusUpdate.ModifiedBy = User?.Identity?.Name ?? _APP_NAME;
            await this._statusUpdate.Apply();
            
            //var update = new UpdateParcelStatus(new[] { p }, a, this._ctx, this._repo, this._featureUpdate, this._statusHelper);
            //update.AcquisitionStatus = request.StatusCode;
            //update.Notes = request.Notes;
            //update.ModifiedBy = User?.Identity?.Name ?? _APP_NAME;
            //await update.Apply();
            
            return new ParcelGraph(p, _docTypes, await _repo.GetDocumentsForParcel(pid));
        }
        #endregion
        #region roe status
        [HttpPut("parcels/{pid}/roe/{statusCode}")]
        public async Task<ActionResult<ParcelGraph>> UpdateRoeStatus(string pid, string statusCode) => await UpdateRoeStatusImpl(pid, Guid.Empty, statusCode, null, DateTimeOffset.UtcNow, null, null);

        [HttpPut("parcels/{pid}/roe")]
        public async Task<ActionResult<ParcelGraph>> UpdateRoeStatus2(string pid, [FromBody] RoeRequest r) => await UpdateRoeStatusImpl(pid, r.AgentId, r.StatusCode, r.Condition, r.ChangeDate, r.EffectiveStart, r.EffectiveEnd);

        private async Task<ActionResult<ParcelGraph>> UpdateRoeStatusImpl(string pid, Guid agentId, string statusCode, string condition, DateTimeOffset changeDate, DateTimeOffset? startDate, DateTimeOffset? endDate)
        {
            var p = await _repo.GetParcel(pid);
            if (p == null)
                return BadRequest();

            var a = await _repo.GetAgent(agentId);
            var ud = new UpdateParcelStatus(new Parcel[] { p }, a, context: _ctx, _repo, _featureUpdate, _statusHelper)
            {
                RoeCondition = condition,
                RoeStatus = statusCode,
                ModifiedBy = User?.Identity?.Name ?? _APP_NAME,

                StatusChangeDate = changeDate,
                ConditionStartDate = startDate,
                ConditionEndDate = endDate
            };

            await ud.Apply();

            return new ParcelGraph(p, _docTypes, await _repo.GetDocumentsForParcel(pid));
        }
        #endregion
        #region community engagement
        [HttpPost("parcel/{pid}/outreach")]
        public async Task<ActionResult<ParcelGraph>> UpdateOutreach(string pid, [FromBody] OutreachRequest r)
        {
            if (!ModelState.IsValid)
                return BadRequest();

            var p = await _repo.GetParcel(pid);
            if (p == null)
                return BadRequest();

            //
            if (p.OutreachStatusCode != r.StatusCode)
            {
                var dt = DateTimeOffset.UtcNow;

                var act = new StatusActivity
                {
                    ActivityDate = r.ChangeDate,
                    AgentId = r.AgentId,
                    ParentParcelId = p.ParcelId,
                    StatusCode = r.StatusCode,
                    OriginalStatusCode = p.OutreachStatusCode
                };
                _ctx.Activities.Add(act);

                if (!string.IsNullOrWhiteSpace(r.Action))
                {
                    var action = new ActionItem
                    {
                        Action = r.Action,
                        DueDate = r.DueDate.Value,
                        ParentActivity = act,
                        Status = ActionStatus.Pending,

                        Created = dt,
                        LastModified = dt,
                        ModifiedBy = User?.Identity?.Name ?? _APP_NAME
                    };
                    _ctx.ActionItem.Add(action);
                }

                p.OutreachStatusCode = r.StatusCode;
                p.LastModified = dt;
                p.ModifiedBy = User?.Identity?.Name ?? _APP_NAME;

                var dv = _statusHelper.GetDomainValue(r.StatusCode);

                try
                {
                    await Task.WhenAll(
                        _featureUpdate.UpdateFeatureOutreach(pid, p.Tracking_Number, dv, r.Action, r.DueDate),
                        _ctx.SaveChangesAsync());
                }
                catch ( Exception e)
                {
                    Trace.TraceError(e.Message);
                }
            }

            return new ParcelGraph(p, _docTypes, await _repo.GetDocumentsForParcel(pid));
        }
        #endregion
        #region score
        [Route("parcels/{pid}/rating/{score}"), HttpPut]
        public async Task<ActionResult<ParcelGraph>> UpdateParcelScore(string pid, int score)
        {
            await UpdateLandownerScore(score, DateTimeOffset.Now, new[] { pid });
            var p = await _repo.GetParcel(pid);
            return Json(new ParcelGraph(p, _docTypes, await _repo.GetDocumentsForParcel(pid)));
        }
        #endregion
        #endregion
        #region logs
        [Route("parcels/{pid}/logs"), HttpPost]
        public async Task<IActionResult> AddContactLog(string pid, [FromBody] LogRequest logRequest)
        {
            if (!ModelState.IsValid)
                return BadRequest(this.ModelState);

            var dt = DateTimeOffset.Now;

            var p = await _repo.GetParcel(pid);
            var a = await _repo.GetAgent(logRequest.AgentName);

            logRequest.ParcelIds.Add(pid);
            var myParcels = logRequest.ParcelIds.Distinct();
            if (!await RecordParcelFirstContact(myParcels))
            {
                Trace.TraceWarning($"AddContactLog:: update feature status for '{pid}' failed");
            }

            // await UpdateLandownerScore(logRequest.Score, dt, myParcels);

            var l = new ContactLog
            {
                Agent = a,
                ContactChannel = logRequest.Channel,
                ProjectPhase = logRequest.Phase,
                DateAdded = logRequest.DateAdded,
                Title = logRequest.Title,
                Notes = logRequest.Notes,
                Created = dt,
                LastModified = dt,
                ModifiedBy = _APP_NAME,
                Parcel = new List<Parcel> { p },
                ContactInfo = new List<ContactInfo>(),
                Landowner_Score = logRequest.Score
            };

            var log = await _repo.AddContactLog(myParcels, logRequest.ContactIds, l);

            var pxs = await _repo.GetParcels(myParcels);
            _statusUpdate.myParcels = pxs;
            //var ptasks = myParcels.Select(px => _repo.GetParcel(px));
            //_statusUpdate.myParcels = await Task.WhenAll(ptasks);
            _statusUpdate.myAgent = a;
            _statusUpdate.StatusChangeDate = logRequest.DateAdded;
            await _statusUpdate.Apply();

            // var touched = await PromptStatus(myParcels);

            var sites = _featureUpdate as ReservoirParcel;
            if (sites != null)
            {
                var logx = Convert(log);
                logx.APN = myParcels.ToArray();
                await sites.Update(logx);
            }

            if (!string.IsNullOrWhiteSpace(logRequest.MapExportUrl))
            {
                var _helper = new FileAttachmentHelper(_repo, _spDocument);
                await _helper.Attach(log, logRequest.ParcelIds, logRequest.MapExportUrl);
            }
            return Json(new ContactLogDto(log));
        }
        [Route("parcels/{pid}/logs/{lid}"), HttpPost]
        public async Task<IActionResult> UpdateContactLog(string pid, Guid lid, [FromBody] LogRequest logRequest)
        {
            if (!ModelState.IsValid)
                return BadRequest(this.ModelState);

            var dt = DateTimeOffset.Now;

            var p = await _repo.GetParcel(pid);
            var a = await _repo.GetAgent(logRequest.AgentName);
            var l = p.ContactLog.Single(cx => cx.ContactLogId == lid);

            l.Agent = a;
            l.ContactChannel = logRequest.Channel;
            l.ProjectPhase = logRequest.Phase;
            l.DateAdded = logRequest.DateAdded;
            l.Title = logRequest.Title;
            l.Notes = logRequest.Notes;
            l.Landowner_Score = logRequest.Score;
            l.Created = dt;
            l.LastModified = dt;
            l.ModifiedBy = _APP_NAME;
            //l.Parcels = new List<Parcel> { p };
            //l.Contacts = new List<ContactInfo>();

            logRequest.ParcelIds.Add(pid);
            var myParcels = logRequest.ParcelIds.Distinct();
            await UpdateLandownerScore(logRequest.Score, dt, myParcels);

            var log = await _repo.UpdateContactLog(myParcels, logRequest.ContactIds, l);
            var touched = await PromptStatus(myParcels);

            var sites = _featureUpdate as ReservoirParcel;
            if (sites != null)
            {
                var logx = Convert(log);
                logx.APN = logRequest.ParcelIds.ToArray();
                await sites.Update(logx);
            }
            return Json(new ContactLogDto(log));
        }

        async Task<int> PromptStatus(IEnumerable<string> pids)
        {
            // way too convoluted. TODO: 2020.3.24
            //

            var touched = 0;
            foreach (var p in pids)
            {
                var parcel = await _repo.GetParcel(p);
                var (t, s) = await _updater.DoUpdate(parcel);
                if (t)
                {
                    touched++;
                    await UpdateStatus(p, s.Code);
                }
            }

            return touched;
        }

        static ReservoirParcel.ContactLog_dto Convert(ContactLog log)
        {
            return new ReservoirParcel.ContactLog_dto
            {
                ContactChannel = log.ContactChannel,
                ContactLogId = log.ContactLogId.ToString("B"),
                ModifiedBy = log.ModifiedBy,
                Notes = log.Notes,
                ProjectPhase = log.ProjectPhase,
                Title = log.Title
            };
        }

        [HttpDelete("contactLogs/{lid:Guid}")]
        public async Task<ActionResult> DeleteContactLog(Guid lid)
        {
            if (await _delete.DeleteContactLog(lid, User.Identity.Name))
                return Ok(lid);
            else
                return BadRequest();
        }
        #region contact status helper
        /// <summary>
        /// update db and feature to Owner_Contacted, if was No_Activities
        /// </summary>
        /// <remarks>group all parcelIds into a single ags rest call</remarks>
        /// <param name="parcelIds"></param>
        /// <returns></returns>
        async Task<bool> RecordParcelFirstContact(IEnumerable<string> parcelIds)
        {
            var good = true;
            //var tasks = new List<Task>();
            //foreach( var pid in parcelIds)
            //{
            //    var p = await _repo.GetParcel(pid);
            //    if ( ParcelStatusHelper.HasNoContact(p))
            //    {
            //        p.ParcelStatusCode = "Owner_Contacted";
            //        //p.ParcelStatus = Parcel.RowStatus.Owner_Contacted;

            //        tasks.Add(_featureUpdate.UpdateFeature(p.Assessor_Parcel_Number, 1));
            //    }
            //}
            //await Task.WhenAll(tasks);

            return good;
        }
        #endregion
        #region landowner score
        /// <summary>
        /// trigger feature update
        /// </summary>
        /// <remarks>this is fairly wasteful. TODO: check if project uses Landowner Score</remarks>
        /// <param name="score"></param>
        /// <param name="ts"></param>
        /// <param name="parcelIds"></param>
        /// <returns></returns>
        async Task<int> UpdateLandownerScore(int score, DateTimeOffset ts, IEnumerable<string> parcelIds)
        {
            var touched = 0;
            if (score >= 0 && score <= 2)
            {
                var tasks = new List<Task>();

                foreach (var pid in parcelIds)
                {
                    var p = await _repo.GetParcel(pid);
                    var oldValue = p.Landowner_Score;
                    if (oldValue != score)
                    {
                        p.Landowner_Score = score;
                        p.LastModified = ts;
                        p.ModifiedBy = _APP_NAME;
                        touched++;

                        tasks.Add(_featureUpdate.UpdateRating(p.Assessor_Parcel_Number, p.Tracking_Number, score));
                        tasks.Add(_repo.UpdateParcel(p));
                    }
                }

                await Task.WhenAll(tasks);
            }
            return touched;
        }
        #endregion
        #endregion
        #region agents
        [Route("agents"), HttpGet]
        public async Task<IEnumerable<AgentDto>> GetAgents()
        {
            var a = await _repo.GetAgents();
            return a.Where(ax => ax.IsActive).Select(ax => new AgentDto(ax));
        }
        #endregion
        #region statistics
        [HttpGet("statistics_no_color")]
        public async Task<StatisticsDto> GetStatistics()
        {
            var s = await _statistics.Snapshot();
            return new StatisticsDto
            {
                NumberOfOwners = s.nOwners,
                NumberOfParcels = s.nParcels,
                ParcelStatus = await _statistics.SnapshotParcelStatus(),
                RoeStatus = await _statistics.SnapshotRoeStatus(),
                ClearStatus = await _statistics.SnapshotClearanceStatus(),
                Access = await _statistics.SnapshotAccessLikelihood()
            };
        }

        [HttpGet("financials")]
        public async Task<StatisticsRepository.Financials> GetFinancials() => await _statistics.GetFinancials();

        //[HttpGet("v2/statistics/{partId}")]
        //public async Task<StatisticsDto> GetStatistics(int partId)
        //{
        //    if (partId <= 0)
        //        return await GetStatistics();

        //    var s = await _statistics.Snapshot(partId);
        //    return new StatisticsDto
        //    {
        //        NumberOfOwners = s.nOwners,
        //        NumberOfParcels = s.nParcels,
        //        ParcelStatus = await _statistics.SnapshotParcelStatus(partId),
        //        RoeStatus = await _statistics.SnapshotRoeStatus(partId),
        //        Access = await _statistics.SnapshotAccessLikelihood(partId)
        //    };
        //}
        #endregion
    }

    #region request dto
    public class LogRequest
    {
        public Guid AgentId { get; set; }
        public string AgentName { get; set; }
        public string Phase { get; set; }
        public List<string> ParcelIds { get; set; }
        public List<Guid> ContactIds { get; set; }
        public string Channel { get; set; }
        public DateTimeOffset DateAdded { get; set; }
        public string Title { get; set; }
        public string Notes { get; set; }
        public string MapExportUrl { get; set; }
        public int Score { get; set; }
    }

    public class ContactRequest
    {
        public string OwnerFirstName { get; set; } = "";
        public string OwnerLastName { get; set; } = "";

        public string OwnerStreetAddress { get; set; } = "";
        public string OwnerCity { get; set; } = "";
        public string OwnerState { get; set; } = "";
        public string OwnerZIP { get; set; } = "";

        public string OwnerEmail { get; set; } = "";
        public string OwnerCellPhone { get; set; } = "";
        public string OwnerWorkPhone { get; set; } = "";
        public string OwnerHomePhone { get; set; } = "";

        public bool IsPrimaryContact { get; set; } = false;
        public string Relations { get; set; } = "";

        public string BusinessName { get; set; } = "";
    }

    public class OwnerRequest
    {
        public string PartyName { get; set; }
        public string OwnerType { get; set; }
    }

    public class AcqRequest
    {
        public Guid AgentId { get; set; }
        public string StatusCode { get; set; }
        public DateTimeOffset ChangeDate { get; set; }
        public string Notes { get; set; }
    }
    public class RoeRequest
    {
        public Guid AgentId { get; set; }
        public string StatusCode { get; set; }
        public DateTimeOffset ChangeDate { get; set; }
        public DateTimeOffset? EffectiveStart { get; set; }
        public DateTimeOffset? EffectiveEnd { get; set; }
        public string Condition { get; set; }
    }

    public class OutreachRequest
    {
        public Guid AgentId { get; set; }
        public string StatusCode { get; set; }
        public DateTimeOffset ChangeDate { get; set; }
        public DateTimeOffset? DueDate { get; set; }
        public string Action { get; set; }
    }
    #endregion
    #region offer dto
    public class OfferRequest
    {
        public string OfferType { get; set; }   // ROE, Option, or Easement
        public DateTimeOffset OfferDate { get; set; }
        public double Amount { get; set; }
        public string Notes { get; set; }
    }
    #endregion
    #region dto
    public class StatisticsDto
    {
        public int NumberOfParcels { get; set; }
        public int NumberOfOwners { get; set; }

        public IEnumerable<StatisticsRepository.SubTotal> ParcelStatus { get; set; }
        public IEnumerable<StatisticsRepository.SubTotal> RoeStatus { get; set; }
        public IEnumerable<StatisticsRepository.SubTotal> ClearStatus { get; set; }
        public IEnumerable<StatisticsRepository.SubTotal> Access { get; set; }
        public IEnumerable<StatisticsRepository.SubTotal> Compensations { get; set; }
    }

    public class AgentDto
    {
        public Guid AgentId { get; set; }
        public string AgentName { get; set; }
        public IEnumerable<ContactLogDto> ContactsLog { get; set; }
        public IEnumerable<DocumentHeader> Documents { get; set; }

        internal AgentDto(Agent a)
        {
            AgentId = a.AgentId;
            AgentName = a.AgentName;
            ContactsLog = a.ContactLog.Select(cx => new ContactLogDto(cx));
            // Documents = a.Documents.Select(dx => new DocumentHeader(dx));
        }
    }
    public class ContactLogDto
    {
        public Guid ContactLogId { get; set; }
        public IEnumerable<RelatedParcelHeader> ParcelIds { get; set; }
        public IEnumerable<ContactInfoDto> ContactIds { get; set; }
        public DateTimeOffset DateAdded { get; set; }
        public string ContactType { get; set; }
        public string AgentName { get; set; }
        public string Phase { get; set; }
        public string Title { get; set; }
        public string Notes { get; set; }
        public int Score { get; set; }

        internal ContactLogDto(ContactLog log)
        {
            ContactLogId = log.ContactLogId;
            AgentName = log.Agent?.AgentName ?? "";
            DateAdded = log.DateAdded.LocalDateTime;
            ContactType = log.ContactChannel;

            ParcelIds = log.Parcel.Select(px => new RelatedParcelHeader(px));
            ContactIds = log.ContactInfo.Select(cx => new ContactInfoDto(cx));

            Phase = log.ProjectPhase;
            Title = log.Title;
            Notes = log.Notes;
            Score = log.Landowner_Score ?? 0;
        }
    }

    public class RelatedParcelHeader
    {
        public Guid ParcelId { get; set; }
        public string APN { get; set; }
        public string Tracking { get; set; }
        public RelatedParcelHeader(Parcel p)
        {
            this.ParcelId = p.ParcelId;
            this.APN = p.Assessor_Parcel_Number;
            this.Tracking = p.Tracking_Number;
        }
    }

    public class OwnerHeader
    {
        public Guid OwnerId { get; set; }
        public string PartyName { get; set; }
    }

    public class ContactInfoDto
    {
        public Guid ContactId { get; set; }
        public string ContactName { get; set; }
        public bool IsPrimary { get; set; }

        public string Relations { get; set; }

        public string OwnerFirstName { get; set; }
        public string OwnerLastName { get; set; }
        public string OwnerStreetAddress { get; set; }
        public string OwnerCity { get; set; }
        public string OwnerState { get; set; }
        public string OwnerZIP { get; set; }
        public string OwnerEmail { get; set; }
        public string OwnerHomePhone { get; set; }
        public string OwnerCellPhone { get; set; }
        public string OwnerWorkPhone { get; set; }

        public string BusinessName { get; set; }

        internal ContactInfoDto(ContactInfo c)
        {
            ContactId = c.ContactId;
            ContactName = $"{c.FirstName ?? ""} {c.LastName ?? ""}".Trim();
            IsPrimary = c.IsPrimaryContact;
            Relations = c.Representation;

            OwnerFirstName = c.FirstName;
            OwnerLastName = c.LastName;
            OwnerStreetAddress = c.StreetAddress;
            OwnerCity = c.City;
            OwnerState = c.State;
            OwnerZIP = c.ZIP;
            OwnerEmail = c.Email;
            OwnerCellPhone = c.CellPhone;
            OwnerWorkPhone = c.WorkPhone;
            OwnerHomePhone = c.HomePhone;

            if (c.Affiliation != null)
                BusinessName = c.Affiliation.Name;
        }
    }

    public class OwnerDto
    {
        public Guid OwnerId { get; set; }
        public string PartyName { get; set; }
        public string OwnerAddress { get; set; }
        public string OwnerType { get; set; }
        public int OwnershipType { get; set; }
        public IEnumerable<ParcelHeaderDto> OwnedParcel { get; set; }
        public IEnumerable<ContactInfoDto> Contacts { get; set; }
        public IEnumerable<ContactLogDto> ContactLogs { get; set; }
        public IEnumerable<DocumentHeader> Documents { get; set; }

        public OwnerDto(Owner o, int oType = 1)
        {
            OwnerId = o.OwnerId;
            PartyName = o.PartyName;
            OwnerAddress = o.OwnerAddress;
            OwnerType = o.OwnerType;
            OwnershipType = oType;

            OwnedParcel = o.Ownership.Where(ox => ox.Parcel.IsActive).Select(ox => new ParcelHeaderDto(ox));
            Contacts = o.ContactInfo.Where(cx => !cx.IsDeleted).Select(cx => new ContactInfoDto(cx));
            ContactLogs = o.ContactInfo
                .Where(cx => cx.ContactLog != null)
                .SelectMany(cx => cx.ContactLog.Select(cxl => new ContactLogDto(cxl))); //  o.ContactLogs.Select(cx => new ContactLogDto(cx));
            Documents = o.Document.Where(dx => !dx.IsDeleted).Select(dx => new DocumentHeader(dx));
        }
    }

    public class ParcelHeaderDto
    {
        public string ParcelId { get; set; }
        public string TractNo { get; set; }
        public string SitusAddress { get; set; }
        public bool IsPrimaryOwner { get; set; }

        public Compensation InitialROEOffer { get; set; }
        public Compensation FinalROEOffer { get; set; }
        public Compensation InitialOptionOffer { get; set; }
        public Compensation FinalOptionOffer { get; set; }
        public Compensation InitialEasementOffer { get; set; }
        public Compensation FinalEasementOffer { get; set; }

        internal ParcelHeaderDto(Ownership o)
        {
            ParcelId = o.Parcel.Assessor_Parcel_Number;
            TractNo = o.Parcel.Tracking_Number;
            SitusAddress = o.Parcel.SitusAddress;
            IsPrimaryOwner = o.IsPrimary(); // .Ownership_t == Ownership.OwnershipType.Primary;

            InitialEasementOffer = OfferHelper.MakeCompensation(o.Parcel, "InitialEasement");
            InitialOptionOffer = OfferHelper.MakeCompensation(o.Parcel, "InitialOption");
            InitialROEOffer = OfferHelper.MakeCompensation(o.Parcel, "InitialROE");
            FinalEasementOffer = OfferHelper.MakeCompensation(o.Parcel, "FinalEasement");
            FinalOptionOffer = OfferHelper.MakeCompensation(o.Parcel, "FinalOption");
            FinalROEOffer = OfferHelper.MakeCompensation(o.Parcel, "FinalROE");
        }

    }

    public class ParcelStatusHistoryDto
    {
        public IEnumerable<StatusDto> History { get; set; }
        public string Status { get; set; }
        public string ParcelUrl { get; set; }
    }

    public class StatusDto
    {
        public enum StageCode { Pending, InProgress, Completed, Aborted };
        public string Label { get; set; }
        public string Code { get; set; }
        public string Category { get; set; }
        public string ParentCode { get; set; }
        public int DisplayOrder { get; set; }
        public string Stage { get; set; }
        public bool IsSet { get; set; } = false;
        [JsonProperty(NullValueHandling =NullValueHandling.Include)]
        public Guid? StatusHistory { get; set; } = null;
        [JsonProperty(NullValueHandling = NullValueHandling.Include)]
        public Guid? StatusAgent { get; set; } = null;
        [JsonProperty(NullValueHandling = NullValueHandling.Include)]
        public DateTime? ActivityDate { get; set; } = null;

        public string Condition { get; set; }
        public DateTime? EffectiveStart { get; set; }
        public DateTime? EffectiveEnd { get; set; }

        public string Action { get; set; }
        public DateTime? DueDate { get; set; }
    }

    public class ActionItemDto
    {
        public Guid ActionItemId { get; set; }
        public string Action { get; set; }
        public string StatusCode { get; set; }
        public DateTimeOffset? DueDate { get; set; }
    }
    #endregion
    #region parcel graph
    public class ParcelGraph
    {
        public string ParcelId { get; set; }
        public string TractNo { get; set; }
        public string ParcelStatusCode { get; set; }
        public string ParcelStatusDate { get; set; }
        public string ParcelStatus => this.ParcelStatusCode;        // to be removed
        public string RoeStatusCode { get; set; }
        public string RoeStatusDate { get; set; }
        public string RoeCondition { get; set; } = "";
        public DateTime? RoeConditionStart { get; set; }
        public DateTime? RoeConditionEnd { get; set; }
        public string OutreachStatusCode { get; set; }
        public int? LandownerScore { get; set; }
        public string SitusAddress { get; set; }
        public double Acreage { get; set; }

        public bool IsImpacted { get; set; } = true;

        public Compensation InitialROEOffer { get; set; }
        public Compensation FinalROEOffer { get; set; }
        public Compensation InitialOptionOffer { get; set; }
        public Compensation FinalOptionOffer { get; set; }
        public Compensation InitialEasementOffer { get; set; }
        public Compensation FinalEasementOffer { get; set; }


        public IEnumerable<OwnerDto> Owners { get; set; }
        public IEnumerable<ContactLogDto> ContactsLog { get; set; }
        public IEnumerable<DocumentHeader> Documents { get; set; }
        public IEnumerable<ActionItemDto> ActionItems { get; set; }

        RoeCondition ActiveCondition(IEnumerable<RoeCondition> conditions)
        {
            if (conditions.Any() == false)
                return default;

            if (conditions.Count() == 1)
                return conditions.First();

            var today = DateTimeOffset.UtcNow;

            var q = from c in conditions
                    where (c.EffectiveStartDate < today || c.EffectiveStartDate == null) && (c.EffectiveEndDate == null || c.EffectiveEndDate > today)
                    select c;

            return q.Any() ? q.OrderBy(cx => cx.EffectiveStartDate).FirstOrDefault() : default;
        }

        internal ParcelGraph(Parcel p, DocTypes docTypes, IEnumerable<Document> d)
        {
            ParcelId = p.Assessor_Parcel_Number;
            TractNo = p.Tracking_Number;
            ParcelStatusCode = p.ParcelStatusCode;
            //ParcelStatus = Enum.GetName(typeof(Parcel.RowStatus), p.ParcelStatus);
            ParcelStatusDate = p.Activities.Where(ax => ax.StatusCode == ParcelStatusCode).OrderBy(ax => ax.ActivityDate).LastOrDefault()?.ActivityDate.LocalDateTime.ToShortDateString() ?? string.Empty;
            RoeStatusCode = p.RoeStatusCode;
            RoeStatusDate = p.Activities.Where(ax => ax.StatusCode == RoeStatusCode).OrderBy(ax => ax.ActivityDate).LastOrDefault()?.ActivityDate.LocalDateTime.ToShortDateString() ?? string.Empty;

            var cod = ActiveCondition(p.Conditions.ToArray());
            if (cod != null)
            {
                // RoeCondition = p.Conditions.FirstOrDefault()?.Condition ?? "";
                RoeCondition = cod.Condition;
                RoeConditionStart = cod.EffectiveStartDate?.LocalDateTime;
                RoeConditionEnd = cod.EffectiveEndDate?.LocalDateTime;
            }

            OutreachStatusCode = p.OutreachStatusCode;
            SitusAddress = p.SitusAddress;

            LandownerScore = p.Landowner_Score;

            Acreage = p.Acreage ?? 0;
            IsImpacted = p.IsImpacted;

            InitialEasementOffer = OfferHelper.MakeCompensation(p, "InitialEasement");
            InitialOptionOffer = OfferHelper.MakeCompensation(p, "InitialOption");
            InitialROEOffer = OfferHelper.MakeCompensation(p, "InitialROE");
            FinalEasementOffer = OfferHelper.MakeCompensation(p, "FinalEasement");
            FinalOptionOffer = OfferHelper.MakeCompensation(p, "FinalOption");
            FinalROEOffer = OfferHelper.MakeCompensation(p, "FinalROE");

            Owners = (p.Ownership.Any()) ? p.Ownership.Where(ox => ox.IsCurrentOwner).Select(ox => new OwnerDto(ox.Owner, ox.Ownership_t))
                // data issue. not common
                : new List<OwnerDto> { new OwnerDto(
                    new Owner { OwnerId = Guid.Empty, PartyName = "Owner Unknwon",
                        ContactInfo = new List<ContactInfo> { new ContactInfo { FirstName = "lost & found", Email = "lost&found@hdrinc.com" } }  }, 1) };

            ContactsLog = p.ContactLog.Where(cx => !cx.IsDeleted).Select(cx => new ContactLogDto(cx));
            Documents = d.Where(dx => !dx.IsDeleted).Select(dx =>
            {
                var dd = new DocumentHeader(dx);
                dd.DisplayCategory = docTypes.Find(dd.DocumentType)?.DisplayCategory ?? dd.DocumentType;
                return dd;
            });

            ActionItems = p.ActionItems.Select(ax => new ActionItemDto { ActionItemId = ax.ActionItemId, Action = ax.Action, StatusCode = Enum.GetName(typeof(ActionStatus), ax.Status), DueDate = ax.DueDate });
        }
    }
    #endregion

    #region offer helper
    static class OfferHelper
    {
        internal static Compensation MakeCompensation(Parcel p, string k)
        {

            switch (k)
            {
                case "InitialEasement":
                    return new Compensation
                    {
                        OfferAmount = p.InitialEasementOffer_OfferAmount ?? 0,
                        OfferDate = p.InitialEasementOffer_OfferDate ?? DateTimeOffset.MinValue,
                        OfferNotes = p.InitialEasementOffer_OfferNotes
                    };
                case "InitialOption":
                    return new Compensation
                    {
                        OfferAmount = p.InitialOptionOffer_OfferAmount ?? 0,
                        OfferDate = p.InitialOptionOffer_OfferDate ?? DateTimeOffset.MinValue,
                        OfferNotes = p.InitialOptionOffer_OfferNotes
                    };
                case "InitialROE":
                    return new Compensation
                    {
                        OfferAmount = p.InitialROEOffer_OfferAmount ?? 0,
                        OfferDate = p.InitialROEOffer_OfferDate ?? DateTimeOffset.MinValue,
                        OfferNotes = p.InitialROEOffer_OfferNotes
                    };
                case "FinalEasement":
                    return new Compensation
                    {
                        OfferAmount = p.FinalEasementOffer_OfferAmount ?? 0,
                        OfferDate = p.FinalEasementOffer_OfferDate ?? DateTimeOffset.MinValue,
                        OfferNotes = p.FinalEasementOffer_OfferNotes
                    };
                case "FinalOption":
                    return new Compensation
                    {
                        OfferAmount = p.FinalOptionOffer_OfferAmount ?? 0,
                        OfferDate = p.FinalOptionOffer_OfferDate ?? DateTimeOffset.MinValue,
                        OfferNotes = p.FinalOptionOffer_OfferNotes
                    };
                case "FinalROE":
                    return new Compensation
                    {
                        OfferAmount = p.FinalROEOffer_OfferAmount ?? 0,
                        OfferDate = p.FinalROEOffer_OfferDate ?? DateTimeOffset.MinValue,
                        OfferNotes = p.FinalROEOffer_OfferNotes
                    };
                default:
                    throw new IndexOutOfRangeException();
            }
        }
    }
    #endregion
}