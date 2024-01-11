using com.hdr.rowmgr.Relocation;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ROWM.Dal;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace ROWM.Controllers
{
    [Produces("application/json")]
    [ApiController]
    public class RelocationController : ControllerBase
    {
        const string _APP_NAME = "RELOCATION";

        readonly RelocationRepository _repo;
        public RelocationController(RelocationRepository r) => (_repo) = (r);

        [HttpGet("api/RelocationActivityTypes"), Obsolete] //, ResponseCache(Duration=60*60)]
        public async Task<IEnumerable<IRelocationActivityType>> GetTypes() => await _repo.GetActivityTypes();

        [HttpGet("api/RelocationActivityPicklist"), Obsolete] // , ResponseCache(Duration = 60 * 60)]
        public async Task<IEnumerable<ActivityTaskPick>> GetPicklist()
        {
            var acts = await _repo.GetActivityTypes();

            return acts.OrderBy(a => a.DisplayOrder)
                .SelectMany(a => a.GetTasks())
                .ToArray();
        }

        /// <summary>
        /// Get Relocation Cases by APN or Tracking Number
        /// </summary>
        /// <param name="pid"></param>
        /// <returns></returns>
        [HttpGet("api/parcels/{pid}/relocations")]
        [ApiConventionMethod(typeof(DefaultApiConventions),nameof(DefaultApiConventions.Get))]
        [ProducesResponseType(StatusCodes.Status200OK, Type = typeof(RelocationDto))]
        //[ProducesResponseType(StatusCodes.Status400BadRequest)]
        //[ProducesResponseType(StatusCodes.Status404NotFound)]
        public async Task<IActionResult> GetRelocation([FromServices] OwnerRepository _oRepo, string pid)
        {
            var p = await _oRepo.GetParcel(pid);
            if (p == null)
                return NotFound($"bad parcel {pid}");

            var rr = await _repo.GetRelocationForParcel(p.ParcelId);
            foreach (var c in rr)
                c.ParcelKey = p.Tracking_Number;

            return new JsonResult(new RelocationDto(p.ParcelId, rr));
            //var r = await _repo.GetRelocation(p.ParcelId);
            //if (r == null)
            //    return NoContent();

            //// this needs to be configurable
            //foreach (var c in r.RelocationCases)
            //    c.ParcelKey = p.Tracking_Number;

            //return new JsonResult(new RelocationDto(r));
        }

        [HttpGet("api/parcels/{guid:Guid}/relocations"), Obsolete]
        [ProducesResponseType(StatusCodes.Status200OK, Type = typeof(RelocationDto))]
        public async Task<IActionResult> GetRelocation(Guid guid)
        {
            if (Guid.Empty == guid)
                return BadRequest();

            var r = await _repo.GetRelocation(guid);

            return new JsonResult(new RelocationDto(r));
        }

        [HttpPost("api/parcels/{pId}/relocations")]
        [ProducesDefaultResponseType(typeof(RelocationDto))]
        [ApiConventionMethod(typeof(DefaultApiConventions), nameof(DefaultApiConventions.Post))]
        [ProducesResponseType(StatusCodes.Status201Created, Type = typeof(RelocationDto))]
        public async Task<IActionResult> AddRelocationCase([FromServices] OwnerRepository _oRepo, [FromServices] IRelocationCaseOps _caseOps, 
            string pId, 
            [FromBody] RequestCase rCase)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            var p = await _oRepo.GetParcel(pId);
            if (p == null)
                return BadRequest($"unknown parcel <{pId}>");

            var r = await _caseOps.AddRelocationCase(p.ParcelId, rCase.DisplaceeName, rCase.Eligibility, rCase.DisplaceeType, rCase.Hs, rCase.Rap);
            return CreatedAtAction(nameof(GetRelocationCase), new { rcId = r.RelocationCaseId });
        }

        //[HttpPost("api/parcels/{guid:Guid}/relocations")]
        //[ProducesResponseType(StatusCodes.Status202Accepted, Type = typeof(RelocationDto))]
        //public async Task<IActionResult> AddRelocationCase(Guid guid, [FromBody] RequestCase rCase)
        //{
        //    if (!ModelState.IsValid)
        //        return BadRequest(ModelState);

        //    var r = await _caseOps.AddRelocationCase(guid, rCase.DisplaceeName, rCase.Eligibility, rCase.DisplaceeType, rCase.Hs, rCase.Rap);
        //    return new JsonResult(new RelocationDto(r));
        //}

        [HttpGet("api/relocations/{rcId}")]
        [ProducesDefaultResponseType(typeof(RelocationCaseDto))]
        [ProducesResponseType(typeof(RelocationCaseDto), StatusCodes.Status200OK)]
        public async Task<IActionResult> GetRelocationCase(Guid rcId)
        {
            var c = await _repo.GetRelocationCase(rcId);
            return new JsonResult(new RelocationCaseDto(c));
        }

        [HttpPost("api/relocations/{rcId}/activities")]
        [ProducesResponseType(StatusCodes.Status201Created, Type = typeof(RelocationDto))]
        [ApiConventionMethod(typeof(DefaultApiConventions), nameof(DefaultApiConventions.Post))]
        public async Task<IActionResult> AddRelocationActivity([FromServices] IRelocationCaseOps _caseOps, Guid rcId, [FromBody] RequestActivity act)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            var r = await _caseOps.AddActivity(rcId
                , act.ActivityCode
                , act.Activity
                , act.Description
                , act.AgentId
                , act.ActivityDate
                , act.Notes
                , act.Money
                , act.YesNo);

            return new JsonResult(new RelocationDto(r));
        }

        [HttpPut("api/relocations/{rcId}"), Obsolete]
        [ProducesResponseType(StatusCodes.Status201Created, Type = typeof(RelocationDto))]
        [ApiConventionMethod(typeof(DefaultApiConventions), nameof(DefaultApiConventions.Put))]
        public async Task<IActionResult> UpdateRelocationCaseData([FromServices] IRelocationCaseOps _caseOps, Guid rcId, [FromBody] RequestActivityData act)
        {
            var r = await ProcessRequest(_caseOps, act);
            return new JsonResult(new RelocationCaseDto(r));
        }

        [HttpGet("api/relocations/{rcId}/activities")]
        [ProducesDefaultResponseType(typeof(DisplaceeHistory))]
        [ApiConventionMethod(typeof(DefaultApiConventions), nameof(DefaultApiConventions.Get))]
        public async Task<DisplaceeHistory> GetRelocationCaseData([FromServices] IRelocationCaseOps _caseOps, Guid rcId)
        {
            var master = await _repo.GetActivityTypes();
            var data = await _caseOps.GetActivities(rcId);
            return new DisplaceeHistory(master, data);
        }

        [HttpPost("api/relocations/{rcId:Guid}/logs/{logId:Guid}")]
        [ApiConventionMethod(typeof(DefaultApiConventions), nameof(DefaultApiConventions.Post))]
        public async Task<IActionResult> AddContactLog(Guid rcId, Guid logId)
        {
            var touched = await _repo.AttachLog(rcId, logId);
            return CreatedAtAction(nameof(GetRelocationCase), new { rcId = rcId });
        }

        [HttpPost("api/relocations/{rcId:Guid}/docs/{docId:Guid}")]
        public async Task<ActionResult> AddDocuments(Guid rcId, Guid docId)
        {
            var touched = await _repo.AttachDoc(rcId, docId);
            return CreatedAtAction(nameof(GetRelocationCase), new { rcId = rcId });
        }
        #region helper
        async Task<IRelocationCase> ProcessRequest(IRelocationCaseOps _caseOps, RequestActivityData data)
        {
            var caseId = data.ParentCaseId;
            var a = await _caseOps.GetActivities(caseId);
            var original = new RequestActivityData(a);

            if (data.Day90NoticeSent != null && original.Day90NoticeSent != data.Day90NoticeSent)
            {
                await _caseOps.AddActivity(caseId, "90day", DisplaceeActivity.sent, "90-Day Notice", Guid.Empty, data.Day90NoticeSent.Value, string.Empty);
            }
            if (data.Day90NoticeDelivered != null && original.Day90NoticeDelivered != data.Day90NoticeDelivered)
            {
                await _caseOps.AddActivity(caseId, "90day", DisplaceeActivity.delivered, "90-Day Notice", Guid.Empty, data.Day90NoticeDelivered.Value, string.Empty);
            }
            if (data.Day30NoticeSent != null && original.Day30NoticeSent != data.Day30NoticeSent)
            {
                await _caseOps.AddActivity(caseId, "30day", DisplaceeActivity.sent, "30-Day Notice", Guid.Empty, data.Day30NoticeSent.Value, string.Empty);
            }
            if (data.Day30NoticeDelivered != null && original.Day30NoticeDelivered != data.Day30NoticeDelivered)
            {
                await _caseOps.AddActivity(caseId, "30day", DisplaceeActivity.delivered, "30-Day Notice", Guid.Empty, data.Day30NoticeDelivered.Value, string.Empty);
            }

            var r = await _repo.GetRelocationCase(caseId);
            return r;
        }
        #endregion
    }
    #region dto
    public class RelocationDto
    {
        public Guid ParcelId { get; set; }
        public IEnumerable<RelocationCaseDto> Cases { get; set; }

        public RelocationDto() { }
        public RelocationDto(Guid parcelid, IEnumerable<IRelocationCase> cases)
        {
            this.ParcelId = parcelid;
            Cases = cases.Select(c => new RelocationCaseDto(c));
        }

        public RelocationDto(IParcelRelocation p)
        {
            this.ParcelId = p.ParcelId;
            Cases = p.RelocationCases.Select(c => new RelocationCaseDto(c));
        }
    }
    public class RelocationCaseDto
    {
        public Guid RelocationCaseId { get; set; }

        public Guid? AgentId { get; set; }
        public int RelocationNumber { get; set; }
        public string Status { get; set; }
        public string DisplaceeType { get; set; }
        public string RelocationType { get; set; }

        public string DisplaceeName { get; set; }
        Guid? ContactInfoId { get; set; }

        public int CompletedSteps { get; set; }

        public IEnumerable<Guid> ContactsLog { get; set; }
        public IEnumerable<Guid> Documents { get; set; }
        //// details
        //IEnumerable<IRelocationEligibilityActivity> EligibilityHistory { get; }
        //IEnumerable<IRelocationDisplaceeActivity> DisplaceeActivities { get; }


        public string AcqFilenamePrefix { get; set; }
        public string CaseUrl { get; set; }

        public RelocationCaseDto(IRelocationCase c)
        {
            RelocationCaseId = c.RelocationCaseId;
            RelocationNumber = c.RelocationNumber;

            AcqFilenamePrefix = c.AcqFilenamePrefix;

            DisplaceeName = c.DisplaceeName;
            Status = Enum.GetName(typeof(RelocationStatus), c.Status);
            DisplaceeType = Enum.GetName(typeof(DisplaceeType), c.DisplaceeType);
            RelocationType = Enum.GetName(typeof(RelocationType), c.RelocationType);

            CompletedSteps = c.CompletedSteps;

            ContactsLog = c.ContactLogIds;
            Documents = c.DocumentIds;

            //var austin = c as RelocationCase;
            //if (austin != null)
            //{
            //    //ContactsLog = austin.Logs.Where(cx => !cx.IsDeleted).Select(cx => new ContactLogDto(cx));
            //    //Documents = austin.Documents.Where(dx => !dx.IsDeleted).Select(dx => new DocumentHeader(dx));
            //}
        }
    }

    #region requests
    public class RequestCase
    {
        public string DisplaceeName { get; set; }
        public string Eligibility { get; set; }
        public string[] DisplaceeType { get; set; }
        //public RelocationType RelocationType { get; set; }
        public double? Hs { get; set; }
        public double? Rap { get; set; }
    }

    public class RequestActivity
    {
        public string ActivityCode { get; set; }
        public DisplaceeActivity Activity { get; set; }
        public string? Description { get; set; }

        [DefaultValue("161DED15-9122-47CF-9A11-4139C3FF5C05")]
        public Guid AgentId { get; set; }
        public DateTimeOffset ActivityDate { get; set; }
        public string? Notes { get; set; }
        public int? Money { get; set; }
        public bool? YesNo { get; set; }
    }
    #endregion

    #region new relocation case status
    public class DisplaceeHistory
    {
        static readonly List<StatusDto> _MASTER = new List<StatusDto> 
        {
            new StatusDto { Label="Initial Interview", Code = "interview", DisplayOrder = 1, IsSet=true, Stage = "" },
            new StatusDto { Label="Client Approval", Code = "txdot", DisplayOrder = 2, IsSet=true, Stage = "" },
            new StatusDto { Label="90-day Notice", Code = "90day", DisplayOrder = 3, IsSet=true, Stage = "" },
            new StatusDto { Label="30-day Notice", Code = "30day", DisplayOrder = 4, IsSet=true, Stage = "" },
            new StatusDto { Label="Actual Date Vacated", Code = "vacated", DisplayOrder = 5, IsSet=true, Stage = "" },
            new StatusDto { Label="Cert of Completion", Code = "certificate", DisplayOrder = 6, IsSet=true, Stage = "" }
        };

        public IEnumerable<StatusDto> History { get; set; }

        public DisplaceeHistory() { }
        public DisplaceeHistory(IEnumerable<IRelocationActivityType> master, IEnumerable<IRelocationDisplaceeActivity> activities)
        {
            var mm = master.OrderBy(mx => mx.DisplayOrder).Select(mx => new StatusDto
            {
                Label = mx.Description,
                Code = mx.ActivityTypeCode,
                DisplayOrder = mx.DisplayOrder,
                IsSet = true,
                Stage = ""
            });

            History = mm // _MASTER
                .OrderBy(sx => sx.DisplayOrder)                
                .Select(sx =>
                {
                    if (activities.Any(ax => ax.ActivityCode == sx.Code))
                    {
                        sx.ActivityDate = activities.Where(ax => ax.ActivityCode == sx.Code).Max(ax => ax.ActivityDate).UtcDateTime;
                        sx.Stage = "Completed";
                    } else
                    {
                        sx.ActivityDate = null;
                        sx.Stage = string.Empty;
                    }
                    return sx;
                });
        }
    }
    #endregion
    #region data upload (obsolete)
    [Obsolete]
    public class RequestActivityData
    {
        public Guid ParentCaseId { get; set; }

        public DateTimeOffset? Day90NoticeSent { get; set; }
        public DateTimeOffset? Day90NoticeDelivered { get; set; }
        public DateTimeOffset? Day30NoticeSent { get; set; }
        public DateTimeOffset? Day30NoticeDelivered { get; set; }
        public DateTimeOffset? DateRequiredToMove { get; set; }
        public DateTimeOffset? DateBenefitsExpire { get; set; }
        public DateTimeOffset? ActualDateVacated { get; set; }
        public bool? VacateExtensionApprovedByClient { get; set; }
        public bool? DemoRequired { get; set; }
        public DateTimeOffset? DisplaceeCertificateOfCompletion { get; set; }
        public DateTimeOffset? SupplementSubmittedToConsultant { get; set; }
        public DateTimeOffset? SupplementSubmittedToClient { get; set; }
        public DateTimeOffset? SupplementApprovedByClient { get; set; }
        public DateTimeOffset? NegotiatedSubmittedToConsultant { get; set; }
        public DateTimeOffset? NegotiatedSubmittedToClient { get; set; }
        public DateTimeOffset? NegotiatedApprovedByClient { get; set; }
        public DateTimeOffset? FixedSubmittedToConsultant { get; set; }
        public DateTimeOffset? FixedSubmittedToClient { get; set; }
        public DateTimeOffset? FixedApprovedByClient { get; set; }
        public DateTimeOffset? TemporarySubmittedToConsultant { get; set; }
        public DateTimeOffset? TemporarySubmittedToClient { get; set; }
        public DateTimeOffset? TemporaryApprovedByClient { get; set; }

        public RequestActivityData() { }
        public RequestActivityData(IEnumerable<IRelocationDisplaceeActivity> act)
        {
            //foreach(var a in act.OrderBy(ax => ax.ActivityDate))
            //{
            //    switch (a.ActivityCode)
            //    {
            //        case "30day":
            //            if (a.Activity == DisplaceeActivity.sent)
            //                Day30NoticeSent = a.ActivityDate;
            //            if (a.Activity == DisplaceeActivity.delivered)
            //                Day30NoticeDelivered = a.ActivityDate;
            //            break;
            //        case "90day":
            //            if (a.Activity == DisplaceeActivity.sent)
            //                Day90NoticeSent = a.ActivityDate;
            //            if (a.Activity == DisplaceeActivity.delivered)
            //                Day90NoticeDelivered = a.ActivityDate;
            //            break;
            //    }
            //}

            DemoRequired = true;
        }
    }
    #endregion
    #endregion
}
