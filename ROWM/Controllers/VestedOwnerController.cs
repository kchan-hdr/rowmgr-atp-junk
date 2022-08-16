using Microsoft.AspNetCore.Mvc;
using ROWM.Dal;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Threading.Tasks;

namespace ROWM.Controllers
{
    [ApiController]
    [Route("api/v2")]
    public class VestedOwnerController : ControllerBase
    {
        #region ctor
        readonly ROWM_Context _ctx;
        readonly OwnerRepository _repo;
        public VestedOwnerController(ROWM_Context c, OwnerRepository o) => (_ctx, _repo) = (c, o);
        #endregion
        [HttpGet("acqUnits/{tractNo}/vested"), ActionName(nameof(GetVestedOwner))]
        [ProducesDefaultResponseTypeAttribute(typeof(IEnumerable<Vested_dto>))]
        public async Task<IActionResult> GetVestedOwner(string tractNo)
        {
            var owners = await _ctx.VestedOwner.Where(vx => vx.TrackingNumber.Equals(tractNo)).ToArrayAsync();
            if (!(owners?.Any() ?? false))
                return new JsonResult(Enumerable.Empty<Vested_dto>()); // NotFound($"no vested owner for parcel <{tractNo}>");

            return new JsonResult(owners.Select(o => new Vested_dto(o, Url)));
        }

        [HttpPost("acqUnits/{tractNo}/vested")]
        public async Task<IActionResult> AddVestedOwner(string tractNo, [FromBody] Vested_req req)
        {
            var parcels = await _ctx.Parcel.Where(px => px.Tracking_Number == tractNo).ToArrayAsync();
            if (!(parcels?.Any() ?? false))
                return NotFound();

            try
            {
                var owners = new List<Vested_req>();
                foreach (var px in parcels)
                {
                    var ox = await UpsertOwner(px, req);
                    owners.Add(new Vested_dto(ox, Url));
                }
                var o = owners.First();
                return new AcceptedResult(Url.Action(nameof(GetVestedOwner), tractNo), o);
            }
            catch (DbUpdateException ex)
            {
                throw;
                //return Problem(_);
            }
        }

        [HttpGet("parcels/{apn}/individualvested")]
        public async Task<IActionResult> GetVestedOwnerByParcel(string apn)
        {
            var owners = await _ctx.VestedOwner.FirstOrDefaultAsync(vx => vx.Tcad.Equals(apn));
            if (owners == null)
                return NotFound($"no vested owner for parcel <{apn}>");

            return new JsonResult(new Vested_dto(owners, Url));
        }

        [HttpPost("parcels/{apn}/individualvested")]
        public async Task<IActionResult> AddVestedOwnerByParcel(string apn, [FromBody] Vested_req req)
        {
            var p = await _ctx.Parcel.SingleOrDefaultAsync(px => px.Assessor_Parcel_Number == apn);
            if (p == null)
                return NotFound();

            try
            {
                var o = await UpsertOwner(p, req);
                return new AcceptedResult(Url.Action(nameof(GetVestedOwner), p.Tracking_Number), new Vested_dto(o, Url));
            }
            catch (DbUpdateException ex)
            {
                throw;
                //return Problem(_);
            }
        }

        [HttpPut("vestedowners/{id:Guid}"), ActionName(nameof(UpdateVestedOwner))]
        public async Task<IActionResult> UpdateVestedOwner(Guid id, [FromBody] Vested_req req)
        {
            var p = await _ctx.Parcel.FindAsync(req.ParcelId);
            if (p == null)
                return BadRequest();

            if (p.ParcelId != req.ParcelId)
                return BadRequest($"invalid vested owner object");

            try
            {
                var o = await UpsertOwner(p, req);
                return new AcceptedResult(Url.Action(nameof(GetVestedOwner), p.Tracking_Number), new Vested_dto(o, Url));
            }
            catch (DbUpdateException ex)
            {
                throw;
            }
        }

        async Task<VestedOwner> UpsertOwner(Parcel p, Vested_req req)
        {
            var myAgent = (await _repo.GetAgent(req.AgentName)) ?? await _repo.GetDefaultAgent();
            var v = (await _ctx.VestedOwner.FindAsync(req.VestedOwnerId)) ?? _ctx.VestedOwner.Add(new VestedOwner());

            v.ParcelId = p.ParcelId;
            v.AcqNo = p.Assessor_Parcel_Number;
            v.TrackingNumber = p.Tracking_Number;
            v.Tcad = p.Assessor_Parcel_Number;

            v.VestedOwnerName = req.OwnerName;
            v.VestedOwnerAddress = req.OwnerAddress;
            v.IsVerified = req.IsVerified;
            v.LastModified = req.LastModified;
            v.ModifiedBy = req.ModifiedBy;
            v.AgentId = myAgent.AgentId;

            await _ctx.SaveChangesAsync();
            return v;
        }
    }

    #region request dto
    public class Vested_req
    {
        public Guid? VestedOwnerId { get; set; }
        public Guid? ParcelId { get; set; }
        public string OwnerName { get; set; }
        public string OwnerAddress { get; set; }
        public bool IsVerified { get; set; }
        public DateTimeOffset LastModified { get; set; }
        public string ModifiedBy { get; set; }
        public string AgentName { get; set; }
        public Guid? TitleDocument { get; set; }

        // endpoints
        public string GetUrl { get; set; }
        public string EditUrl { get; set; }

        public Vested_req() { }
        public Vested_req(VestedOwner o, IUrlHelper urlHelper = null)
        {
            VestedOwnerId = o.VestedOwnerId;
            ParcelId = o.ParcelId;
            OwnerName = o.VestedOwnerName;
            OwnerAddress = o.VestedOwnerAddress;
            IsVerified = o.IsVerified;
            AgentName = o.Agent.AgentName;
            LastModified = o.LastModified;
            ModifiedBy = o.ModifiedBy;

            if (urlHelper != null)
            {
                EditUrl = $"/api/v2/vestedowners/{VestedOwnerId}";
            }
        }
    }

    public class Vested_dto : Vested_req
    {
        public string TCad { get; set; }
        public string AcquisitionNumber { get; set; }

        public Vested_dto() { }
        public Vested_dto(VestedOwner o, IUrlHelper urlHelper =null) : base(o,urlHelper)
        {
            this.TCad = o.Tcad;
            this.AcquisitionNumber = o.AcqNo;
        }
    }
    #endregion
}
