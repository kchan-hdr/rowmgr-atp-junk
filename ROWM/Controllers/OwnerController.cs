using Austin_Costs;
using Microsoft.AspNetCore.Mvc;
using ROWM.Dal;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace ROWM.Controllers
{
    /// <summary>
    /// this version of "owners" combine the old, but unused, endpoints & Vested-Owner
    /// most importantly it implements change of ownership
    /// </summary>
    [Route("api/v2")]
    [ApiController]
    public class OwnerController : ControllerBase
    {
        readonly OwnerRepository _repo;
        readonly OwnershipRepository _Orepo;
        readonly CostEstimateContext _ctx;
        readonly IOwnershipHelper _helper;
        public OwnerController(OwnerRepository r, OwnershipRepository o, CostEstimateContext c, IOwnershipHelper h) => (_repo, _Orepo, _ctx, _helper) = (r, o, c, h);

        [Route("owners/{id:Guid}"), HttpGet]
        public async Task<OwnerDto2> GetOwner(Guid id) => new OwnerDto2(await _repo.GetOwner(id));

        [HttpGet("owners", Name ="All Owners")]
        public async Task<IEnumerable<OwnerDto2>> AllOwners() =>
            (await _helper.AllOwners()).Select(ox => new OwnerDto2(ox));

        [HttpGet("owners/{name}", Name = "Find Owners")]
        public async Task<IEnumerable<OwnerDto2>> FindOwner(string name) =>
            (await _repo.FindOwner(name)).Select(ox => new OwnerDto2(ox));


        [HttpPost("owners")]
        public async Task<ActionResult<OwnerDto2>> AddOwner([FromBody] OwnershipRequest o)
        {
            var myLord = await GetOrCreateOwner(o);
            return Ok(new OwnerDto2(await ChangeOwnerImpl(myLord, o)));
        }

        async Task<Owner> GetOrCreateOwner(OwnershipRequest o)
        {
            Owner myLord = null;
            if (o.OwnerId.HasValue)
                myLord = await _helper.GetOwner(o.OwnerId.Value);

            if (myLord == null)
            {
                var ow = await _repo.FindOwner(o.Name);
                if (ow.Any())
                {
                    var potential = ow.Where(ox => ox.PartyName.Equals(o.Name, StringComparison.InvariantCultureIgnoreCase));
                    Trace.TraceWarning($"submitting a duplicated user {o.Name}");

                    // user didn't search
                    myLord = potential.FirstOrDefault();
                }
            }

            //myLord ??= await h.AddOwner(o.Name, o.Address, o.OwnerType);
            if (myLord == null)
                myLord = await _helper.AddOwner(o.Name, o.Address, o.OwnerType);


            return myLord;
        }

        [HttpPut("owners/{id:Guid}")]
        public async Task<ActionResult<OwnerDto2>> UpdateOwner(Guid id, [FromBody] OwnerRequest o)
        {
            var ow = await _repo.GetOwner(id);
            if (ow == null)
                return BadRequest();

            ow.PartyName = o.PartyName;
            ow.OwnerType = o.OwnerType;

            ow = await _repo.UpdateOwner(ow);

            return new OwnerDto2(ow);
        }

        [HttpPost("owners/{id:Guid}/parcels")]
        public async Task<ActionResult<OwnerDto2>> ChangeOwner(Guid id, [FromBody] OwnershipRequest o)
        {
            var owner = await _helper.GetOwner(id);
            if (owner == null)
                return BadRequest();

            return Ok(new OwnerDto2(await ChangeOwnerImpl(owner, o)));
        }

        async Task<Owner> ChangeOwnerImpl(Owner owner, OwnershipRequest o)
        {
            var parcels = new List<Guid>();
            parcels.AddRange(await _helper.GetParcelsByApn(o.Parcels));
            //parcels.AddRange(await _helper.GetParcelsByAcquisitionUnit(o.AcquisitionUnits));

            _ = await _Orepo.SplitAcquisitionUnit(o.Parcels);

            return await _helper.NewOwnership(parcels, owner.OwnerId);
        }

        [HttpGet("parcels/{pid}/owners")]
        public async Task<ActionResult<IEnumerable<Ownership>>> GetOwnersForParcel(string pid)
        {
            var p = await _repo.GetParcel(pid);
            if (p == null)
                return BadRequest();
             
            var ox = await _helper.GetOwners(p.ParcelId);

            return Ok(value: ox);
        }

        /// <summary>
        /// mostly to maintain the original endpoint. don't expect much useage
        /// </summary>
        /// <param name="_docTypes"></param>
        /// <param name="pid"></param>
        /// <param name="o"></param>
        /// <returns></returns>
        [HttpPost("parcels/{pid}/owners")]
        public async Task<ActionResult<ParcelGraph>> SetOwner([FromServices] DocTypes _docTypes, [FromServices] geographia.ags.IFeatureUpdate_Austin fu, string pid, [FromBody] OwnershipRequest o)
        {
            var p = await _repo.GetParcel(pid);
            if (p == null)
                return BadRequest();

            var myLord = await GetOrCreateOwner(o);
            if (myLord.OwnerId == o.OwnerId)
            {
                await _Orepo.ChangeOwnerDetails(o);
            } else
            {
                await ChangeOwnerImpl(myLord, o);
            }
            await fu.UpdateFeature_Ex(pid, new Dictionary<string, dynamic>() { 
                { "OwnerName", myLord.PartyName },
                { "OwnerAddress", myLord.OwnerAddress }
            });

            return new ParcelGraph(p, _docTypes, await _repo.GetDocumentsForParcel(pid));
        }

        #region vested owner
        [HttpGet("parcels/{pid}/vestedOwners")]
        public async Task<ActionResult<VestedOwnerDto>> GetVestedOwnerForParcel(string pid, [FromServices]OwnershipRepository _Orepo)
        {
            var myVestedOwner = await _Orepo.GetVestedOwnerForParcel(pid);
            //if (myVestedOwner == null) return NotFound();
            return Ok(value: myVestedOwner);
        }

        [HttpPost("parcels/{pid}/vestedOwners")]
        public async Task<ActionResult<VestedOwnerDto>> SetVestedOwner(string pid, [FromBody] OwnershipRequest o)
        {
            var p = await _repo.GetParcel(pid);
            if (p == null)
                return BadRequest();

            var myVestOwner = await _Orepo.AddVestedOwner(o.Name, o.Address, p.ParcelId, o.OwnerType);

            return myVestOwner;
        }

        [HttpGet("owners/{name}/vestedOwners", Name = "Find Vested Owners")]
        public async Task<IEnumerable<VestedOwnerDto>> FindVestedOwner(string name, [FromServices] OwnershipRepository _Orepo) =>
            (await _Orepo.FindVestedOwner(name));

        #endregion

        ///
        [HttpGet("parcels/{id}/TCad_Owner")]
        public async Task<ActionResult<OwnerDto2>> GetOwner([FromServices] ROWM_Context context, [FromServices] OwnerRepository o, string id)
        {
            var owners = await context.Database.SqlQuery<string>("SELECT [PartyName] FROM Austin.TCAD_OWNER WHERE [Tracking_Number] = @pid",
                    new System.Data.SqlClient.SqlParameter("pid", id)
                ).FirstOrDefaultAsync();
            if (owners == null)
                return NotFound();

            var d = await o.FindOwner(owners);
            if (d == null || !d.Any())
                return NotFound();

            return Ok(new OwnerDto2(d.First()));
        }

        #region helper
        // split acquisition
        [Obsolete("use Split from OwnershipRepositry")]
        async Task<bool> SplitAcquisitionUnit(IEnumerable<string> parcels)
        {
            var pid = string.Join(",", parcels.Select(px => $"'{px}'"));

            var query = $"SELECT DISTINCT acq.* FROM [austin].[cost_estimate_parcel] p1 left join [austin].[cost_estimate_parcel] acq ON p1.[Acquisition_Parcel_No] = acq.[Acquisition_Parcel_No] WHERE p1.[TCAD_PROP_ID] in ( {pid})";

            // are all parcels belong to the same unit
            var acqParcel = _ctx.Database.SqlQuery<AcqParcel>(query);
            var acqs = await acqParcel.ToArrayAsync();

            if (!acqs.Any())
                return false;   // not found
            if (acqs.Select(ax => ax.Acquisition_Parcel_No).Distinct().Count() > 1)
                return false;   // this is unusual. not supported (changing owners from different acq units

            var unitNumber = acqs.First().Acquisition_Parcel_No;
            if (acqs.Count() == parcels.Count())
                return true;

            // do split then
            var kill = await _ctx.AcquisitionKeys.Where(kx => kx.AcqNo == unitNumber).ToListAsync();
            foreach(var p in kill)
            {
                var nkey = new AcquisitionKey
                {
                    AcqNo = p.AcqNo + (parcels.Contains(p.TrackingNumber) ? "_b" : "_a"),
                    PropId = p.PropId
                };
                _ctx.AcquisitionKeys.Add(nkey);
                _ctx.AcquisitionKeys.Remove(p);
            }

            return await _ctx.SaveChangesAsync() > 0;
        }

        #endregion
    }

    public class OwnershipRequest_ 
    {
        public Guid? OwnerId { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string OwnerType { get; set; }


        public IEnumerable<string> AcquisitionUnits { get; set; }   // ATP uses funny acq units
        public IEnumerable<string> Parcels { get; set; }    // optional parcels
    }

    #region more slim-down
    public class OwnerDto2
    {
        public Guid OwnerId { get; set; }
        public string PartyName { get; set; }
        public string OwnerAddress { get; set; }
        public string OwnerType { get; set; }
        public int OwnershipType { get; set; }
        public IEnumerable<ParcelMinHeaderDto> OwnedParcel { get; set; }

        public OwnerDto2(Owner o, int oType = 1)
        {
            OwnerId = o.OwnerId;
            PartyName = o.PartyName;
            OwnerAddress = o.OwnerAddress;
            OwnerType = o.OwnerType;
            OwnershipType = oType;

            OwnedParcel = o.Ownership.Select(ox => new ParcelMinHeaderDto(ox));
        }
    }

    public class ParcelMinHeaderDto
    {
        public string ParcelId { get; set; }
        public string TractNo { get; set; }
        public string SitusAddress { get; set; }
        public bool IsPrimaryOwner { get; set; }
        public bool IsRelinquished { get; set; }

        internal ParcelMinHeaderDto(Ownership o)
        {
            ParcelId = o.Parcel.Assessor_Parcel_Number;
            TractNo = o.Parcel.Tracking_Number;
            SitusAddress = o.Parcel.SitusAddress;
            IsPrimaryOwner = o.IsPrimary(); // .Ownership_t == Ownership.OwnershipType.Primary;
            IsRelinquished = o.Ownership_t == (int)Ownership.OwnershipType.Relinquished;
        }
    }
    #endregion
}
