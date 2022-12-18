using Microsoft.AspNetCore.Mvc;
using ROWM.Dal;
using System;
using System.Collections.Generic;
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
        public OwnerController(OwnerRepository r) => (_repo) = (r);

        [Route("owners/{id:Guid}"), HttpGet]
        public async Task<OwnerDto2> GetOwner(Guid id) => new OwnerDto2(await _repo.GetOwner(id));

        [HttpGet("owners", Name ="All Owners")]
        public async Task<IEnumerable<OwnerDto2>> AllOwners([FromServices] IOwnershipHelper h) =>
            (await h.AllOwners()).Select(ox => new OwnerDto2(ox));

        [HttpGet("owners/{name}", Name = "Find Owners")]
        public async Task<IEnumerable<OwnerDto2>> FindOwner(string name) =>
            (await _repo.FindOwner(name)).Select(ox => new OwnerDto2(ox));


        [HttpPost("owners")]
        public async Task<ActionResult<OwnerDto2>> AddOwner([FromServices] IOwnershipHelper h, [FromBody] OwnershipRequest o)
        {
            Owner myLord = null;

            var ow = await _repo.FindOwner(o.Name);
            if (ow.Any())
            {
                var potential = ow.Where(ox => ox.PartyName.Equals(o.Name, StringComparison.InvariantCultureIgnoreCase));
                Trace.TraceWarning($"submitting a duplicated user {o.Name}");

                // user didn't search
                myLord = potential.FirstOrDefault();
            }

            //myLord ??= await h.AddOwner(o.Name, o.Address, o.OwnerType);
            if (myLord == null)
                myLord = await h.AddOwner(o.Name, o.Address, o.OwnerType);

            return Ok(new OwnerDto2(await ChangeOwnerImpl(h, myLord, o)));
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
        public async Task<ActionResult<OwnerDto2>> ChangeOwner([FromServices] IOwnershipHelper h, Guid id, [FromBody] OwnershipRequest o)
        {
            var owner = await h.GetOwner(id);
            if (owner == null)
                return BadRequest();

            return Ok(new OwnerDto2(await ChangeOwnerImpl(h, owner, o)));
        }

        async Task<Owner> ChangeOwnerImpl(IOwnershipHelper h, Owner owner, OwnershipRequest o)
        {
            var parcels = new List<Guid>();
            parcels.AddRange(await h.GetParcelsByApn(o.Parcels));
            parcels.AddRange(await h.GetParcelsByAcquisitionUnit(o.AcquisitionUnits));

             return await h.NewOwnership(parcels, owner.OwnerId);
        }

        [HttpGet("parcels/{pid}/owners")]
        public async Task<ActionResult<IEnumerable<Ownership>>> GetOwnersForParcel([FromServices] IOwnershipHelper h, string pid)
        {
            var p = await _repo.GetParcel(pid);
            if (p == null)
                return BadRequest();
             
            var ox = await h.GetOwners(p.ParcelId);

            return Ok(value: ox);
        }
    }

    public class OwnershipRequest 
    {
        public string Name { get; set; }
        public string Address { get; set; }
        public string OwnerType { get; set; }


        public IEnumerable<string> AcquisitionUnits { get; set; }   // ATP uses funny acq units
        public IEnumerable<string> Parcels { get; set; }    // optional parcels
    }

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
}
