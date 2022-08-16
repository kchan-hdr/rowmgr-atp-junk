using Microsoft.AspNetCore.Mvc;
using ROWM.Dal;
using System;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;

namespace ROWM.Controllers
{
    [Route("api/v2")]
    [ApiController]
    public class AcquisitionUnitController : ControllerBase
    {
        readonly OwnerRepository _repo;
        readonly ROWM_Context _ctx;
        public AcquisitionUnitController(OwnerRepository o, ROWM_Context c) => (_repo, _ctx) = (o, c);

        [HttpGet("acqUnits/{tractNo}")]
        public async Task<ActionResult> GetUnitParcelsByApn(string tractNo)
        {
            var parcels = await _ctx.Parcel.Where(px => px.Tracking_Number == tractNo).ToArrayAsync();
            if (!(parcels?.Any() ?? false))
                return NotFound();

            return new JsonResult(parcels.Select(px => new PackageParcel(px)));
        }

        [HttpGet("parcels/{pid:Guid}/acq")]
        public async Task<IActionResult> GetUnitParcels(Guid pid)
        {
            var p = await _ctx.Parcel.FindAsync(pid);
            if (p == null)
                return NotFound();

            var parcels = await _ctx.Parcel.Where(px => px.Tracking_Number == p.Tracking_Number).ToArrayAsync();

            return new JsonResult(parcels.Select(px => new PackageParcel(px)));
        }
    }

    public class PackageParcel
    {
        public Guid ParcelGuid { get; set; }
        public string Apn { get; set; }
        public string TCad { get => this.Apn; }
        public string TractNo { get; set; }
        public string AcquisitionUnitNumber { get => this.TractNo; }  // alias this guy to help
        public string SitusAddress { get; set; }

        public PackageParcel() { }
        public PackageParcel(Parcel p)
        {
            ParcelGuid = p.ParcelId;
            Apn = p.Assessor_Parcel_Number;
            TractNo = p.Tracking_Number;
            SitusAddress = p.SitusAddress;
        }
    }
}
