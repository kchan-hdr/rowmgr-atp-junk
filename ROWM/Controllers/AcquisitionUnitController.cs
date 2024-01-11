using Austin_Costs;
using Microsoft.AspNetCore.Mvc;
using ROWM.Dal;
using System;
using System.Collections.Generic;
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
        public async Task<ActionResult<IEnumerable<PackageParcel>>> GetUnitParcelsByApn(string tractNo)
        {
            return Ok(await GetParcels(tractNo));

            //var parcels = await _ctx.Parcel.Where(px => px.Tracking_Number == tractNo).ToArrayAsync();
            //if (!(parcels?.Any() ?? false))
            //    return NotFound();

            //return Ok(parcels.Select(px => new PackageParcel(px)));
        }

        [HttpGet("parcels/{pid}/acq")]
        public async Task<ActionResult<IEnumerable<PackageParcel>>> GetUnitParcels(string pid)
        {
            //var p = await _ctx.Parcel.FindAsync(pid)
            //if (p == null)
            //    return NotFound();

            return Ok(await GetParcels(pid));
        }

        async Task<IEnumerable<PackageParcel>> GetParcels(string pid)
        { 
            /// TODO: need to update database schema. using COST_ESTIMATE_PARCEL
            /// 
            var same = _ctx.Database.SqlQuery<AcqParcel>("SELECT acq.* FROM [austin].[cost_estimate_parcel] p1 left join [austin].[cost_estimate_parcel] acq ON p1.[Acquisition_Parcel_No] = acq.[Acquisition_Parcel_No] WHERE p1.[TCAD_PROP_ID] = @pid",
                    new System.Data.SqlClient.SqlParameter("pid", pid));

            var sameacq = await same.ToListAsync();
            
            if (!sameacq.Any())
                return Enumerable.Empty<PackageParcel>();

            var acq = sameacq.First().Acquisition_Parcel_No;
            var candidate = sameacq.Select(a => a.TCAD_PROP_ID).ToArray();

            var parcels = await _ctx.Parcel.Where(px => candidate.Contains(px.Tracking_Number)).ToArrayAsync();

            return parcels.Select(px => new PackageParcel(px, acq));
        }
    }

    #region private records
    public class AcqParcel
    {
        public string Acquisition_Parcel_No { get; set; }
        public string TCAD_PROP_ID { get; set; }
    }
    #endregion

    public class PackageParcel
    {
        public Guid ParcelGuid { get; set; }
        public string Apn { get; set; }
        public string TCad { get => this.Apn; }
        public string TractNo { get; set; }
        public string AcquisitionUnitNumber { get; set; }
        public string SitusAddress { get; set; }

        public PackageParcel() { }
        public PackageParcel(Parcel p, string acq = "")
        {
            ParcelGuid = p.ParcelId;
            Apn = p.Assessor_Parcel_Number;
            TractNo = p.Tracking_Number;
            SitusAddress = p.SitusAddress;
            AcquisitionUnitNumber = acq;
        }
    }
}
