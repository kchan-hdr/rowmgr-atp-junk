using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using static ROWM.Dal.OwnerDto2;
using static ROWM.Dal.ParcelOwnershipHelper;

namespace ROWM.Dal
{
    public class OwnershipRepository
    {
        #region ctor
        private readonly ROWM_Context _ctx;

        readonly OwnerRepository _repo;

        readonly ParcelOwnershipHelper _helper;

        public OwnershipRepository(ROWM_Context c, OwnerRepository r, ParcelOwnershipHelper h) => (_ctx, _repo, _helper) = (c, r, h);
        #endregion

        IQueryable<VestedOwner> ActiveVestedOwners() => _ctx.VestedOwner.Where(ox => !ox.IsDeleted);

        public async Task<Owner> GetOrCreateOwner(OwnershipRequest o)
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
          
            if (myLord == null)
                myLord = await _helper.AddOwner(o.Name, o.Address, o.OwnerType);


            return myLord;
        }

        public async Task<Owner> ChangeOwner(Owner owner, OwnershipRequest o)
        {
            var parcels = new List<Guid>(await _helper.GetParcelsByApn(o.Parcels));
            //parcels.AddRange(await _helper.GetParcelsByAcquisitionUnit(o.AcquisitionUnits));
            if (!await SplitAcquisitionUnit(o.Parcels))
            {
                throw new Exception("Error while splitting acquisition unit.");
            }

            _ = await EmptyVestedOwner(parcels);

            return await _helper.NewOwnership(parcels, owner.OwnerId);
        }

        public async Task<bool>  EmptyVestedOwner(IEnumerable<Guid> parcels)
        {
            bool anyDeleted = false;
            foreach (var parcel in parcels)
            {
                var myVO = await _ctx.VestedOwner.Where(ox => !ox.IsDeleted).FirstOrDefaultAsync(ox => ox.ParcelId == parcel);
                if (myVO != null)
                {
                    myVO.IsDeleted = true;
                    myVO.LastModified = DateTime.UtcNow;
                    myVO.ModifiedBy = "TCADOwnerRepo";
                    anyDeleted = true;
                }
            }
            await _ctx.SaveChangesAsync();
            return anyDeleted;
        }

        public async Task<VestedOwnerDto> GetVestedOwnerForParcel(string pid)
        {
            var myParcel = await _repo.GetParcel(pid);
            VestedOwner owner = await ActiveVestedOwners().FirstOrDefaultAsync(ox => ox.ParcelId == myParcel.ParcelId);
            return owner != null ? new VestedOwnerDto(owner) : null;
        }

        public async Task<IEnumerable<VestedOwnerDto>> FindVestedOwner(string name)
        {
            IEnumerable<VestedOwner> owners = await ActiveVestedOwners().Where(ox => ox.VestedOwnerName.Contains(name)).ToArrayAsync();
            return owners.Select(owner => new VestedOwnerDto(owner));
        }

        public async Task<VestedOwnerDto> AddVestedOwner(string name, string addr, Guid pid, string ownerType = "")
        {
            if (pid == Guid.Empty)
                throw new ArgumentOutOfRangeException(nameof(pid), "PID cannot be null or empty Guid.");

            try
            {
                // Check any old active vestOwners and retire them all
                var oldList = await ActiveVestedOwners().Where(ox => ox.ParcelId == pid).ToListAsync();

                foreach (var old in oldList)
                {
                    old.IsDeleted = true;
                    old.LastModified = DateTime.UtcNow;
                    old.ModifiedBy = "VestedOwnerRepo";
                }

                var vestedOwner = _ctx.VestedOwner.Add(new VestedOwner
                {
                    VestedOwnerName = name,
                    VestedOwnerAddress = addr,
                    OwnerType = ownerType,
                    IsDeleted = false,
                    Created = DateTimeOffset.UtcNow,
                    LastModified = DateTimeOffset.UtcNow,
                    ModifiedBy = "VestedOwnerRepo",
                    ParcelId = pid,
                    //TitleDocumentId = null
                });

                await _ctx.SaveChangesAsync();

                return new VestedOwnerDto(vestedOwner);
            }
            catch (DbUpdateException ex)
            {
                throw new ApplicationException("Error saving changes to the database", ex);
            }
            catch (Exception ex)
            {
                throw new ApplicationException("An error occurred", ex);
            }

        }


        // split acquisition unit
        public async Task<bool> SplitAcquisitionUnit(IEnumerable<string> parcels)
        {
            List<Guid> parcelIds = new List<Guid>();

            foreach (string parcel in parcels)
            {
                var myParcel = await _repo.GetParcel(parcel);
                if (myParcel != null)
                {
                    parcelIds.Add(myParcel.ParcelId);
                }
            }

            var pids = string.Join(",", parcelIds.Select(px => $"'{px}'"));

            var q = from a in _ctx.AcqParcel
                    where parcelIds.Contains(a.ParcelId)
                    select a;
            var acqs = await q.ToArrayAsync();
            // var query = $"SELECT DISTINCT acq.* FROM[ROWM].[Acquisition_Parcel] p1 left join[ROWM].[Acquisition_Parcel] acq ON p1.[Acquisition_Unit_No] = acq.[Acquisition_Unit_No] WHERE p1.[ParcelId] in ({pids})";

            // are all parcels belong to the same unit
            //var acqParcel = _ctx.Database.SqlQuery<AcqParcel>(query);
            //var acqs = await acqParcel.ToArrayAsync();

            if (!acqs.Any())
                return false;   // not found
            if (acqs.Select(ax => ax.Acquisition_Unit_No).Distinct().Count() > 1)
                return false;   // this is unusual. not supported (changing owners from different acq units

            var unitNumber = acqs.First().Acquisition_Unit_No;
            if (acqs.Count() == parcels.Count())
                return true;    // All parcels within the acq unit is selected

            // do split then
            var kill = await _ctx.AcqParcel.Where(kx => kx.Acquisition_Unit_No == unitNumber).ToListAsync();
            foreach (var p in kill)
            {
                // get extention of exisitng acq unit and prepare for the suffix number increase
                int dotIndex = p.Acquisition_Unit_No.LastIndexOf('.');
                int extention = dotIndex >= 0 ? int.Parse(p.Acquisition_Unit_No.Substring(dotIndex + 1)) : -1;
                int e1 = extention + 2;
                int e2 = extention + 3;
                string s1 = "." + e1.ToString("D2");
                string s2 = "." + e2.ToString("D2");

                // create new AcqParcel
                var nkey = new AcqParcel
                {
                    Acquisition_Unit_No = p.Acquisition_Unit_No.Split('.')[0] + (parcelIds.Contains(p.ParcelId) ? s1 : s2),
                    ParcelId = p.ParcelId,
                    //TCAD_PROP_ID = p.TCAD_PROP_ID
                };

                _ctx.AcqParcel.Remove(p);
                _ctx.AcqParcel.Add(nkey);
            }

            return await _ctx.SaveChangesAsync() > 0;
        }

    }

    #region Dtos
    public class OwnershipRequest
    {
        public Guid? OwnerId { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string OwnerType { get; set; }

        //public IEnumerable<string> AcquisitionUnits { get; set; }   // ATP uses funny acq units
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
    public class VestedOwnerDto
    {
        public Guid OwnerId { get; set; }
        public string PartyName { get; set; }
        public string OwnerAddress { get; set; }
        public IEnumerable<ParcelMinHeaderDto> OwnedParcel { get; set; }

        public VestedOwnerDto(VestedOwner o)
        {
            OwnerId = o.VestedOwnerId;
            PartyName = o.VestedOwnerName;
            OwnerAddress = o.VestedOwnerAddress;
        }
    }
    #endregion
}


