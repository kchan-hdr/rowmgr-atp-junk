using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Data.Entity.Validation;
using System.Linq;
using System.Threading.Tasks;

namespace ROWM.Dal
{
    public interface IOwnershipHelper
    {
        Task<Owner> GetOwner(Guid id);
        Task<IEnumerable<Owner>> AllOwners();
        Task<Owner> AddOwner(string name, string addr, string ownerType = "");
        Task<IEnumerable<Ownership>> GetOwners(Guid parcelId);
        Task<Owner> NewOwnership(IEnumerable<Guid> parcelId, Guid ownerId);

        Task<Guid[]> GetParcelsByAcquisitionUnit(IEnumerable<string> trackno);
        Task<Guid[]> GetParcelsByApn(IEnumerable<string> apn);
    }

    public class ParcelOwnershipHelper : IOwnershipHelper
    {
        #region ctor
        readonly ROWM_Context _ctx;
        public ParcelOwnershipHelper(ROWM_Context c) => (_ctx) = (c);
        #endregion

        public Task<Owner> GetOwner(Guid id) => _ctx.Owner.FindAsync(id);
        public async Task<Owner> AddOwner(string name, string addr, string ownerType = "" )
        {
            var o = _ctx.Owner.Add(new Owner
            {
                PartyName = name,
                OwnerAddress = addr,
                OwnerType = ownerType,
                IsDeleted = false,
                Created = DateTimeOffset.UtcNow,
                ModifiedBy = nameof(ParcelOwnershipHelper)
            });

            await _ctx.SaveChangesAsync();

            return o;
        }

        public async Task<IEnumerable<Owner>> AllOwners()
        {
            var q = from ow in _ctx.Owner
                    where !ow.IsDeleted
                    select new
                    {
                        ow.OwnerId,
                        ow.PartyName,
                        ow.OwnerAddress,
                        Ownership = ow.Ownership.Select(os => new
                        {
                            os.Ownership_t,
                            Parcel = new
                            {
                                os.Parcel.ParcelId,
                                os.Parcel.Assessor_Parcel_Number,
                                os.Parcel.Tracking_Number
                            }
                        })
                    };

            return (await q.ToArrayAsync())
                .Select(ow => new Owner
                {
                    OwnerId = ow.OwnerId,
                    PartyName = ow.PartyName,
                    OwnerAddress = ow.OwnerAddress,
                    Ownership = ow.Ownership.Select(os => new Ownership
                    {
                        Ownership_t = os.Ownership_t,
                        Parcel = new Parcel
                        {
                            ParcelId = os.Parcel.ParcelId,
                            Assessor_Parcel_Number = os.Parcel.Assessor_Parcel_Number,
                            Tracking_Number = os.Parcel.Tracking_Number
                        }
                    }).ToArray()
                });
        }

        // including old/changed ownerships
        public async Task<IEnumerable<Ownership>> GetOwners(Guid parcelId)
        {
            var q = from p in _ctx.Parcel
                    where p.ParcelId == parcelId
                    select
                    p.Ownership.Select(ox => new
                    {
                        ox.Ownership_t,
                        Owner = new
                        {
                            ox.Owner.OwnerId,
                            ox.Owner.PartyName,
                            ox.Owner.OwnerAddress,
                            ox.Owner.OwnerType
                        },
                        Parcel = new
                        {
                            p.ParcelId,
                            p.Assessor_Parcel_Number,
                            p.Tracking_Number
                        }
                    });

            return (await q.ToArrayAsync())
                .SelectMany(p => p.Select(oxx => new Ownership
                {
                    Ownership_t = oxx.Ownership_t,
                    OwnerId = oxx.Owner.OwnerId,
                    Owner = new Owner
                    {
                        OwnerId = oxx.Owner.OwnerId,
                        PartyName = oxx.Owner.PartyName,
                        OwnerAddress = oxx.Owner.OwnerAddress,
                        OwnerType = oxx.Owner.OwnerType
                    },
                    ParcelId = oxx.Parcel.ParcelId,
                    Parcel = new Parcel
                    {
                        ParcelId = oxx.Parcel.ParcelId,
                        Assessor_Parcel_Number = oxx.Parcel.Assessor_Parcel_Number,
                        Tracking_Number = oxx.Parcel.Tracking_Number
                    }                    
                }));
        }

        // change ownership
        public async Task<Owner> NewOwnership(IEnumerable<Guid> parcelId, Guid ownerId)
        {
            if (!parcelId.Any())
                throw new ArgumentOutOfRangeException(nameof(parcelId));

            var dt = DateTimeOffset.UtcNow;
            var myParcel = await _ctx.Parcel.Include(px => px.Ownership.Select(ox => ox.Owner)).Where(px => parcelId.Contains(px.ParcelId)).ToListAsync();

            foreach( var p in myParcel)
            {
                foreach (var os in p.Ownership)
                {
                    // retire old 
                    if (!os.Owner.IsDeleted)
                    {
                        os.Ownership_t = (int)Ownership.OwnershipType.Relinquished;
                        os.LastModified = dt;
                        os.ModifiedBy = nameof(ParcelOwnershipHelper);
                    }
                }

                // add new owner
                p.Ownership.Add(new Ownership
                {
                    OwnerId = ownerId,
                    Ownership_t = (int)Ownership.OwnershipType.Primary,
                    Created = dt,
                });
            }

            try
            {
                await _ctx.SaveChangesAsync();
            }
            catch( DbEntityValidationException validation)
            {
                throw;
            }
            catch( DbUpdateException update)
            {
                throw;
            }   
            return await GetOwner(ownerId);
        }

        #region private records
        public class AcqParcel
        {
            public string Acquisition_Parcel_No { get; set; }
            public string TCAD_PROP_ID { get; set; }
        }
        #endregion

        #region acquisition units
        public Task<Guid[]> GetParcelsByAcquisitionUnit(IEnumerable<string> trackno) => trackno == null ? Task.FromResult(Enumerable.Empty<Guid>().ToArray()) : _ctx.Parcel.Where(px => trackno.Contains(px.Tracking_Number)).Select(px => px.ParcelId).ToArrayAsync();
        public Task<Guid[]> GetParcelsByApn(IEnumerable<string> apn) => apn == null ? Task.FromResult(Enumerable.Empty<Guid>().ToArray()) : _ctx.Parcel.Where(px => apn.Contains(px.Assessor_Parcel_Number)).Select(px => px.ParcelId).ToArrayAsync();
        #endregion

    }
}
