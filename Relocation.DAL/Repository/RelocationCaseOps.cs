using com.hdr.rowmgr.Relocation;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Data.SqlTypes;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace ROWM.Dal
{
    #region noop
    public class RelocationCaseNoOp : IRelocationCaseOps
    {
        public Task<IParcelRelocation> AddActivity(Guid caseId, string activityCode, DisplaceeActivity atc, string desc, Guid agentId, DateTimeOffset date, string notes, int? money = null, bool? bValue = null)
        {
            throw new NotImplementedException();
        }

        public Task<IRelocationCase> AddRelocationCase(Guid parcelId, string displaceeName, RelocationStatus eligibility, DisplaceeType displaceeType, RelocationType reloType, FinancialAssistType ft, double fa)
        {
            throw new NotImplementedException();
        }

        public Task<IRelocationCase> AddRelocationCase(Guid parcelId, string displaceeName, string eligibility, string[] displaceeType, double? hs, double? rap)
        {
            throw new NotImplementedException();
        }

        public Task<IParcelRelocation> ChangeEligibility(Guid caseId, RelocationStatus eligibility, Guid agentId, DateTimeOffset date, string notes)
        {
            throw new NotImplementedException();
        }

        public Task<IEnumerable<IRelocationDisplaceeActivity>> GetActivities(Guid caseId)
        {
            throw new NotImplementedException();
        }

        public Task<IEnumerable<IRelocationActivityType>> GetActivityTypes()
        {
            throw new NotImplementedException();
        }

        public Task<IParcelRelocation> GetRelocation(Guid parcelId)
        {
            throw new NotImplementedException();
        }

        public Task<IRelocationCase> GetRelocationCase(Guid caseId)
        {
            throw new NotImplementedException();
        }

        public Task<IEnumerable<IRelocationCase>> GetRelocationCases(Guid parcelId)
        {
            throw new NotImplementedException();
        }
    }
    #endregion
    public class RelocationCaseOps : IRelocationCaseOps
    {
        readonly RelocationContext _context;
        readonly RelocationRepository _repo;

        public RelocationCaseOps(RelocationContext c, RelocationRepository r) => (_context, _repo) = (c, r);


        #region case eligibility
        public async Task<IEnumerable<IRelocationActivityType>> GetActivityTypes() => await _repo.GetActivityTypes();

        public async Task<IParcelRelocation> GetRelocation(Guid parcelId) => await _repo.GetRelocation(parcelId);

        public async Task<IEnumerable<IRelocationCase>> GetRelocationCases(Guid parcelId)
        {
            var p = await _repo.GetRelocation(parcelId) ?? throw new KeyNotFoundException(nameof(parcelId));
            return p.RelocationCases;
            //try
            //{
            //    var p = await _repo.GetRelocation(parcelId);
            //    if (p == null)
            //    {
            //        throw new KeyNotFoundException(nameof(parcelId));
            //    }
            //    return p.RelocationCases;
            //}
            //catch (KeyNotFoundException)
            //{
            //    return Enumerable.Empty<IRelocationCase>();
            //}
        }

        public async Task<IRelocationCase> GetRelocationCase(Guid caseId) => await _repo.GetRelocationCase(caseId);

        public async Task<IRelocationCase> AddRelocationCase(Guid parcelId, string displaceeName, string eligibility, string[] displaceeType, double? hs, double? rap)
        {
            if (string.IsNullOrWhiteSpace(displaceeName))
            {
                throw new ArgumentException("displaceeName cannot be null or empty.", nameof(displaceeName));
            }

            if (string.IsNullOrWhiteSpace(eligibility))
            {
                throw new ArgumentException("eligibility cannot be null or empty.", nameof(eligibility));
            }

            if (displaceeType == null || displaceeType.Length == 0)
            {
                throw new ArgumentException("displaceeType cannot be null or an empty list.", nameof(displaceeType));
            }

            var ( dt, rt ) = Decode(displaceeType[0]);

            var (ft, fa) = DecodeFinAssisType(hs, rap);

            if (eligibility.Equals("eligible on hold", StringComparison.InvariantCultureIgnoreCase))
                eligibility = "EligibleOnHold";

            if (Enum.TryParse<RelocationStatus>(eligibility, true, out var elig))
                //&& Enum.TryParse<DisplaceeType>(displaceeType[0], true, out var dType))
            {
                return await AddRelocationCase(parcelId, displaceeName, elig, dt, rt, ft, fa);
            }

            throw new KeyNotFoundException();
        }

        public async Task<IRelocationCase> AddRelocationCase(Guid parcelId, string displaceeName, RelocationStatus eligibility, DisplaceeType displaceeType, RelocationType reloType, FinancialAssistType ft, double fa)
        {
            if (!Enum.IsDefined(typeof(DisplaceeType), displaceeType))
            {
                throw new ArgumentException("Invalid displaceeType", nameof(displaceeType));
            }
            if (!Enum.IsDefined(typeof(RelocationType), reloType))
            {
                throw new ArgumentException("Invalid reloType", nameof(reloType));
            }
            var pr = await _repo.GetRelocation(parcelId) ?? _context.Relocations.Add(new ParcelRelocation { ParcelId = parcelId, Created = DateTimeOffset.UtcNow });
            var c = pr.AddCase(displaceeName, eligibility, displaceeType, reloType, ft, fa);
            await _repo.SaveRelocation(pr as ParcelRelocation);
            return c;
        }

        public async Task<IParcelRelocation> ChangeEligibility(Guid caseId, RelocationStatus eligibility, Guid agentId, DateTimeOffset date, string notes)
        {
            var c = await _context.RelocationCases.FindAsync(caseId) ?? throw new KeyNotFoundException(nameof(caseId));

            var origin = c.Status;
            c.Status = eligibility;

            c.History.Add(new RelocationEligibilityActivity
            {
                ActivityDate = date,
                AgentId = agentId,
                NewStatus = eligibility,
                OriginalStatus = origin,
                Notes = notes
            });

            var p = c.ParentRelocation;
            p.LastModified = DateTimeOffset.UtcNow;
            p.ModifiedBy = "ROWM ATP change";

            await _repo.SaveRelocation(p);

            return p;
        }

        #region decode relocation types 
        /// <summary>
        /// temporary mapping to ramp up ATP
        /// </summary>
        /// <param name="relocationType"></param>
        /// <returns></returns>
        public static (DisplaceeType,RelocationType) Decode(string relocationType)
        {
            DisplaceeType dt;
            RelocationType rt;

            switch( relocationType)
            {
                case "Non-Residential":
                    dt = DisplaceeType.BusinessTenant;
                    rt = RelocationType.Business;
                    break;

                case "Non-Residential (Landlord)":
                    dt = DisplaceeType.Landlord;
                    rt = RelocationType.Business;
                    break;

                case "Residential Owner":
                    dt = DisplaceeType.Owner;
                    rt = RelocationType.Residential;
                    break;

                case "Residential Tenant":
                    dt = DisplaceeType.ResidentialTenant;
                    rt = RelocationType.Residential;
                    break;

                case "Personal Property":
                    dt = DisplaceeType.PersonalProperty;
                    rt = RelocationType.PersonalProperty;
                    break;

                case "Business Owner":
                    dt = DisplaceeType.Owner;
                    rt = RelocationType.Business;
                    break;

                case "Business Tenant":
                    dt = DisplaceeType.BusinessTenant;
                    rt = RelocationType.Business;
                    break;

                case "OAS":
                    dt = DisplaceeType.OAS;
                    rt = RelocationType.OAS;
                    break;

                default:
                    dt = DisplaceeType.Owner;
                    rt = RelocationType.PersonalProperty;
                    //throw new KeyNotFoundException();
                    break;
            }

            return (dt, rt);
        }
        #endregion

        public static (FinancialAssistType, double) DecodeFinAssisType(double? hs, double? rap)
        {
            FinancialAssistType ft;
            double fa;

            if (hs.HasValue)
            {
                ft = FinancialAssistType.HousingSupplement;
                fa = hs.Value;
            }
            else if (rap.HasValue)
            {
                ft = FinancialAssistType.RentAssist;
                fa = rap.Value;
            }
            else
            {
                ft = FinancialAssistType.None;
                fa = 0;
            }

            return (ft, fa);
        }
        #endregion

        #region case activity
        public async Task<IParcelRelocation> AddActivity(Guid caseId, string activityCode, DisplaceeActivity act, string desc, Guid agentId, DateTimeOffset date, string notes, int? money, bool? bValue)
        {
            var c = await _context.RelocationCases.FindAsync(caseId) ?? throw new KeyNotFoundException(nameof(caseId));

            c.Activities.Add(new RelocationDisplaceeActivity
            {
                ActivityCode = activityCode,
                Activity = act,
                ActivityDescription = desc,
                ActivityDate = date,
                AgentId = agentId,
                Notes = notes,

                MoneyValue = money,
                BooleanValue = bValue
            });

            var p = c.ParentRelocation;
            p.LastModified = DateTimeOffset.UtcNow;
            p.ModifiedBy = "ROWM ATP tracking";

            await _repo.SaveRelocation(p);

            return p;
        }

        public async Task<IEnumerable<IRelocationDisplaceeActivity>> GetActivities(Guid caseId)
        {
            var c = await _context.RelocationCases.FindAsync(caseId) ?? throw new KeyNotFoundException(nameof(caseId));
            return c.Activities;
        }

        #endregion
    }
}
