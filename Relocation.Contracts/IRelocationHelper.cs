using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace com.hdr.rowmgr.Relocation
{
    public interface IRelocationHelper
    {
        Task<IParcelRelocation> GetRelocation(Guid parcelId);
        Task<IEnumerable<IRelocationCase>> GetRelocationCases(Guid parcelId);
        Task<IParcelRelocation> AddRelocationCase(Guid parcelId, string displaceeName, RelocationStatus eligibility, DisplaceeType displaceeType, RelocationType reloType);

        Task<IEnumerable<IRelocationActivityType>> GetActivityTypes();
    }

    public enum RelocationStatus { Unk = 0, Eligible = 1, EligibleOnHold, Optional, Ineligible }

    public enum DisplaceeType { Landlord = 1, Owner, BusinessTenant, ResidentialTenant, PersonalProperty, OAS }
    public enum RelocationType { Business = 1, Residential, PersonalProperty, OAS /* Outdoor advertisement sign */ }
    public enum FinancialAssistType { Unk = 0, HousingSupplement = 1, RentAssist, None }

}
