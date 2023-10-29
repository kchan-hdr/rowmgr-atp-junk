using com.hdr.rowmgr.Relocation;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ROWM.Dal
{
    public interface IRelocationCaseOps
    {
        Task<IParcelRelocation> GetRelocation(Guid parcelId);
        Task<IEnumerable<IRelocationCase>> GetRelocationCases(Guid parcelId);
        Task<IRelocationCase> GetRelocationCase(Guid caseId);
        Task<IRelocationCase> AddRelocationCase(Guid parcelId, string displaceeName, string eligibility, string[] displaceeType, double? hs, double? rap);
        Task<IRelocationCase> AddRelocationCase(Guid parcelId, string displaceeName, RelocationStatus eligibility, DisplaceeType displaceeType, RelocationType reloType, FinancialAssistType ft, double fa);
        
        Task<IParcelRelocation> ChangeEligibility(Guid caseId, RelocationStatus eligibility, Guid agentId, DateTimeOffset date, string notes);

        Task<IEnumerable<IRelocationDisplaceeActivity>> GetActivities(Guid caseId);
        Task<IParcelRelocation> AddActivity(Guid caseId, string activityCode, DisplaceeActivity act, string desc, Guid agentId, DateTimeOffset date, string notes, int? money = null, bool? bValue = null);

        Task<IEnumerable<IRelocationActivityType>> GetActivityTypes();
    }
}
