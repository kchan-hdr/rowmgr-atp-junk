using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ROWM.Reports_Ex
{
    public class NoReports : IRowmReports
    {
        #region init
        readonly IEnumerable<ReportDef> _reports = new List<ReportDef>();
        #endregion
        public Task<ReportPayload> GenerateReport(ReportDef d)
        {
            throw new NotImplementedException();
        }

        public Task<ReportPayload> GenerateReport(string reportCode)
        {
            throw new NotImplementedException();
        }

        public Task<IEnumerable<ReportDef>> GetReports() => Task.FromResult(_reports);
    }
}
