using ROWM.Dal;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace WeeklyReports
{
    public class Recipients
    {
        readonly ROWM_Context _Context;
        public Recipients(ROWM_Context c) => _Context = c;

        internal async Task<IEnumerable<DistributionList>> GetRecipients(int part) =>
            await _Context.DistributionList.Where(l => l.ProjectPartId == part && l.IsActive)?.ToArrayAsync() ?? throw new KeyNotFoundException($"unknown line {part}");

        internal async Task Sent(int part, DateTimeOffset dt)
        {
            var r = await GetRecipients(part);
            foreach (var rx in r)
                rx.LastSent = dt;

            try
            {
                await _Context.SaveChangesAsync();
            }
            catch (System.Data.Entity.Infrastructure.DbUpdateException e)
            {
                Trace.TraceError(e.Message);
            }
        }
    }
}
