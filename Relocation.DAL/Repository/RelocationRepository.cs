using com.hdr.rowmgr.Relocation;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROWM.Dal
{
    public class RelocationRepository
    {
        #region ctor
        readonly RelocationContext _context;
        public RelocationRepository(RelocationContext c)
        {
            _context = c ?? new RelocationContext("name=ROWM_Context");
        }
        #endregion

        public async Task<IEnumerable<IRelocationActivityType>> GetActivityTypes() =>
            await _context.RelocationActivity_Type
                .Where(t => t.IsActive)
                .ToListAsync();

        public async Task<bool> HasRelocation(Guid parcelId) => await _context.Relocations.AnyAsync(r => r.ParcelId == parcelId);
        public async Task<IParcelRelocation> GetRelocation(Guid parcelId) => 
            await _context.Relocations
                .SingleOrDefaultAsync(r => r.ParcelId == parcelId);

        public async Task<IEnumerable<IRelocationCase>> GetRelocationForParcel(Guid parcelId)
        {
            var q = from r in _context.Relocations.Include(pr => pr.Cases).AsNoTracking()
                    where r.ParcelId == parcelId
                    select r.Cases.Select(c => new
                    {
                        c.RelocationCaseId,
                        c.DisplaceeName,
                        c.DisplaceeType,
                        c.RelocationNumber,
                        c.RelocationType,
                        c.Status,
                        Docx = c.Documents.Select(dx => dx.DocumentId),
                        Logs = c.Logs.Select(lx => lx.ContactLogId),
                        Acts = c.ActionItems.Select(ax => ax.ActionItemId),
                        Steps = c.Activities.Select(a => a.ActivityCode).Distinct()
                    });

            var rr = await q.FirstOrDefaultAsync();
            if (rr is null || rr.Count() <= 0)
                return Array.Empty<RelocationCase>();

            var relx = rr?.Select(r =>
            {
                var rx = new RelocationCase
                {
                    DisplaceeName = r.DisplaceeName,
                    DisplaceeType = r.DisplaceeType,
                    RelocationCaseId = r.RelocationCaseId,
                    RelocationNumber = r.RelocationNumber,
                    RelocationType = r.RelocationType,
                    Status = r.Status
                };
                foreach (var a in r.Steps)
                    rx.Activities.Add(new RelocationDisplaceeActivity { ActivityCode = a });

                foreach (var d in r.Logs)
                    rx.Logs.Add(new ContactLog { ContactLogId = d });
                foreach (var d in r.Docx)
                    rx.Documents.Add(new Document { DocumentId = d });
                foreach (var a in r.Acts)
                    rx.ActionItems.Add(new ActionItem { ActionItemId = a });

                return rx;
            });

            return relx;
        }

        public async Task<IRelocationCase> GetRelocationCase(Guid caseId) => await _context.RelocationCases.FindAsync(caseId);
        public async Task<RelocationCase> GetRelocationCaseWithLogs(Guid caseId) => await _context.RelocationCases.Include(rc => rc.Logs).FirstOrDefaultAsync(rc => rc.RelocationCaseId == caseId);

        internal ParcelRelocation MakeNewRelocation => _context.Relocations.Add(new ParcelRelocation());

        internal async Task<IParcelRelocation> SaveRelocation(ParcelRelocation r)
        {
            if (_context.Entry<ParcelRelocation>(r).State == EntityState.Detached)
                _context.Entry<ParcelRelocation>(r).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
                return r;
            }
            catch (Exception e)
            {
                Trace.TraceError(e.Message);
                throw;
            }
        }

        #region contact logs
        public async Task<int> AttachLog(Guid caseId, Guid logId)
        {
            var c = await _context.RelocationCases.FindAsync(caseId);
            var l = await _context.ContactLogs.FindAsync(logId);

            c.Logs.Add(l);

            try
            {
                return await _context.SaveChangesAsync();
            }
            catch (Exception e)
            {
                Trace.TraceError(e.Message);
                throw;
            }
        }
        #endregion

        public async Task<int> AttachDoc(Guid rcId, Guid docId)
        {
            var relo = await _context.RelocationCases.FindAsync(rcId);
            var doc = await _context.Documents.FindAsync(docId);
            relo.Documents.Add(doc);
            return await _context.SaveChangesAsync();
        }

        #region actionItems
        public async Task<int> AttachActionItem(Guid caseId, Guid actId)
        {
            var c = await _context.RelocationCases.FindAsync(caseId);
            var a = await _context.ActionItems.FindAsync(actId);

            c.ActionItems.Add(a);

            try
            {
                return await _context.SaveChangesAsync();
            }
            catch (Exception e)
            {
                Trace.TraceError(e.Message);
                throw;
            }
        }
        #endregion

    }
}
