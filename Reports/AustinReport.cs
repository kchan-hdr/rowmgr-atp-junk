using ROWM.Dal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ROWM.Reports
{
    public class AustinReport : IRowmReports
    {
        readonly ROWM_Context _context;
        readonly OwnerRepository _repo;
        readonly IStatisticsRepository _stats;
        public AustinReport(ROWM_Context c, OwnerRepository r, IStatisticsRepository s) => (_context, _repo, _stats) = (c, r, s);

        public async Task<ReportPayload> GenerateReport(ReportDef d)
        {
            _ = d ?? throw new ArgumentNullException("missing ReportDef");

            var code = d.ReportCode;

            if (code.StartsWith("en") && int.TryParse(code.Substring(2), out var project))
            {
                var parts = _context.ProjectParts.FirstOrDefault(px => px.ProjectPartId == project);

                var title = parts.Caption;
                var printDate = $"as of {DateTime.Today.ToLongDateString()}";
                var titleDate = DateTime.Today.ToString("M.dd.yyyy");

                var data = await _repo.GetEngagement(project);
                var pieData = await _stats.Snapshot("engagement", project);

                //var e = new ExcelExport.EngagementExport(data);
                //var bytes = e.Export();
                var e = new ExcelExport.GeneratedClass();
                var bytes = e.CreatePackage(title, printDate, data, pieData);
                var p = new ReportPayload { Content = bytes, Filename = $"{d.Caption}_{titleDate}.xlsx", Mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" };
                return p;
            }
                

            throw new KeyNotFoundException($"unknown report {code}");
        }

        public IEnumerable<ReportDef> GetReports()
        {
            var parts = _context.ProjectParts
                .Where(pp => pp.IsActive)
                .OrderBy(pp => pp.DisplayOrder)
                .AsEnumerable()
                .Select(pp => new ReportDef { Caption = $"{pp.Caption} Community Engagement Report", DisplayOrder = pp.DisplayOrder ?? 0, ReportCode = $"en{pp.ProjectPartId}" })
                .ToArray();

            return parts;
        }
    }
}
