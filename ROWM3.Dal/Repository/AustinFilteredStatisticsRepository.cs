using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;

namespace ROWM.Dal.Repository
{
    public class AustinFilteredStatisticsRepository : FilteredStatisticsRepository
    {
        readonly Lazy<List<AtpLookup>> _mapping;

        public AustinFilteredStatisticsRepository(ROWM_Context c) : base(c) 
        {
            _mapping = new Lazy<List<AtpLookup>>(MakeMapping);
            _baseParcels = new Lazy<IEnumerable<SubTotal>>(MakeBaseAcquisition);
        }

        public override async Task<IEnumerable<SubTotal>> SnapshotParcelStatus(int? part = null)
        {
            var q1 = await (
                    from p in ActiveParcels(part)
                           group p by p.ParcelStatusCode into psg
                           select new { psg.Key, c = psg.Count()  }
                    ).ToArrayAsync();

            var q = from p in q1
                    join sx in _mapping.Value on p.Key equals sx.Code
                    group p by sx.ChartCategory into psg
                    select new SubTotal { Title = psg.Key, Count = psg.Sum(px => px.c) };

            return from b in _baseParcels.Value
                   join psg in q on b.Title equals psg.Title into matq
                   from sub in matq.DefaultIfEmpty()
                   select new SubTotal { Title = b.Title, Caption = b.Caption, DomainValue = b.DomainValue, Count = sub?.Count ?? 0 };
        }

        IEnumerable<SubTotal> MakeBaseAcquisition()
        {
            var q = from sx in _mapping.Value
                    group sx by sx.ChartCategory into g
                    select new SubTotal
                    {
                        Title = g.Key,
                        Caption = g.Key,
                        DomainValue = g.Max(mx => mx.DomainValue).ToString(),
                        Count = 0
                    };

            return q.ToList();
        }

        List<AtpLookup> MakeMapping() =>
            _context.Database.SqlQuery<AtpLookup>("SELECT Code, ChartCategory, DomainValue FROM Austin.AcquisitionChart")
                .ToList();

        public class AtpLookup
        {
            internal string Code { get; set; }
            internal string ChartCategory { get; set; }
            internal int DomainValue { get; set; }
        }
    }
}
