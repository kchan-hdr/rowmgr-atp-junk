using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Austin_Costs
{
    public interface ICostEstimateRepository
    {
        IEnumerable<CostEstimate> Get(string key);
        Task<IEnumerable<ProjDetailsCostEstimate>> GetEx(string key);
    }

    public class CostEstimateRepository : ICostEstimateRepository
    {
        readonly CostEstimateContext _ctx;
        public CostEstimateRepository(CostEstimateContext c) => (_ctx) = (c);



        public IEnumerable<CostEstimate> Get(string parcelKey) => _ctx.Estimates.Where(ex => ex.PropId == parcelKey);

        public async Task<IEnumerable<ProjDetailsCostEstimate>> GetEx(string propId)
        {

            var k = await _ctx.AcquisitionKeys.SingleOrDefaultAsync(kx => kx.PropId == propId);

            if (k == null)
                return Enumerable.Empty<ProjDetailsCostEstimate>();

            var est = await _ctx.ProjDetailsCostEstimates.Where(kx => kx.AcqNo == k.AcqNo).ToListAsync();
            //    .SelectMany(kx => kx.Estimates);

            return est;
        }
    }
}
