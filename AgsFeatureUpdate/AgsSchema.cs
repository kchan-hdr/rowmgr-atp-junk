using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace geographia.ags
{
    internal class AgsSchema
    {
        readonly FeatureService_Base _service;
        internal bool hasSchema;
        static SemaphoreSlim mutux = new SemaphoreSlim(1,1);
        Dictionary<string, long> _layers;

        internal AgsSchema(FeatureService_Base service) => _service = service;

        internal async Task<int> GetId(string layerName)
        {
            if (!hasSchema)
            {
                mutux.Wait();
                if (!hasSchema)
                    hasSchema = await GetLayers();
                mutux.Release();
            }

            return (int)_layers[layerName];
        }

        async Task<bool> GetLayers()
        {
            var info = await _service.Layers<AgsInfo>();

            _layers = new Dictionary<string, long>(); // (list.Select(lx => new KeyValuePair<string, long>(lx.Name.ToLower(), lx.Id)));

            var list = new List<IdInfo>();
            list.AddRange(info.Layers);
            list.AddRange(info.Tables);

            _layers = list.Distinct(new IdInfo())
                .ToDictionary(i => i.Name.ToLower(), i => i.Id);

            //foreach (var lx in list)
            //{
            //    var n = lx.Name.ToLower();
            //    if ( !_layers.ContainsKey(n))
            //        _layers.Add(n, lx.Id);
            //}

            return _layers.Count > 0;
        }

        #region layer desc
        public class AgsInfo
        {
            public string Description { get; set; }
            public IdInfo[] Layers { get; set; }
            public IdInfo[] Tables { get; set; }
        }

        public class IdInfo : EqualityComparer<IdInfo>
        {
            public long Id { get; set; }
            public string Name { get; set; }

            // equality
            public override bool Equals(IdInfo x, IdInfo y) => x.Name == y.Name;
            public override int GetHashCode(IdInfo obj) => (obj.Name).GetHashCode();
        }
        #endregion
    }
}
