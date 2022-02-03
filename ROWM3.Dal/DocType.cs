using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ROWM.Dal
{
    public class DocType
    {
        public string DocTypeName { get; set; }
        public string Description { get; set; }
        public int DisplayOrder { get; set; }
        public bool IsDisplayed { get; set; }
        public string ParentType { get; set; }
        public string DisplayCategory { get; set; }
        public string TitleHint { get; set; }

        [JsonIgnore]
        public string FolderPath { get; set; }

        [JsonIgnore]
        public Parcel_Status Milestone { get; set; }

        public DocType() { }
    }

    public class DocTypes : Collection<DocType>
    {
        #region static
        readonly Lazy<IEnumerable<DocType>> _Master;
        public IEnumerable<DocType> AllTypes => _Master.Value.OrderBy(dt => dt.DisplayOrder);
        public IEnumerable<DocType> Types => _Master.Value.OrderBy(dt => dt.DisplayOrder);
        public DocType Find(string n) => _Master.Value.SingleOrDefault(dt => dt.DocTypeName.Equals(n.Trim(), StringComparison.CurrentCultureIgnoreCase));
        public DocType Default => _Master.Value.Single(dt => dt.DocTypeName.Equals("Other"));
        #endregion

        readonly ROWM_Context _ctx;

        public DocTypes(ROWM_Context c)
        {
            _ctx = c;

            _Master = new Lazy<IEnumerable<DocType>>(DocTypeLoad, System.Threading.LazyThreadSafetyMode.PublicationOnly);
        }

        IEnumerable<DocType> DocTypeLoad() => _ctx.Document_Type.AsNoTracking()
            .Where(dt => dt.IsActive)
            .Select(dt => new DocType 
            { 
                DocTypeName = dt.DocTypeName
                , Description = dt.Description
                , DisplayOrder = dt.DisplayOrder
                , FolderPath = dt.FolderPath
                , IsDisplayed = string.IsNullOrEmpty(dt.ParentType)
                , Milestone = dt.Milestone 
                , ParentType = dt.ParentType
                , DisplayCategory = dt.DisplayCategory
                , TitleHint = dt.TitleHint
            })
            .ToArray();
    }
}
