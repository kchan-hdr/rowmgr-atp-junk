using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Austin_Costs
{
    [Table("Cost_Estimate_Parcel", Schema ="Austin")]
    public class AcquisitionKey
    {
        [Key, Column("Acquisition_Parcel_No", Order =1)]
        public string AcqNo { get; private set; }

        [Key, Column("TCAD_PROP_ID", Order =2)]
        public string PropId { get; private set; }


        //[ForeignKey(nameof(AcqNo))]
        //virtual public ICollection<ProjDetailsCostEstimate> Estimates { get; private set; }

        // alias
        public string TrackingNumber => this.PropId;
    }
}
