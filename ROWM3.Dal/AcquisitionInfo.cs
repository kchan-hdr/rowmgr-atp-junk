using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ROWM.Dal
{
    [Table("Acquisition_Parcel", Schema ="Austin")]
    public class AcqParcel
    {        
        public string Acquisition_Unit_No { get; set; }
        [Key]
        public Guid ParcelId { get; set; }
        public string TCAD_PROP_ID { get; set; }
    }
}
