using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROWM.Dal
{
    [Table("Vested_Owner", Schema ="Austin")]
    public class VestedOwner
    {
        [Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public Guid VestedOwnerId { get; set; }

        public Guid ParcelId { get; set; }

        [Column("Acquisition_Parcel_No")]
        public string AcqNo { get; set; }

        [Column("TCAD_PROP_ID")]
        public string Tcad { get; set; }

        [Column("Tracking_Number")]
        public string TrackingNumber { get; set; }

        [Column("Vested_Owner_Name")]
        public string VestedOwnerName { get; set; }

        [Column("Vested_Owner_Address")]
        public string VestedOwnerAddress { get; set; }

        [Column("Is_Verified")]
        public bool IsVerified { get; set; } = false;

        [Column("Source_Title_Document")]
        public Guid? TitleDocument { get; set; }

        [Column("Agent")]
        public Guid? AgentId { get; set; }

        public DateTimeOffset LastModified { get; set; } = DateTimeOffset.Now;
        public string ModifiedBy { get; set; }


        //navigation
        [ForeignKey(nameof(ParcelId))]
        virtual public Parcel ParentParcel { get; set; }

        [ForeignKey(nameof(AgentId))]
        virtual public Agent Agent { get; set; }
    }
}
