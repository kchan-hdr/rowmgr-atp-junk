using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROWM.Dal
{
    [Table("Vested_Owner", Schema ="ROWM")]
    public class VestedOwner
    {
        [Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public Guid VestedOwnerId { get; set; }

        public Guid ParcelId { get; set; }

        [NotMapped][Column("Acquisition_Parcel_No")]
        public string AcqNo { get; set; }

        [NotMapped][Column("TCAD_PROP_ID")]
        public string Tcad { get; set; }

        [NotMapped][Column("Tracking_Number")]
        public string TrackingNumber { get; set; }

        [Column("OwnerName")]
        public string VestedOwnerName { get; set; }

        [Column("OwnerAddress")]
        public string VestedOwnerAddress { get; set; }

        [NotMapped][Column("Is_Verified")]
        public bool IsVerified { get; set; } = false;
        public bool IsDeleted { get; set; } = false;

        [NotMapped][Column("Source_Title_Document")]
        public Guid? TitleDocument { get; set; }

        [NotMapped][Column("Agent")]
        public Guid? AgentId { get; set; }

        public DateTimeOffset Created { get; set; } = DateTimeOffset.UtcNow;
        public DateTimeOffset LastModified { get; set; } = DateTimeOffset.UtcNow;
        public string ModifiedBy { get; set; }

        public string OwnerType { get; set; }

        //navigation
        [ForeignKey(nameof(ParcelId))]
        virtual public Parcel ParentParcel { get; set; }

        //[ForeignKey(nameof(AgentId))]
        //virtual public Agent Agent { get; set; }
    }
}
