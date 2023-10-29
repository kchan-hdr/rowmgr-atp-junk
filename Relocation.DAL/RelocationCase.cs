using com.hdr.rowmgr.Relocation;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ROWM.Dal
{
    [Table("Relocation_Case", Schema = "Austin")]
    public partial class RelocationCase
    {
        [Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public Guid RelocationCaseId { get; set; }

        [ForeignKey(nameof(ParentRelocation))]
        public Guid ParentRelocationId { get; set; }
        public virtual ParcelRelocation ParentRelocation { get; set; }

        public Guid? AgentId { get; set; }

        public int RelocationNumber { get; set; }

        //[Column(TypeName ="nvarchar(20)")]
        public RelocationStatus Status { get; set; }

        //[Column(TypeName = "nvarchar(20)")]
        public DisplaceeType DisplaceeType { get; set; }

        //[Column(TypeName = "nvarchar(20)")]
        public RelocationType RelocationType { get; set; }

        public FinancialAssistType? FinancialAssistType { get; set; }

        public double? FinancialAssistAmount { get; set; }

        public virtual ICollection<RelocationEligibilityActivity> History { get; } = new HashSet<RelocationEligibilityActivity>();
        public virtual ICollection<RelocationDisplaceeActivity> Activities { get; } = new HashSet<RelocationDisplaceeActivity>();

        [StringLength(200)]
        public string DisplaceeName { get; set; }
        public Guid? ContactInfoId { get; set; }

        public virtual ICollection<ContactLog> Logs { get; } = new HashSet<ContactLog>();
        public virtual ICollection<Document> Documents { get; } = new HashSet<Document>();
    }
}
