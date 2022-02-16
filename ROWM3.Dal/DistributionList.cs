using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ROWM.Dal
{
    [Table("Distribution_List", Schema ="Austin")]
    public class DistributionList
    {
        [Key]
        public Guid DistributionMemberId { get; private set; }
        public int ProjectPartId { get; private set; }
        public string Mail { get; private set; }
        public bool IsActive { get; private set; }
        public bool CcMode { get; private set; }
        public DateTimeOffset? LastSent { get; set; } = DateTimeOffset.UtcNow;
    }
}
