namespace ROWM.Dal
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("ROWM.Document_Type")]
    public partial class Document_Type
    {
        [Key]
        [StringLength(200)]
        public string DocTypeName { get; set; }

        public string Description { get; set; }

        [StringLength(400)]
        public string FolderPath { get; set; }

        public int DisplayOrder { get; set; }

        public bool IsActive { get; set; }


        [ForeignKey("Milestone")]
        public string MilestoneCode { get; set; }
        public virtual Parcel_Status Milestone { get; set; }

        [Obsolete("not used")]
        public string ParentType { get; set; }

        public string DisplayCategory { get; set; }
        public string TitleHint { get; set; }
    }
}
