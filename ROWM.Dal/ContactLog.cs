﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ROWM.Dal
{
    [Table("ContactLog", Schema ="ROWM")]
    public class ContactLog
    {
        // public enum Channel { InPerson = 1, Phone, TextMessage, Email, Letter, NotesToFile, Followup = 10, Research = 20 }

        [Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public Guid ContactLogId { get; set; }

        public DateTimeOffset DateAdded { get; set; }

        [ForeignKey("ContactAgent")]
        public Guid ContactAgentId { get; set; }
        public virtual Agent ContactAgent { get; set; }

        [StringLength(20)]
        public string ContactChannel { get; set; }      // fk Channel_Master

        [StringLength(20)]
        public string ProjectPhase { get; set; }        // fk Purpose_Master

        [StringLength(200)]
        public string Title { get; set; }

        [StringLength(int.MaxValue)]
        public string Notes { get; set; }

        //
        public virtual ICollection<ContactInfo> Contacts { get; set; }
        public virtual ICollection<Parcel> Parcels { get; set; }

        // audit
        [Required]
        public DateTimeOffset Created { get; set; }
        public DateTimeOffset LastModified { get; set; }
        [StringLength(50)]
        public string ModifiedBy { get; set; }
    }
}
