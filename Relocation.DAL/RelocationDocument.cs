using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace Relocation.DAL
{
    [Table("Relocation_Case_Documents", Schema = "Austin")]
    public class RelocationCaseDocument
    {
        public Guid DocumentId { get; set; }
        public Guid RelocationCaseId { get; set; }
    }
}
