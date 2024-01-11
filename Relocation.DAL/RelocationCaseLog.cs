using System;
using System.ComponentModel.DataAnnotations.Schema;

namespace Relocation.DAL
{
    [Table("Relocation_Case_ContactLogs", Schema = "Austin")]
    public class RelocationCaseLog
    {
        public Guid ContactLogId { get; set; }
        public Guid RelocationCaseId { get; set; }
    }
}
