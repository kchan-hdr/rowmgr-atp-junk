using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Relocation.DAL
{
    [Table("Relocation_Case_ContactLogs", Schema = "Austin")]
    public class RelocationCaseLog
    {
        public Guid ContactLogId { get; set; }
        public Guid RelocationCaseId { get; set; }
    }
}
