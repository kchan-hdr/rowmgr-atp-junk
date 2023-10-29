using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Relocation.DAL
{
    [Table("Relocation_Case_Documents", Schema = "Austin")]
    public class RelocationCaseDocument
    {
        public Guid DocumentId { get; set; }
        public Guid RelocationCaseId { get; set; }
    }
}
