using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExpenseTracking.Dal.Entities
{
    [Table("Expense_Category", Schema = "ROWM")]
    public class ExpenseCategory
    {
        [Key]
        [StringLength(50)]
        public string ExpenseCategoryName { get; set; } = string.Empty;

        public string Description { get; set; }

        public int? DisplayOrder { get; set; }

        public bool? IsActive { get; set; }

        public virtual ICollection<ExpenseType> ExpenseTypes { get; } = new HashSet<ExpenseType>();
    }
}
