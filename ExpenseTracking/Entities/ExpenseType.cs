using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ExpenseTracking.Dal.Entities
{
    [Table("Expense_Type", Schema = "ROWM")]
    public class ExpenseType
    {
        [Key]
        [StringLength(50)]
        public string ExpenseTypeName { get; set; } = string.Empty;

        [StringLength(50), Required, ForeignKey(nameof(Category))]
        public string ExpenseCategoryName { get; set; } = string.Empty;
        public virtual ExpenseCategory Category { get; set; }

        public string Description { get; set; }

        [StringLength(400)]
        public string FolderPath { get; set; }

        public int? DisplayOrder { get; set; }

        public bool? IsActive { get; set; }

        public virtual ICollection<Expense> Expenses { get; } = new HashSet<Expense>(); // Collection navigation containing dependents
    }
}