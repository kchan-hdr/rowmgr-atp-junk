using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExpenseTracking.Dal.Entities
{
    [Table("Expense_Activity", Schema = "ROWM")]
    public class ExpenseActivity
    {
        [Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public Guid ActivityId { get; set; }

        public DateTimeOffset ActivityDate { get; set; }

        [StringLength(100)]
        public string ActivityDescription { get; set; }

        public string ActivityNotes { get; set; }

        public Guid ParentExpenseId { get; set; }

        public Guid? ChildExpenseId { get; set; }

        public Guid? AgentId { get; set; }

        [ForeignKey(nameof(ParentExpenseId))]
        public virtual Expense ParentExpense { get; set; }

        [ForeignKey(nameof(ChildExpenseId))]
        public virtual Expense ChildExpense { get; set; }
    }
}
