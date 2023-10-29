using System;

namespace ExpenseTracking.Dal.Entities.Dtos
{
    public class ExpenseHeaderDto
    {
        public Guid ExpenseId { get; set; }

        public string ExpenseTitle { get; set; }

        public double? ExpenseAmount { get; set; }

        public DateTimeOffset? SentDate { get; set; }

        public bool? IsFileAttached { get; set; }

        public ExpenseHeaderDto(Expense expense) {
            ExpenseId = expense.ExpenseId;
            ExpenseTitle = expense.ExpenseTitle;
            ExpenseAmount = expense.ExpenseAmount;  
            SentDate = expense.SentDate;
            IsFileAttached = expense.IsFileAttached;
        }

        public ExpenseHeaderDto(Guid expenseId, string expenseTitle, double? expenseAmount, DateTimeOffset? sentDate, bool? isFileAttached)
        {
            ExpenseId = expenseId;
            ExpenseTitle = expenseTitle;
            ExpenseAmount = expenseAmount;
            SentDate = sentDate;
            IsFileAttached = isFileAttached;
        }
    }
}
