using System.Collections.Generic;

namespace ExpenseTracking.Dal.Entities.Dtos
{

    public class ExpenseTypeDto
    {
        public string ExpenseTypeName { get; set; } = string.Empty;

        public string ExpenseCategory { get; set; } = string.Empty;

        public string Description { get; set; }

        public string FolderPath { get; set; }

        public int? DisplayOrder { get; set; }

        public double? TotalExpenseAmount { get; set; }

        public int TotalExpenseCount { get; set; }

        public bool? IsActive { get; set; }

        public IEnumerable<ExpenseHeaderDto> Expenses { get; } = new HashSet<ExpenseHeaderDto>();

        public ExpenseTypeDto() { }

        public ExpenseTypeDto(ExpenseType expenseType, IEnumerable<ExpenseHeaderDto> selectedExpenses, double totalExpenseAmount, int totalExpenseCount)
        {
            ExpenseTypeName = expenseType.ExpenseTypeName;
            ExpenseCategory = expenseType.ExpenseCategoryName;
            Description = expenseType.Description;
            FolderPath = expenseType.FolderPath;
            DisplayOrder = expenseType.DisplayOrder;
            IsActive = expenseType.IsActive;
            Expenses = selectedExpenses;
            TotalExpenseAmount = totalExpenseAmount;
            TotalExpenseCount = totalExpenseCount;
        }
    }
}