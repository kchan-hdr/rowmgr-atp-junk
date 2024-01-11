using ExpenseTracking.Dal.Entities;
using ExpenseTracking.Dal.Entities.Dtos;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;

namespace ExpenseTracking.Dal
{

    public class ExpenseTracking_Op : IExpenseTracking
    {
        readonly ExpenseContext _context;

        public ExpenseTracking_Op(ExpenseContext context)
        {
            _context = context;
        }

        public async Task<Expense> AddExpense(ExpenseRequestDto expenseRequestDto)
        {
            if (!expenseRequestDto.ValidateExpenseTypeName())
            {
                throw new ArgumentException("Invalid expense type.");
            }
            if (!expenseRequestDto.ValidateRecordedDateAndSentDate())
            {
                throw new ArgumentException("Recorded date cannot be before sent date.");
            }
            if (!expenseRequestDto.ValidateRequiredAmount())
            {
                throw new ArgumentException("Invalid expense amount.");
            }
            Expense expense = new Expense(expenseRequestDto);
            _context.Expense.Add(expense);
            await _context.SaveChangesAsync();
            return await GetExpense(expense.ExpenseId);
        }

        public async Task<Expense> GetExpense(Guid expenseId)
        {
            return await _context.Expense.SingleOrDefaultAsync(e => e.ExpenseId == expenseId)
                 ?? throw new KeyNotFoundException(nameof(expenseId));
        }

        public async Task<Expense> UpdateExpense(Guid expenseId, ExpenseRequestDto expenseRequestDto)
        {
            Expense expense = await GetExpense(expenseId);
            Expense updatedExpense = expense.Update(expenseRequestDto);
            await _context.SaveChangesAsync();
            return updatedExpense;
        }

        public async Task<Expense> DeleteExpense(Guid expenseId)
        {
            Expense expense = await GetExpense(expenseId);
            expense.IsDeleted = true;
            await _context.SaveChangesAsync();

            return expense;
        }

        public async Task<IEnumerable<ExpenseTypeDto>> GetExpenses(string parcelTCAD, string category)
        {
            if (string.IsNullOrEmpty(parcelTCAD) /*||  more validation conditions (not exist in Parcel table)  */)
            {
                throw new ArgumentException("Invalid parcelTCAD.", nameof(parcelTCAD));
            }

            var expenseTypes = _context.ExpenseType.AsNoTracking().Where(e => e.ExpenseCategoryName == category);

            var expenses = _context.Expense.AsNoTracking()
                .Where(e => e.ParcelTCAD == parcelTCAD)
                .Where(e => !e.IsDeleted)
                .Where(e => e.ExpenseType.ExpenseCategoryName == category)
                .GroupBy(e => e.ExpenseTypeName);

            var query = from et in expenseTypes
                        join e in expenses on et.ExpenseTypeName equals e.Key into types
                        from etx in types.DefaultIfEmpty()  
                        select new { expenseType = et, expenses = etx.ToList() };

            var myList = await query.OrderBy(et => et.expenseType.DisplayOrder).ToArrayAsync();
            return myList.Select(t => new ExpenseTypeDto(t.expenseType, t.expenses?.Select(e => new ExpenseHeaderDto(e)) ?? Enumerable.Empty<ExpenseHeaderDto>(), t.expenses?.Sum(e => e.ExpenseAmount) ?? 0M, t.expenses?.Count() ?? 0));
        }

        public async Task<IEnumerable<ExpenseTypeDto>> GetExpenses(Guid relocationCaseId, string category)
        {
            if (relocationCaseId == Guid.Empty /*||  more validation conditions (not exist in Parcel table)  */)
            {
                throw new ArgumentException("Invalid parcelTCAD.", nameof(relocationCaseId));
            }
            var expenseTypes = _context.ExpenseType.AsNoTracking().Where(e => e.ExpenseCategoryName == category);

            var expenses = _context.Expense.AsNoTracking()
                .Where(e => e.RelocationCaseId == relocationCaseId)
                .Where(e => !e.IsDeleted)
                .Where(e => e.ExpenseType.ExpenseCategoryName == category)
                .GroupBy(e => e.ExpenseTypeName);

            var query = from et in expenseTypes
                        join e in expenses on et.ExpenseTypeName equals e.Key into types
                        from etx in types.DefaultIfEmpty()
                        select new { expenseType = et, expenses = etx.ToList() };

            var myList = await query.OrderBy(et => et.expenseType.DisplayOrder).ToArrayAsync();
            return myList.Select(t => new ExpenseTypeDto(t.expenseType, t.expenses?.Select(e => new ExpenseHeaderDto(e)) ?? Enumerable.Empty<ExpenseHeaderDto>(), t.expenses?.Sum(e => e.ExpenseAmount) ?? 0, t.expenses?.Count() ?? 0));
        }

        static IEnumerable<ExpenseTypeDto> MakeExpenseTypeResult((ExpenseType, IList<Expense>)[] expenseT) =>
            expenseT.Select(t =>
                    new ExpenseTypeDto(t.Item1, t.Item2?.Select(e => new ExpenseHeaderDto(e)) ?? Enumerable.Empty<ExpenseHeaderDto>(), t.Item2?.Sum(e => e.ExpenseAmount) ?? 0, t.Item2?.Count() ?? 0));


        public async Task<IEnumerable<string>> GetExpensesTypeNames()
        {
            return await _context.ExpenseType.Select(et => et.ExpenseTypeName).ToListAsync();
        }

        public async Task<IEnumerable<string>> GetExpensesTypeNames(string category)
        {
            return await _context.ExpenseType
                .Where(et => et.ExpenseCategoryName == category)
                .Select(et => et.ExpenseTypeName).ToListAsync();
        }

        public async Task<IEnumerable<ExpenseHeaderDto>> GetExpensesByType(string parcelTCAD, string expenseType)
        {
            if (string.IsNullOrEmpty(parcelTCAD) /*||  more validation conditions (not exist in Parcel table)  */)
            {
                throw new ArgumentException("Invalid parcelTCAD.", nameof(parcelTCAD));
            }

            ExpenseType et = await _context.ExpenseType.FirstAsync(ex => ex.ExpenseTypeName == expenseType); // ?? throw new KeyNotFoundException(nameof(expenseType));
            if (et.IsActive == false)
            {
                throw new ArgumentException("Invalid expenseType. It is not active", nameof(expenseType));
            }

            var expenseHeaderDtos = et.Expenses
                .Where(expense => expense.ParcelTCAD == parcelTCAD && !expense.IsDeleted)
                .Select(expense => new ExpenseHeaderDto(expense))
                .ToList();

            return expenseHeaderDtos;
        }

        public async Task<IEnumerable<ExpenseHeaderDto>> GetExpensesByType(Guid relocationCaseId, string expenseType)
        {
            if (relocationCaseId == Guid.Empty /*||  more validation conditions (not exist in Parcel table)  */)
            {
                throw new ArgumentException("Invalid relo Id.", nameof(relocationCaseId));
            }

            ExpenseType et = await _context.ExpenseType.FirstOrDefaultAsync(ex => ex.ExpenseTypeName == expenseType) ?? throw new KeyNotFoundException(nameof(expenseType));
            if (et.IsActive == false)
            {
                throw new ArgumentException("Invalid expenseType. It is not active", nameof(expenseType));
            }

            var expenseHeaderDtos = et.Expenses
                .Where(expense => expense.RelocationCaseId == relocationCaseId && !expense.IsDeleted)
                .Select(expense => new ExpenseHeaderDto(expense))
                .ToList();

            return expenseHeaderDtos;
        }

        [Obsolete("not used. included in get expense type")]
        public async Task<decimal> CalculateExpenseByType(string parcelTCAD, string expenseType)
        {
            var expenseHeaderDtos = await GetExpensesByType(parcelTCAD, expenseType);

            var totalExpenseAmount = expenseHeaderDtos.Sum(dto => dto.ExpenseAmount ?? 0.0M);

            return totalExpenseAmount;
        }

        public async Task<decimal> CalculateExpenseByRelocationCase(Guid relocationCaseId)
        {
            if (relocationCaseId == Guid.Empty /*||  more validation conditions (not exist in RelocationCase table)  */)
            {
                throw new ArgumentException("Invalid parcelTCAD.", nameof(relocationCaseId));
            }
            var expenses = _context.Expense
                .Where(expense => expense.RelocationCaseId == relocationCaseId && !expense.IsDeleted)
                .Select(expense => expense.ExpenseAmount);
            
            if (await expenses.AnyAsync())
            {
                return expenses.Sum();
            }
            else
            {
                return 0;
            }
        }

        public async Task<decimal> CalculateExpenseByParcel(string parcelTCAD)
        {
            if (string.IsNullOrEmpty(parcelTCAD) /*||  more validation conditions (not exist in Parcel table)  */)
            {
                throw new ArgumentException("Invalid parcelTCAD.", nameof(parcelTCAD));
            }

            var expenses = _context.Expense
                .Where(expense => expense.ParcelTCAD == parcelTCAD && !expense.IsDeleted)
                .Select(expense => expense.ExpenseAmount);

            if (await expenses.AnyAsync())
            {
                return expenses.Sum();
            }
            else
            {
                return 0;
            }
        }

        public async Task<decimal> CalculateReloExpenseByParcel(string parcelTCAD)
        {
            if (string.IsNullOrEmpty(parcelTCAD) /*||  more validation conditions (not exist in Parcel table)  */)
            {
                throw new ArgumentException("Invalid parcelTCAD.", nameof(parcelTCAD));
            }
            var expenses = _context.Expense
                .Where(expense => expense.ParcelTCAD == parcelTCAD && expense.RelocationCaseId.HasValue && expense.RelocationCaseId != Guid.Empty && !expense.IsDeleted)
                .Select(expense => expense.ExpenseAmount);

            if (await expenses.AnyAsync())
            {
                return expenses.Sum();
            }
            else
            {
                return 0;
            }
        }
    }
}