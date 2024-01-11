using ExpenseTracking.Dal.Entities;
using ExpenseTracking.Dal.Entities.Dtos;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExpenseTracking.Dal
{

    public interface IExpenseTracking
    {
        Task<Expense> AddExpense(ExpenseRequestDto expenseRequestDto); //Create a new Expense

        Task<Expense> GetExpense(Guid expenseId);

        Task<Expense> UpdateExpense(Guid expenseId, ExpenseRequestDto expenseRequestDto);

        Task<Expense> DeleteExpense(Guid expenseId);

        Task<IEnumerable<ExpenseTypeDto>> GetExpenses(string parcelTCAD, string categoty); //Return a list of Expense Type with according expense under each type for a parcel
        Task<IEnumerable<ExpenseTypeDto>> GetExpenses(Guid relocationId, string category); // relocation expenses

        Task<IEnumerable<string>> GetExpensesTypeNames(); //Return all the Expense type name

        Task<IEnumerable<ExpenseHeaderDto>> GetExpensesByType(string parcelTCAD, string expenseType); //Get a list of expense for a paticular expense type for a parcel

        Task<decimal> CalculateExpenseByType(string parcelTCAD, string expenseType);//Type total expense

        Task<decimal> CalculateExpenseByRelocationCase(Guid relocationCaseId);//Relo case total expense

        Task<decimal> CalculateExpenseByParcel(string parcelTCAD); //Parcel total expense
        Task<decimal> CalculateReloExpenseByParcel(string parcelTCAD);
    }

    public class ExpenseTracking_NoOp : IExpenseTracking
    {
        public Task<Expense> AddExpense(ExpenseRequestDto expenseRequestDto)
        {
            throw new NotImplementedException("AddExpense is not implemented.");
        }

        public Task<Expense> GetExpense(Guid expenseId)
        {
            throw new NotImplementedException("GetExpense is not implemented.");
        }

        public Task<Expense> UpdateExpense(Guid expenseId, ExpenseRequestDto expenseRequestDto)
        {
            throw new NotImplementedException("UpdateExpense is not implemented.");
        }

        public Task<Expense> DeleteExpense(Guid expenseId)
        {
            throw new NotImplementedException("DeleteExpense is not implemented.");
        }

        public Task<IEnumerable<ExpenseTypeDto>> GetExpenses(string parcelTCAD, string category)
        {
            throw new NotImplementedException("GetExpensesType is not implemented.");
        }

        public Task<IEnumerable<string>> GetExpensesTypeNames()
        {
            throw new NotImplementedException("GetExpensesType is not implemented.");
        }

        public Task<IEnumerable<ExpenseHeaderDto>> GetExpensesByType(string parcelTCAD, string expenseType)
        {
            throw new NotImplementedException("GetExpensesByType is not implemented.");
        }

        public Task<decimal> CalculateExpenseByType(string parcelTCAD, string expenseType)
        {
            throw new NotImplementedException("CalculateExpenseByType is not implemented.");
        }

        public Task<decimal> CalculateExpenseByRelocationCase(Guid relocationCaseId)
        {
            throw new NotImplementedException("CalculateExpenseByRelocationCase is not implemented.");
        }

        public Task<decimal> CalculateExpenseByParcel(string parcelTCAD)
        {
            throw new NotImplementedException("CalculateExpenseByParcel is not implemented.");
        }
        public Task<decimal> CalculateReloExpenseByParcel(string parcelTCAD)
        {
            throw new NotImplementedException("CalculateExpenseByParcel is not implemented.");
        }

        public Task<IEnumerable<ExpenseTypeDto>> GetExpenses(Guid relocationId, string category)
        {
            throw new NotImplementedException();
        }
    }
}