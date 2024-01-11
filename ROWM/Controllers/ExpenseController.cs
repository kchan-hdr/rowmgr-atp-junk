using ExpenseTracking.Dal;
using ExpenseTracking.Dal.Entities;
using ExpenseTracking.Dal.Entities.Dtos;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ROWM.Controllers
{

    [Produces("application/json")]
    [ApiController]
    public class ExpenseController : ControllerBase
    {
        const string _APP_NAME = "EXPENSE";

        readonly IExpenseTracking _expenseOp;
        readonly ILogger _logger;

        public ExpenseController(ILogger<ExpenseController> logger, IExpenseTracking expense)
        {
            _expenseOp = expense;
            _logger = logger;
        }

        [ProducesResponseType(StatusCodes.Status201Created)]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        [Produces("application/json")]
        [HttpPost("api/parcels/{pId}/expenses")]
        public async Task<ActionResult<Expense>> CreateExpense([FromBody] ExpenseRequestDto expenseRequestDto, string pId)
        {
            try
            {
                var addedExpense = await _expenseOp.AddExpense(expenseRequestDto);
                return Ok(new ExpenseJson(addedExpense));
            }
            catch (ArgumentException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError("An error prevented the expense from being added: {message}", ex.Message);
                return StatusCode(500, "An error occurred while creating the expense.");
            }
        }

        [ProducesResponseType(200)]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        [Produces("application/json")]
        [HttpPut("api/expenses/{expenseId}")]
        public async Task<ActionResult<Expense>> UpdateExpense(Guid expenseId, [FromBody] ExpenseRequestDto expenseRequestDto)
        {
            try
            {
                var updatededExpense = await _expenseOp.UpdateExpense(expenseId, expenseRequestDto);
                return Ok(new ExpenseJson(updatededExpense));
            }
            catch (ArgumentException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError("An error prevented the expense from being updated: {message}", ex.Message);
                return StatusCode(500, "An error occurred while updating the expense.");
            }
        }

        [ProducesResponseType(200)]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        [Produces("application/json")]
        [HttpDelete("api/expenses/{expenseId}")]
        public async Task<ActionResult<Expense>> DeleteExpense(Guid expenseId)
        {
            try
            {
                var deletedExpense = await _expenseOp.DeleteExpense(expenseId);
                return Ok(new ExpenseJson(deletedExpense));
            }
            catch (ArgumentException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError("An error prevented the expense from being deleted: {message}", ex.Message);
                return StatusCode(500, "An error occurred while deleting the expense.");
            }
        }

        [ProducesResponseType(200)]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        [Produces("application/json")]
        [HttpGet("api/expenses/{expenseId}")]
        public async Task<ActionResult<Expense>> GetExpense(Guid expenseId)
        {
            try
            {
                var myExpense = await _expenseOp.GetExpense(expenseId);
                return Ok(new ExpenseJson(myExpense));
            }
            catch (ArgumentException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError("An error prevented the expense from being fetched: {message}", ex.Message);
                return StatusCode(500, "An error occurred while fetching the expense.");
            }
        }

        [ProducesResponseType(200)]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        [Produces("application/json")]
        [HttpGet("api/parcels/{pId}/expenses")]
        public async Task<ActionResult<IEnumerable<ExpenseTypeDto>>> GetParcelExpense(string pId)
        {
            try
            {
                var expenseTypes = await _expenseOp.GetExpenses(pId, "Acquisition");
                return Ok(expenseTypes);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError("An error prevented the expense from being fetched: {message}", ex.Message);
                return StatusCode(500, "An error occurred while creating the expense.");
            }
        }

        [ProducesResponseType(200)]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        [Produces("application/json")]
        [HttpGet("api/relocation/{reloId}/expenses")]
        public async Task<ActionResult<IEnumerable<ExpenseTypeDto>>> GetRelocationExpense(Guid reloId)
        {
            try
            {
                var expenseTypes = await _expenseOp.GetExpenses(reloId, "Relocation");
                return Ok(expenseTypes);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError("An error prevented the expense from being fetched: {message}", ex.Message);
                return StatusCode(500, "An error occurred while creating the expense.");
            }
        }

        [ProducesResponseType(200)]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        [Produces("application/json")]
        [HttpGet("api/parcels/{pId}/expenses/total")]
        public async Task<ActionResult<double>> CalculateExpenseByParcel(string pId)
        {
            try
            {
                var totalExpense = await _expenseOp.CalculateExpenseByParcel(pId);
                return Ok(totalExpense);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError("An error prevented the expense info from being fetched: {message}", ex.Message);
                return StatusCode(500, "An error occurred while fetching the number.");
            }
        }

        [ProducesResponseType(200)]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        [Produces("application/json")]
        [HttpGet("api/parcels/{pId}/expenses/relocationTotal")]
        public async Task<ActionResult<double>> CalculateReloExpenseByParcel(string pId)
        {
            try
            {
                var totalExpense = await _expenseOp.CalculateReloExpenseByParcel(pId);
                return Ok(totalExpense);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError("An error prevented the expense info from being fetched: {message}", ex.Message);
                return StatusCode(500, "An error occurred while fetching the number.");
            }
        }

        [ProducesResponseType(200)]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        [Produces("application/json")]
        [HttpGet("api/relocation/{reloId}/expenses/total")]
        public async Task<ActionResult<double>> CalculateExpenseByRelocationCase(Guid reloId)
        {
            try
            {
                var totalExpense = await _expenseOp.CalculateExpenseByRelocationCase(reloId);
                return Ok(totalExpense);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError("An error prevented the expense info from being fetched: {message}", ex.Message);
                return StatusCode(500, "An error occurred while fetching the number.");
            }
        }
    }


    #region dto

    public class ExpenseJson
    {
        public Guid ExpenseId { get; set; }

        public string ExpenseTypeName { get; set; }

        public string ExpenseTitle { get; set; }

        public decimal ExpenseAmount { get; set; }

        public string Notes { get; set; }

        public DateTimeOffset? SentDate { get; set; }

        public DateTimeOffset? ReceivedDate { get; set; }

        public string ReceiptNumber { get; set; }

        public string TrackingNumber { get; set; }

        public DateTimeOffset? RecordedDate { get; set; }

        public bool? IsFileAttached { get; set; }

        public string AzureBlobId { get; set; }

        public DateTimeOffset Created { get; set; }

        public DateTimeOffset? LastModified { get; set; }

        public string ModifiedBy { get; set; }

        public bool IsDeleted { get; set; } = false;

        public Guid? ContactLogId { get; set; }

        public string ParcelTCAD { get; set; }

        public Guid? RelocationCaseId { get; set; }

        public ExpenseJson(Expense e)
        {
            ExpenseId = e.ExpenseId;
            AzureBlobId = e.AzureBlobId;
            ContactLogId = e.ContactLogId;
            ExpenseAmount = e.ExpenseAmount;
            ExpenseTitle = e.ExpenseTitle;
            ExpenseTypeName = e.ExpenseTypeName;
            IsDeleted = e.IsDeleted;
            IsFileAttached = e.IsFileAttached;
            Notes = e.Notes;
            ParcelTCAD = e.ParcelTCAD;
            ReceiptNumber = e.ReceiptNumber;
            ReceivedDate = e.ReceivedDate;
            RecordedDate = e.RecordedDate;
            RelocationCaseId = e.RelocationCaseId;
            SentDate = e.SentDate;
            TrackingNumber = e.TrackingNumber;
        }
    }
    #endregion
}