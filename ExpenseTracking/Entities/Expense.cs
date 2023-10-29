using ExpenseTracking.Dal.Entities.Dtos;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ExpenseTracking.Dal.Entities
{

    [Table("Expense", Schema = "ROWM")]
    public class Expense
    {
        [Key, DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public Guid ExpenseId { get; set; }

        [StringLength(50), Required, ForeignKey(nameof(ExpenseType))]
        public string ExpenseTypeName { get; set; }
        public virtual ExpenseType ExpenseType { get; set; }

        [StringLength(100)]
        public string ExpenseTitle { get; set; }

        public double ExpenseAmount { get; set; }

        public string Notes { get; set; }

        public DateTimeOffset? SentDate { get; set; }

        public DateTimeOffset? ReceivedDate { get; set; }

        [StringLength(100)]
        public string ReceiptNumber { get; set; }

        [StringLength(100)]
        public string TrackingNumber { get; set; }

        public DateTimeOffset? RecordedDate { get; set; }

        public bool? IsFileAttached { get; set; }

        public string AzureBlobId { get; set; }

        public DateTimeOffset Created { get; set; }

        public DateTimeOffset? LastModified { get; set; }

        [StringLength(50)]
        public string ModifiedBy { get; set; }

        public bool IsDeleted { get; set; } = false;

        public Guid? ContactLogId { get; set; }

        [StringLength(100)]
        public string ParcelTCAD { get; set; }

        public Guid? RelocationCaseId { get; set; }

        public virtual ICollection<ExpenseActivity> ExpenseParentActivities { get; } = new HashSet<ExpenseActivity>();

        public virtual ICollection<ExpenseActivity> ExpenseChildActivities { get; } = new HashSet<ExpenseActivity>();

        public Expense() { }

        public Expense(ExpenseRequestDto expenseRequestDto)
        {
            ExpenseTypeName = expenseRequestDto.ExpenseTypeName;
            ExpenseTitle = expenseRequestDto.ExpenseTitle;
            ExpenseAmount = expenseRequestDto.ExpenseAmount ?? throw new ArgumentNullException(nameof(expenseRequestDto.ExpenseAmount));
            Notes = expenseRequestDto.Notes;
            SentDate = expenseRequestDto.SentDate;
            ReceivedDate = expenseRequestDto.ReceivedDate;
            ReceiptNumber = expenseRequestDto.ReceiptNumber;
            TrackingNumber = expenseRequestDto.TrackingNumber;
            RecordedDate = expenseRequestDto.RecordedDate;
            IsFileAttached = expenseRequestDto.IsFileAttached;
            AzureBlobId = expenseRequestDto.AzureBlobId;
            Created = DateTimeOffset.UtcNow;
            LastModified = DateTimeOffset.UtcNow;
            ModifiedBy = "ROWM ATP Expense Tracking";
            IsDeleted = false;
            ContactLogId = expenseRequestDto.ContactLogId;
            ParcelTCAD = expenseRequestDto.ParcelTCAD;
            RelocationCaseId = expenseRequestDto.RelocationCaseId;
        }

        public Expense Update(ExpenseRequestDto expenseRequestDto)
        {
            ExpenseTypeName = expenseRequestDto.ExpenseTypeName;
            ExpenseTitle = expenseRequestDto.ExpenseTitle ?? ExpenseTitle;
            ExpenseAmount = expenseRequestDto.ExpenseAmount ?? ExpenseAmount;
            Notes = expenseRequestDto.Notes ?? Notes;
            SentDate = expenseRequestDto.SentDate ?? SentDate;
            ReceivedDate = expenseRequestDto.ReceivedDate ?? ReceivedDate;
            ReceiptNumber = expenseRequestDto.ReceiptNumber ?? ReceiptNumber;
            TrackingNumber = expenseRequestDto.TrackingNumber ?? TrackingNumber;
            RecordedDate = expenseRequestDto.RecordedDate ?? RecordedDate;
            IsFileAttached = expenseRequestDto.IsFileAttached ?? IsFileAttached;
            AzureBlobId = expenseRequestDto.AzureBlobId ?? AzureBlobId;
            LastModified = DateTimeOffset.UtcNow;
            ModifiedBy = "ROWM ATP Expense Tracking Update";
            IsDeleted = expenseRequestDto.IsDeleted ?? IsDeleted;
            ContactLogId = expenseRequestDto.ContactLogId ?? ContactLogId;
            return this;
        }
    }
}