using System;

namespace ExpenseTracking.Dal.Entities.Dtos
{

    public class ExpenseRequestDto
    {
        //public Guid? ExpenseId { get; set; }

        public string ExpenseTypeName { get; set; }

        public string ExpenseTitle { get; set; }

        public decimal? ExpenseAmount { get; set; }

        public string Notes { get; set; }

        public DateTimeOffset? SentDate { get; set; }

        public DateTimeOffset? ReceivedDate { get; set; }

        public string ReceiptNumber { get; set; }

        public string TrackingNumber { get; set; }

        public DateTimeOffset? RecordedDate { get; set; }

        public bool? IsFileAttached { get; set; }

        public string AzureBlobId { get; set; }

        public bool? IsDeleted { get; set; }

        public Guid? ContactLogId { get; set; }

        public string ParcelTCAD { get; set; }

        public Guid? RelocationCaseId { get; set; }


        // validation
        public bool ValidateRequiredAmount() => this.ExpenseAmount.HasValue && this.ExpenseAmount > 0 && this.ExpenseAmount < 99999;

        public bool ValidateParcelTCAD() => !string.IsNullOrEmpty(this.ParcelTCAD);

        public bool ValidateExpenseTypeName() => !string.IsNullOrEmpty(this.ExpenseTypeName);

        public bool ValidateRecordedDateAndSentDate()
        {
            if (RecordedDate.HasValue && SentDate.HasValue && RecordedDate < SentDate)
            {
                return false;
            }
            return true;
        }
    }
}