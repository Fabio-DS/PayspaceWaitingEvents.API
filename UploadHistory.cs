using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace PaySpaceWaitingEvents.API.Models
{
    public class UploadHistory
    {
        public int Id { get; set; }
        public DateTime UploadTimestamp { get; set; }
        public int CompanyId { get; set; }

        [MaxLength(100)]
        public string CompanyName { get; set; }

        [MaxLength(50)]
        public string CompanyCode { get; set; }

        [MaxLength(50)]
        public string Frequency { get; set; }

        [MaxLength(50)]
        public string Run { get; set; }

        public int TotalEntries { get; set; }
        public int SuccessfulEntries { get; set; }
        public int FailedEntries { get; set; }
        public bool Success { get; set; }

        [MaxLength(255)]
        public string FileName { get; set; }

        public virtual ICollection<UploadHistoryEntry> Entries { get; set; } = new List<UploadHistoryEntry>();
    }

    public class UploadHistoryEntry
    {
        public int Id { get; set; }
        public int UploadHistoryId { get; set; }

        [MaxLength(50)]
        public string EmployeeId { get; set; }

        [MaxLength(100)]
        public string EmployeeName { get; set; }

        [MaxLength(50)]
        public string PayElementId { get; set; }

        [MaxLength(50)]
        public string ComponentCode { get; set; }

        public bool Success { get; set; }

        [MaxLength(500)]
        public string ErrorMessage { get; set; }

        public decimal? Amount { get; set; }

        [MaxLength(50)]
        public string InputType { get; set; }

        public decimal? InputQuantity { get; set; }

        public virtual UploadHistory UploadHistory { get; set; }
    }
}