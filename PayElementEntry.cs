public class PayElementEntry
{
    public string BODId { get; set; }
    public string RecordNumber { get; set; }
    public string PersonId { get; set; } // Employee Number
    public string EmployeeId { get; set; }
    public string PayServId { get; set; }
    public string EmployeeName { get; set; }
    public string Event { get; set; }
    public string Action { get; set; }
    public string Category { get; set; }
    public string FieldLabel { get; set; }
    public string Value { get; set; }
    public string RowId { get; set; }
    public DateTime EndDate { get; set; }
    public string PayElementType { get; set; }

    // Pay Element specific
    public string PayElementId { get; set; }
    public string PaySpaceCompCode { get; set; }
    public string UnitType { get; set; }
    public decimal? NumberOfUnits { get; set; }
    public decimal? Amount { get; set; }

    // Personal Data Specific

    public string EmployeeFirstName { get; set; }
    public string EmployeeLastName { get; set; }
    public string EmployeeNumber { get; set; }
    public DateTime? BirthDate { get; set; }
    public string Title { get; set; }
    public string Gender { get; set; }
    public string Language { get; set; }
    public string CitizenshipCountry { get; set; }
    public string? EmailAddress { get; set; }

    //Employment Status
    public string TerminationReason { get; set; }
    public DateTime? TerminationDate { get; set; }

    //Position Data
    public string PositionTitle { get; set; }
    public string CostCenter { get; set; }

    // For mapping custom fields for other categories
    public Dictionary<string, string> CustomFields { get; set; } = new();

    // For date-driven categories
    public DateTime StartDate { get; set; }
}
