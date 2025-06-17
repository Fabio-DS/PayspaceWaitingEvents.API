using System;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using Microsoft.Extensions.Logging;
using System.Globalization;

namespace PaySpaceWaitingEvents.API.Services
{
    public class WaitingEventsService : IWaitingEventsService
    {
        private readonly ILogger<WaitingEventsService> _logger;

        public WaitingEventsService(ILogger<WaitingEventsService> logger)
        {
            _logger = logger;
        }

        private string GetStringValue(IRow row, Dictionary<string, int> columnMap, string columnName)
        {
            if (columnMap.TryGetValue(columnName, out int colIndex) && colIndex >= 0)
            {
                var cell = row.GetCell(colIndex);
                return cell?.ToString()?.Trim();
            }
            // _logger.LogDebug($"Column '{columnName}' not found in header map or index was invalid for row {row.RowNum + 1}.");
            return null;
        }

        private DateTime? GetDateTimeValue(string cellValue, string context)
        {
            if (DateTime.TryParse(cellValue, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var result))
            {
                return result;
            }
            if (!string.IsNullOrWhiteSpace(cellValue))
            {
                _logger.LogWarning($"Could not parse DateTime from '{cellValue}' ({context}).");
            }
            return null;
        }

        private DateTime? GetDateTimeValueFromColumn(IRow row, Dictionary<string, int> columnMap, string columnName)
        {
            string cellValue = GetStringValue(row, columnMap, columnName);
            return GetDateTimeValue(cellValue, $"dedicated column '{columnName}', Row {row.RowNum + 1}");
        }

        private decimal? GetDecimalValue(string cellValue, string context)
        {
            if (decimal.TryParse(cellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
            {
                return result;
            }
            if (!string.IsNullOrWhiteSpace(cellValue))
            {
                _logger.LogWarning($"Could not parse decimal from '{cellValue}' ({context}).");
            }
            return null;
        }

        public async Task<(List<PayElementEntry>, string)> ProcessWaitingEventsFile(Stream stream)
        {
            var aggregatedEntries = new Dictionary<string, PayElementEntry>();
            string logicalIdPrefix = string.Empty;

            using var workbook = new XSSFWorkbook(stream);
            var sheet = workbook.GetSheetAt(0);
            if (sheet == null || sheet.LastRowNum < 0)
            {
                _logger.LogWarning("Excel sheet is empty or contains no data rows.");
                return (new List<PayElementEntry>(), logicalIdPrefix);
            }

            var headerRow = sheet.GetRow(0);
            if (headerRow == null)
            {
                _logger.LogWarning("Excel sheet is missing a header row.");
                return (new List<PayElementEntry>(), logicalIdPrefix);
            }

            var columnMap = new Dictionary<string, int>();
            _logger.LogInformation("Reading Excel headers:");
            for (int i = 0; i < headerRow.LastCellNum; i++)
            {
                var cell = headerRow.GetCell(i);
                if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString()))
                {
                    string headerName = cell.ToString().Trim();
                    columnMap[headerName] = i;
                    _logger.LogInformation($"Header: '{headerName}' found at index {i}");
                }
            }

            // Check if critical columns for aggregation were found
            bool personIdColFound = columnMap.ContainsKey("Person ID");
            bool recordNumColFound = columnMap.ContainsKey("Record Number");
            _logger.LogInformation($"Column 'Person ID' found in header: {personIdColFound}");
            _logger.LogInformation($"Column 'Record Number' found in header: {recordNumColFound}");

            if (!personIdColFound || !recordNumColFound)
            {
                _logger.LogError("CRITICAL: 'Person ID' or 'Record Number' column not found in Excel header. Aggregation will likely fail. Please check Excel column names.");
                // Optionally, return early or handle this error more gracefully
            }


            for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null)
                {
                    _logger.LogWarning($"Row {rowIndex + 1} is null, skipping.");
                    continue;
                }

                var currentLogicalId = GetStringValue(row, columnMap, "Logical ID");
                if (string.IsNullOrEmpty(logicalIdPrefix) && !string.IsNullOrEmpty(currentLogicalId))
                {
                    var match = Regex.Match(currentLogicalId, @"^(.+)-\d+$");
                    logicalIdPrefix = match.Success ? match.Groups[1].Value : currentLogicalId;
                }

                string personIdStr = GetStringValue(row, columnMap, "Person ID");
                string recordNumberStr = GetStringValue(row, columnMap, "Record Number");

                // Detailed logging for first 20 data rows (rowIndex 1 to 20)
                if (rowIndex <= 20)
                {
                    _logger.LogInformation($"Row {rowIndex + 1}: Raw PersonID='{personIdStr}', Raw RecordNumber='{recordNumberStr}'");
                }

                if (string.IsNullOrWhiteSpace(personIdStr) || string.IsNullOrWhiteSpace(recordNumberStr))
                {
                    _logger.LogWarning($"Skipping Excel row {rowIndex + 1} due to missing Person ID ('{personIdStr}') or Record Number ('{recordNumberStr}'). This row will not be processed.");
                    continue;
                }

                string entryKey = $"{personIdStr}_{recordNumberStr}";
                if (rowIndex <= 20)
                {
                    _logger.LogInformation($"Row {rowIndex + 1}: Generated entryKey='{entryKey}'");
                }


                var eventType = GetStringValue(row, columnMap, "Event");
                var action = GetStringValue(row, columnMap, "Action");
                var categoryName = GetStringValue(row, columnMap, "Category / Form Name");
                var fieldLabel = GetStringValue(row, columnMap, "Field Label");
                var value = GetStringValue(row, columnMap, "Value");

                DateTime? startDateFromDedicatedColumn = GetDateTimeValueFromColumn(row, columnMap, "Start Date");
                DateTime? endDateFromDedicatedColumn = GetDateTimeValueFromColumn(row, columnMap, "End Date");

                PayElementEntry entry;
                if (!aggregatedEntries.TryGetValue(entryKey, out entry))
                {
                    entry = new PayElementEntry
                    {
                        PersonId = personIdStr,
                        RecordNumber = recordNumberStr,
                        EmployeeId = personIdStr,
                        EmployeeName = GetStringValue(row, columnMap, "Employee Name"),
                        RowId = currentLogicalId,
                        Event = eventType,
                        Action = action,
                        Category = categoryName
                    };
                    aggregatedEntries[entryKey] = entry;
                }

                if (!string.IsNullOrWhiteSpace(eventType)) entry.Event = eventType;
                if (!string.IsNullOrWhiteSpace(action)) entry.Action = action;
                if (!string.IsNullOrWhiteSpace(categoryName)) entry.Category = categoryName;
                if (!string.IsNullOrWhiteSpace(currentLogicalId)) entry.RowId = currentLogicalId;


                if (string.IsNullOrWhiteSpace(entry.Event) || string.IsNullOrWhiteSpace(entry.Category) || string.IsNullOrWhiteSpace(fieldLabel) && string.IsNullOrWhiteSpace(value)) // Allow rows with no field/value if other data present
                {
                    _logger.LogInformation($"Row {rowIndex + 1} (Key: {entryKey}) may have limited data for field-specific parsing: Event='{entry.Event}', Category='{entry.Category}', FieldLabel='{fieldLabel}', Value='{value}'.");
                }

                switch (entry.Event)
                {
                    case "Hiring":
                    case "Data Change":
                        if (categoryName == "Deployment")
                        {
                            if (fieldLabel?.Equals("Position Title", StringComparison.OrdinalIgnoreCase) == true) entry.PositionTitle = value;
                            else if (fieldLabel?.Equals("Termination Reason", StringComparison.OrdinalIgnoreCase) == true) entry.TerminationReason = value;
                            else if (fieldLabel?.Equals("Termination Date", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                entry.TerminationDate = GetDateTimeValue(value, $"Value for {entry.Event}/Deployment/Termination Date, Row {rowIndex + 1}");
                            }
                        }
                        else if (categoryName == "Cost Assignment")
                        {
                            if (fieldLabel?.Equals("Cost Center Code", StringComparison.OrdinalIgnoreCase) == true) entry.CostCenter = value;
                            if (startDateFromDedicatedColumn.HasValue && entry.StartDate == default) entry.StartDate = startDateFromDedicatedColumn.Value;
                            else if (fieldLabel?.Equals("Start Date", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                entry.StartDate = GetDateTimeValue(value, $"Value for {entry.Event}/Cost Assignment/Start Date, Row {rowIndex + 1}") ?? entry.StartDate;
                            }
                        }
                        else if (categoryName == "Personal Data")
                        {
                            if (fieldLabel?.Equals("Given Name", StringComparison.OrdinalIgnoreCase) == true) entry.EmployeeFirstName = value;
                            else if (fieldLabel?.Equals("Family Name", StringComparison.OrdinalIgnoreCase) == true) entry.EmployeeLastName = value;
                            // ... (other personal data fields with case-insensitive compare)
                            else if (fieldLabel?.Equals("Preferred Salutation", StringComparison.OrdinalIgnoreCase) == true) entry.Title = value;
                            else if (fieldLabel?.Equals("Birth Date", StringComparison.OrdinalIgnoreCase) == true)
                                entry.BirthDate = GetDateTimeValue(value, $"Value for {entry.Event}/Personal Data/Birth Date, Row {rowIndex + 1}");
                            else if (fieldLabel?.Equals("Gender Code", StringComparison.OrdinalIgnoreCase) == true) entry.Gender = value;
                            else if (fieldLabel?.Equals("Primary Language Code", StringComparison.OrdinalIgnoreCase) == true) entry.Language = ConvertLanguageCode(value);
                            else if (fieldLabel?.Equals("Citizenship Country Code", StringComparison.OrdinalIgnoreCase) == true) entry.CitizenshipCountry = ConvertCountryCode(value);
                        }
                        else if (categoryName?.Contains("Pay Element", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            ProcessPayElementFields(entry, fieldLabel, value, entry.Event, categoryName, startDateFromDedicatedColumn, endDateFromDedicatedColumn, rowIndex + 1);
                        }
                           else if (categoryName == "Communication") 
                        { 
                            if (fieldLabel?.Equals("Email Business Uri", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                entry.EmailAddress = value;
                            }
                        }
                        break;

                    case "Termination":
                        if (categoryName == "Deployment")
                        {
                            if (fieldLabel?.Equals("Termination Reason", StringComparison.OrdinalIgnoreCase) == true) entry.TerminationReason = value;
                            else if (fieldLabel?.Equals("Termination Date", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                entry.TerminationDate = GetDateTimeValue(value, $"Value for Termination/Deployment/Termination Date, Row {rowIndex + 1}");
                            }
                        }
                        break;

                    case "Pay Element":
                        if (categoryName?.Contains("Pay Element", StringComparison.OrdinalIgnoreCase) == true)
                        {
                            _logger.LogInformation($"DEBUG: Row {rowIndex + 1} (Key {entryKey}, Event {entry.Event}) calling ProcessPayElementFields. FieldLabel='{fieldLabel}', Value='{value}'");
                            ProcessPayElementFields(entry, fieldLabel, value, entry.Event, categoryName, startDateFromDedicatedColumn, endDateFromDedicatedColumn, rowIndex + 1);
                        }
                        break;
                }
            }

            _logger.LogInformation($"Aggregation complete. Total unique PersonId/RecordNumber keys found: {aggregatedEntries.Count}");
            var finalEntries = new List<PayElementEntry>();
            foreach (var kvp in aggregatedEntries)
            {
                if (IsPopulated(kvp.Value)) { finalEntries.Add(kvp.Value); }
            }
            _logger.LogInformation($"Filtering complete. Extracted {finalEntries.Count} populated entries.");
            return (finalEntries, logicalIdPrefix);
        }

        private void ProcessPayElementFields(PayElementEntry entry, string fieldLabel, string value, string eventType, string categoryName, DateTime? dedicatedStartDate, DateTime? dedicatedEndDate, int excelRowNum)
        {
            string normalizedFieldLabel = fieldLabel?.Trim().ToLowerInvariant();

            switch (normalizedFieldLabel)
            {
                case "pay element":
                case "pay element id":
                    entry.PayElementId = value; break;
                case "pay element type": entry.PayElementType = value; break;
                case "currency code": entry.PaySpaceCompCode = value; break;
                case "unit type": entry.UnitType = value?.Trim(); break;
                case "number of units":
                case "units":
                    var units = GetDecimalValue(value, $"Units for {entry.PersonId}_{entry.RecordNumber}, Row {excelRowNum}");
                    if (units.HasValue) entry.NumberOfUnits = units.Value;
                    break;
                case "amount":
                    var amt = GetDecimalValue(value, $"Amount for {entry.PersonId}_{entry.RecordNumber}, Row {excelRowNum}");
                    if (amt.HasValue) entry.Amount = amt.Value;
                    break;
                case "start date":
                    var payElStartDate = GetDateTimeValue(value, $"Start Date (Field Label) for {entry.PersonId}_{entry.RecordNumber}, Row {excelRowNum}");
                    if (payElStartDate.HasValue) entry.StartDate = payElStartDate.Value;
                    else if (dedicatedStartDate.HasValue && entry.StartDate == default) entry.StartDate = dedicatedStartDate.Value; 
                    break;
                case "end date":
                    var payElEndDate = GetDateTimeValue(value, $"End Date (Field Label) for {entry.PersonId}_{entry.RecordNumber}, Row {excelRowNum}");
                    if (payElEndDate.HasValue) entry.EndDate = payElEndDate.Value;
                    else if (dedicatedEndDate.HasValue && entry.EndDate == default) entry.EndDate = dedicatedEndDate.Value;
                    break;
                default:
                    if (!string.IsNullOrWhiteSpace(fieldLabel) &&
                        !normalizedFieldLabel.Contains("position") && 
                        !normalizedFieldLabel.Contains("cost center")) 
                    {
                        _logger.LogWarning($"Unhandled Field Label within PayElement processing: '{fieldLabel}' (Normalized: '{normalizedFieldLabel}') for Event: '{eventType}', Category: '{categoryName}', Entry: {entry.PersonId}_{entry.RecordNumber}, Row: {excelRowNum}. Value: '{value}'");
                    }
                    break;
            }

            if (entry.StartDate == default && dedicatedStartDate.HasValue) entry.StartDate = dedicatedStartDate.Value;
            if (entry.EndDate == default && dedicatedEndDate.HasValue) entry.EndDate = dedicatedEndDate.Value;
        }

        private bool IsPopulated(PayElementEntry e)
        {
            bool hasCoreEmploymentData = !string.IsNullOrWhiteSpace(e.PositionTitle) || !string.IsNullOrWhiteSpace(e.CostCenter);
            bool hasTerminationData = !string.IsNullOrWhiteSpace(e.TerminationReason) || e.TerminationDate.HasValue;
            bool hasPersonalData = !string.IsNullOrWhiteSpace(e.EmployeeFirstName) || !string.IsNullOrWhiteSpace(e.EmployeeLastName) || e.BirthDate.HasValue ||
                                   !string.IsNullOrWhiteSpace(e.Title) || !string.IsNullOrWhiteSpace(e.Gender) || !string.IsNullOrWhiteSpace(e.Language) || !string.IsNullOrWhiteSpace(e.CitizenshipCountry);
            bool hasPayElementData = !string.IsNullOrWhiteSpace(e.PayElementId) && (e.Amount.HasValue || e.NumberOfUnits.HasValue);
            bool hasSignificantDatedEvent = (e.StartDate != default || e.EndDate != default) &&
                                            (!string.IsNullOrWhiteSpace(e.PayElementId) || !string.IsNullOrWhiteSpace(e.CostCenter) || !string.IsNullOrWhiteSpace(e.PositionTitle));

            if (hasCoreEmploymentData || hasTerminationData || hasPersonalData || hasPayElementData || hasSignificantDatedEvent)
            {
                return true;
            }

            _logger.LogWarning(
                $"Filtering out aggregated entry (Key: {e.PersonId}_{e.RecordNumber}, Event: {e.Event}, Category: {e.Category}, RowId: {e.RowId}). Criteria check: " +
                $"hasCoreEmploymentData={hasCoreEmploymentData} (PosTitle='{e.PositionTitle}', CostCenter='{e.CostCenter}'); " +
                $"hasTerminationData={hasTerminationData} (TermReason='{e.TerminationReason}', TermDate='{e.TerminationDate}'); " +
                $"hasPersonalData={hasPersonalData} (FirstName='{e.EmployeeFirstName}', LastName='{e.EmployeeLastName}', BirthDate='{e.BirthDate}'); " +
                $"hasPayElementData={hasPayElementData} (PayElemId='{e.PayElementId}', Amount='{e.Amount}', Units='{e.NumberOfUnits}', UnitType='{e.UnitType}'); " +
                $"hasSignificantDatedEvent={hasSignificantDatedEvent} (StartDate='{e.StartDate}', EndDate='{e.EndDate}')"
            );
            return false;
        }

        private string ConvertLanguageCode(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return code;
            return code.ToUpperInvariant() switch { "EN" => "English", "AF" => "Afrikaans", "PT" => "Portuguese", _ => code };
        }

        private string ConvertCountryCode(string code)
        {
            if (string.IsNullOrWhiteSpace(code)) return code;
            return code.ToUpperInvariant() switch { "ZA" => "South Africa", "PT" => "Portugal", "US" => "United States", _ => code };
        }
    }
}
