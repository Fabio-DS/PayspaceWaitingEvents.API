using System;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using Microsoft.Extensions.Logging;

namespace PaySpaceWaitingEvents.API.Services
{
    public class WaitingEventsService : IWaitingEventsService
    {
        private readonly ILogger<WaitingEventsService> _logger;

        public WaitingEventsService(ILogger<WaitingEventsService> logger)
        {
            _logger = logger;
        }

        public async Task<(List<PayElementEntry>, string)> ProcessWaitingEventsFile(Stream stream)
        {
            var entries = new List<PayElementEntry>();
            string logicalIdPrefix = string.Empty;

            using var workbook = new XSSFWorkbook(stream);
            var sheet = workbook.GetSheetAt(0);
            var headerRow = sheet.GetRow(0);
            var columnMap = new Dictionary<string, int>();

            for (int i = 0; i < headerRow.LastCellNum; i++)
            {
                var cell = headerRow.GetCell(i);
                if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString()))
                    columnMap[cell.ToString().Trim()] = i;
            }

            for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null) continue;

                var currentLogicalId = row.GetCell(columnMap.GetValueOrDefault("Logical ID"))?.ToString();
                if (string.IsNullOrEmpty(logicalIdPrefix) && !string.IsNullOrEmpty(currentLogicalId))
                {
                    var match = Regex.Match(currentLogicalId, @"^(.+)-\d+$");
                    logicalIdPrefix = match.Success ? match.Groups[1].Value : currentLogicalId;
                }

                var eventType = row.GetCell(columnMap.GetValueOrDefault("Event"))?.ToString();
                var categoryName = row.GetCell(columnMap.GetValueOrDefault("Category / Form Name"))?.ToString();
                var fieldLabel = row.GetCell(columnMap.GetValueOrDefault("Field Label"))?.ToString();
                var value = row.GetCell(columnMap.GetValueOrDefault("Value"))?.ToString();

                if (string.IsNullOrWhiteSpace(eventType) || string.IsNullOrWhiteSpace(categoryName) || string.IsNullOrWhiteSpace(fieldLabel))
                    continue;

                var entry = new PayElementEntry
                {
                    RowId = currentLogicalId,
                    RecordNumber = row.GetCell(columnMap.GetValueOrDefault("Record Number"))?.ToString(),
                    PersonId = row.GetCell(columnMap.GetValueOrDefault("Person ID"))?.ToString(),
                    EmployeeId = row.GetCell(columnMap.GetValueOrDefault("Person ID"))?.ToString(),
                    EmployeeName = row.GetCell(columnMap.GetValueOrDefault("Employee Name"))?.ToString(),
                    Event = eventType,
                    Action = row.GetCell(columnMap.GetValueOrDefault("Action"))?.ToString(),
                    Category = categoryName
                };

                switch (eventType)
                {
                    case "Hiring":
                        if (categoryName == "Deployment" && fieldLabel == "Position Title")
                        {
                            entry.PositionTitle = value;
                        }
                        else if (categoryName == "Cost Assignment" && fieldLabel == "Cost Center Code")
                        {
                            entry.CostCenter = value;
                            if (DateTime.TryParse(row.GetCell(columnMap.GetValueOrDefault("Start Date"))?.ToString(), out var startDate))
                                entry.StartDate = startDate;
                        }
                        else if (categoryName == "Personal Data")
                        {
                            if (fieldLabel == "Given Name") entry.EmployeeFirstName = value;
                            else if (fieldLabel == "Family Name") entry.EmployeeLastName = value;
                            else if (fieldLabel == "Preferred Salutation") entry.Title = value;
                            else if (fieldLabel == "Birth Date" && DateTime.TryParse(value, out var birth)) entry.BirthDate = birth;
                            else if (fieldLabel == "Gender Code") entry.Gender = value;
                            else if (fieldLabel == "Primary Language Code") entry.Language = ConvertLanguageCode(value);
                            else if (fieldLabel == "Citizenship Country Code") entry.CitizenshipCountry = ConvertCountryCode(value);
                        }
                        else if (categoryName == "Pay Element")
                        {
                            entry.PayElementId = fieldLabel;
                            if (decimal.TryParse(value, out var amount))
                                entry.Amount = amount;
                        }
                        break;

                    case "Termination":
                        if (categoryName == "Deployment")
                        {
                            if (fieldLabel == "Termination Reason") entry.TerminationReason = value;
                            else if (fieldLabel == "Termination Date" && DateTime.TryParse(value, out var termDate))
                                entry.TerminationDate = termDate;
                        }
                        break;

                    case "Data Change":
                        if (categoryName == "Deployment" && fieldLabel == "Position Title")
                        {
                            entry.PositionTitle = value;
                        }
                        else if (categoryName == "Cost Assignment" && fieldLabel == "Cost Center Code")
                        {
                            entry.CostCenter = value;
                            if (DateTime.TryParse(row.GetCell(columnMap.GetValueOrDefault("Start Date"))?.ToString(), out var csStartDate))
                                entry.StartDate = csStartDate;
                        }
                        else if (categoryName == "Personal Data")
                        {
                            if (fieldLabel == "Given Name") entry.EmployeeFirstName = value;
                            else if (fieldLabel == "Family Name") entry.EmployeeLastName = value;
                            else if (fieldLabel == "Preferred Salutation") entry.Title = value;
                            else if (fieldLabel == "Birth Date" && DateTime.TryParse(value, out var birth)) entry.BirthDate = birth;
                            else if (fieldLabel == "Gender Code") entry.Gender = value;
                            else if (fieldLabel == "Primary Language Code") entry.Language = ConvertLanguageCode(value);
                            else if (fieldLabel == "Citizenship Country Code") entry.CitizenshipCountry = ConvertCountryCode(value);
                        }
                        break;

                    case "Pay Element":
                        if (categoryName == "Pay Element")
                        {
                            switch (fieldLabel)
                            {
                                case "Pay Element":
                                    entry.PayElementId = value;
                                    break;

                                case "Pay Element Type":
                                    entry.PayElementType = value;
                                    break;

                                case "Currency Code":
                                    entry.PaySpaceCompCode = value;
                                    break;

                                case "Unit Type":
                                    entry.UnitType = value;
                                    break;

                                case "Number of Units":
                                case "Units":
                                    if (decimal.TryParse(value, out var units))
                                        entry.NumberOfUnits = units;
                                    break;

                                case "Amount":
                                    if (decimal.TryParse(value, out var amt))
                                        entry.Amount = amt;
                                    break;

                                default:
                                    break;
                            }
                        }
                        break;

                }

                // Only include entries that were actually populated
                if (!string.IsNullOrWhiteSpace(entry.PositionTitle) ||
                    !string.IsNullOrWhiteSpace(entry.CostCenter) ||
                    !string.IsNullOrWhiteSpace(entry.PayElementId) ||
                    !string.IsNullOrWhiteSpace(entry.EmployeeFirstName) ||
                    !string.IsNullOrWhiteSpace(entry.TerminationReason) ||
                    entry.TerminationDate.HasValue ||
                    entry.BirthDate.HasValue ||
                    entry.Amount > 0)
                {
                    entries.Add(entry);
                }
            }

            return (entries, logicalIdPrefix);
        }

        private string ConvertLanguageCode(string code)
        {
            return code.ToUpperInvariant() switch
            {
                "EN" => "English",
                "AF" => "Afrikaans",
                "PT" => "Portuguese",
                _ => code //will do the rest later  
            };
        }

        private string ConvertCountryCode(string code)
        {
            return code.ToUpperInvariant() switch
            {
                "ZA" => "South Africa",
                "PT" => "Portugal",
                "US" => "United States",
                _ => code //will do the rest later  
            };
        }
    }
}