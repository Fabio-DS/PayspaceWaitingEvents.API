using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using NPOI.XSSF.UserModel;
using PaySpaceWaitingEvents.API.Data;
using PaySpaceWaitingEvents.API.Models;
using PaySpaceWaitingEvents.API.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace PaySpaceWaitingEvents.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class WaitingEventsController : ControllerBase
    {
        private readonly IWaitingEventsService _waitingEventsService;
        private readonly ILogger<WaitingEventsController> _logger;
        private readonly ApplicationDbContext _dbContext;
        private readonly IPaySpaceApiService _paySpaceApiService;
        private readonly IUploadHistoryService _uploadHistoryService;
        private static List<PayElementEntry> _lastProcessedEntries;
        private static List<PayElementEntry> _lastMappedEntries;

        public WaitingEventsController(
            IWaitingEventsService waitingEventsService,
            ApplicationDbContext dbContext,
            ILogger<WaitingEventsController> logger,
            IPaySpaceApiService paySpaceApiService,
            IUploadHistoryService uploadHistoryService)
        {
            _waitingEventsService = waitingEventsService;
            _dbContext = dbContext;
            _logger = logger;
            _paySpaceApiService = paySpaceApiService;
            _uploadHistoryService = uploadHistoryService;
        }

        [HttpPost("upload")]
        public async Task<IActionResult> UploadFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded");
            }

            try
            {
                _logger.LogInformation($"Starting file processing: {file.FileName}");

                using var stream = file.OpenReadStream();
                _logger.LogInformation("File stream opened");

                var (entries, logicalIdPrefix) = await _waitingEventsService.ProcessWaitingEventsFile(stream);
                _lastProcessedEntries = entries;
                _lastMappedEntries = null;

                var response = new
                {
                    Message = "File processed successfully",
                    TotalEntriesFound = entries.Count,
                    LogicalIdPrefix = logicalIdPrefix,
                    Summary = new
                    {
                        UniqueEmployees = entries.Select(e => e.EmployeeId).Distinct().Count(),
                        PayElements = entries.Select(e => e.PayElementId).Distinct().Count(),
                        // This summary needs care: Summing 'Amount' might be misleading if values are in NumberOfUnits
                        // For a simple summary, TotalAmount from Amount field is kept, but acknowledge its limitation.
                        TotalAmount = entries.Sum(e => e.Amount ?? 0)
                    }
                };

                _logger.LogInformation($"Successfully processed {entries.Count} entries");
                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing waiting events file");
                return StatusCode(500, new
                {
                    error = "Error processing file",
                    details = ex.Message,
                    fileName = file.FileName
                });
            }
        }

        [HttpGet("company-by-logical-id/{logicalIdPrefix}")]
        public async Task<IActionResult> GetCompanyByLogicalId(string logicalIdPrefix)
        {
            try
            {
                var legalEntity = await _dbContext.LegalEntities
                    .Where(e => e.LogicalIdPrefix == logicalIdPrefix && e.IsActive)
                    .FirstOrDefaultAsync();

                if (legalEntity == null)
                {
                    return NotFound(new { message = "No company found with the specified logical ID prefix" });
                }

                return Ok(new
                {
                    id = legalEntity.Id,
                    companyCode = legalEntity.CompanyCode,
                    companyName = legalEntity.CompanyName
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error finding company by logical ID prefix: {logicalIdPrefix}");
                return StatusCode(500, new { error = "Error retrieving company", details = ex.Message });
            }
        }

        [HttpGet("db-companies")]
        public async Task<IActionResult> GetDatabaseCompanies()
        {
            try
            {
                _logger.LogInformation("Getting companies from database");

                var companies = await _dbContext.LegalEntities
                    .Where(e => e.IsActive)
                    .Select(e => new
                    {
                        id = e.Id,
                        companyCode = e.CompanyCode,
                        companyName = e.CompanyName
                    })
                    .ToListAsync();

                return Ok(companies);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving companies from database");
                return StatusCode(500, new { error = "Error retrieving companies", details = ex.Message });
            }
        }

        [HttpGet("db-frequencies/{legalEntityId}")]
        public async Task<IActionResult> GetDatabaseFrequencies(int legalEntityId)
        {
            try
            {
                _logger.LogInformation($"Getting frequencies for legal entity ID {legalEntityId} from database");

                var frequencies = await _dbContext.PayElementMappings
                    .Where(m => m.LegalEntityId == legalEntityId && m.IsActive)
                    .Select(m => m.Frequency)
                    .Distinct()
                    .ToListAsync();

                _logger.LogInformation($"Found {frequencies.Count} frequencies in database for legal entity {legalEntityId}");
                return Ok(frequencies);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error retrieving frequencies from database for legal entity {legalEntityId}");
                return StatusCode(500, new { error = "Error retrieving frequencies", details = ex.Message });
            }
        }

        [HttpGet("companies")]
        public IActionResult GetCompanies()
        {
            try
            {
                var companies = _paySpaceApiService.GetAvailableCompanies();
                return Ok(companies);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving companies from PaySpace");
                return StatusCode(500, new { error = "Error retrieving companies from PaySpace", details = ex.Message });
            }
        }

        [HttpGet("frequencies/{companyId}")]
        public IActionResult GetFrequencies(int companyId)
        {
            try
            {
                _logger.LogInformation($"Getting frequencies for company ID: {companyId}");

                var frequencies = _paySpaceApiService.GetAvailableFrequencies(companyId);
                return Ok(frequencies);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error retrieving frequencies for company {companyId}");
                return StatusCode(500, new { error = "Error retrieving frequencies", details = ex.Message });
            }
        }

        [HttpGet("runs/{companyId}/{frequency}")]
        public IActionResult GetRuns(int companyId, string frequency)
        {
            try
            {
                _logger.LogInformation($"Getting runs for company ID: {companyId}, frequency: {frequency}");

                var runs = _paySpaceApiService.GetCompanyRuns(companyId, frequency);
                return Ok(runs);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error retrieving runs for company {companyId} and frequency {frequency}");
                return StatusCode(500, new { error = "Error retrieving runs", details = ex.Message });
            }
        }

        [HttpPost("map/{companyId}/{frequency}")]
        public async Task<IActionResult> MapEntries(int companyId, string frequency)
        {
            if (_lastProcessedEntries == null || !_lastProcessedEntries.Any())
            {
                return NotFound("No processed entries found. Please upload a file first.");
            }

            try
            {
                var mappings = await _dbContext.PayElementMappings
                    .Where(m => m.LegalEntityId == companyId && m.Frequency == frequency && m.IsActive)
                    .ToDictionaryAsync(m => m.PayElementId, m => m.ComponentCode);

                var mappedEntries = new List<PayElementEntry>();
                var unmappedElements = new HashSet<string>();

                foreach (var entry in _lastProcessedEntries)
                {
                    if (mappings.TryGetValue(entry.PayElementId, out string componentCode))
                    {
                        // Create a new PayElementEntry for the mapped list to avoid modifying _lastProcessedEntries
                        var mappedEntry = new PayElementEntry
                        {
                            BODId = entry.BODId,
                            RecordNumber = entry.RecordNumber,
                            PersonId = entry.PersonId,
                            EmployeeId = entry.EmployeeId,
                            PayServId = entry.PayServId,
                            EmployeeName = entry.EmployeeName,
                            Event = entry.Event,
                            Action = entry.Action,
                            Category = entry.Category,
                            FieldLabel = entry.FieldLabel,
                            Value = entry.Value, // This is the raw 'Value' from excel row, not the derived one for export
                            RowId = entry.RowId,
                            EndDate = entry.EndDate,
                            PayElementType = entry.PayElementType,
                            PayElementId = entry.PayElementId,
                            PaySpaceCompCode = componentCode, // Mapped Component Code
                            UnitType = entry.UnitType,
                            NumberOfUnits = entry.NumberOfUnits,
                            Amount = entry.Amount,
                            EmployeeFirstName = entry.EmployeeFirstName,
                            EmployeeLastName = entry.EmployeeLastName,
                            EmployeeNumber = entry.EmployeeNumber,
                            BirthDate = entry.BirthDate,
                            Title = entry.Title,
                            Gender = entry.Gender,
                            Language = entry.Language,
                            CitizenshipCountry = entry.CitizenshipCountry,
                            TerminationReason = entry.TerminationReason,
                            TerminationDate = entry.TerminationDate,
                            PositionTitle = entry.PositionTitle,
                            CostCenter = entry.CostCenter,
                            CustomFields = new Dictionary<string, string>(entry.CustomFields),
                            StartDate = entry.StartDate
                        };
                        mappedEntries.Add(mappedEntry);
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(entry.PayElementId)) // Only add if PayElementId was present
                        {
                            unmappedElements.Add(entry.PayElementId);
                        }
                    }
                }

                _lastMappedEntries = mappedEntries;

                if (unmappedElements.Any())
                {
                    return BadRequest(new
                    {
                        Message = "Mapping failed. There are unmapped pay elements.",
                        TotalOriginalEntries = _lastProcessedEntries.Count,
                        TotalMappedEntries = mappedEntries.Count, // Should be _lastProcessedEntries.Count - unmappedElements.Count
                        UnmappedPayElementsCount = unmappedElements.Count,
                        UnmappedPayElements = unmappedElements.ToList()
                    });
                }

                var response = new
                {
                    Message = "Entries mapped successfully",
                    TotalOriginalEntries = _lastProcessedEntries.Count,
                    TotalMappedEntries = mappedEntries.Count,
                    Summary = new
                    {
                        UniqueEmployees = mappedEntries.Select(e => e.EmployeeId).Distinct().Count(),
                        PayElements = mappedEntries.Select(e => e.PayElementId).Distinct().Count(),
                        TotalAmount = mappedEntries.Sum(e => e.Amount ?? 0), // Again, this sum is on the 'Amount' field
                    }
                };

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error mapping entries for company {companyId} and frequency {frequency}");
                return StatusCode(500, new { error = "Error mapping entries", details = ex.Message });
            }
        }

        [HttpGet("export-original")]
        public IActionResult ExportOriginalToExcel()
        {
            if (_lastProcessedEntries == null || !_lastProcessedEntries.Any())
            {
                return NotFound("No processed entries found. Please upload a file first.");
            }
            // Pass "OriginalEntries" to distinguish it, though the export logic is now unified
            return ExportToExcel(_lastProcessedEntries, "ProcessedEntries");
        }

        [HttpGet("export-mapped")]
        public IActionResult ExportMappedToExcel()
        {
            if (_lastMappedEntries == null || !_lastMappedEntries.Any())
            {
                return NotFound("No mapped entries found. Please map entries first.");
            }

            return ExportToExcel(_lastMappedEntries, "MappedEntries");
        }

        [HttpPost("upload-to-payspace/{companyId}")]
        public async Task<IActionResult> UploadToPayspace(int companyId, [FromQuery] string frequency, [FromQuery] string run)
        {
            if (_lastMappedEntries == null || !_lastMappedEntries.Any())
            {
                return NotFound("No mapped entries found. Please map entries first.");
            }

            try
            {
                _logger.LogInformation($"Starting upload to PaySpace for company ID {companyId}, frequency {frequency}, run {run}");
                _logger.LogInformation($"Uploading {_lastMappedEntries.Count} mapped entries");

                var entriesByCategory = _lastMappedEntries
                    .GroupBy(e =>
                    {
                        if (e.Event == "Hiring") return "Hiring";
                        if (e.Event == "Pay Element") return "PayElement";
                        if (e.Event == "Data Change" && e.Category == "Personal Data") return "Personal Data";
                        if (e.Event == "Data Change" && (e.Category == "Deployment" || e.Category == "Approver" || e.Category == "Cost Assignment")) return "Position Data";
                        if (e.Event == "Data Change" && e.Category == "Pay Rate") return "Pay Rate";
                        if (e.Event == "Data Change" && e.Category == "Payment Instruction") return "Payment Instruction";
                        if (e.Event == "Termination") return "Employment";
                        return "Data Change"; // Default category if none of the above
                    })
                    .ToDictionary(g => g.Key, g => g.ToList());

                var apiResponse = await _paySpaceApiService.SubmitAllCategoriesAsync(companyId, frequency, run, entriesByCategory);

                _logger.LogInformation($"PaySpace API call completed with success={apiResponse.Success}, " +
                                      $"successful entries={apiResponse.SuccessfulEntries}, " +
                                      $"failed entries={apiResponse.FailedEntries}");
                var responseData = new
                {
                    success = apiResponse.Success,
                    message = apiResponse.Message,
                    details = apiResponse.SuccessfulEntries > 0
                        ? $"Successfully processed {apiResponse.SuccessfulEntries} entries"
                        : "No entries were processed successfully",
                    failureDetails = apiResponse.FailedEntries > 0
                        ? $"Failed to process {apiResponse.FailedEntries} entries"
                        : null,
                    company = $"PaySpace Company ID {companyId}",
                    frequency = frequency,
                    run = run,
                    entriesCount = _lastMappedEntries.Count,
                    successfulEntries = apiResponse.SuccessfulEntries,
                    failedEntries = apiResponse.FailedEntries,
                    errors = apiResponse.Errors.Any() ? apiResponse.Errors : null,
                    results = apiResponse.Results, // These are results from PaySpace API
                    uploadHistorySaved = false,
                    uploadHistoryId = 0,
                    entriesSavedToDb = 0 // Renamed for clarity
                };
                try
                {
                    string companyName = "Unknown";
                    string companyCode = "Unknown";

                    try
                    {
                        var companies = _paySpaceApiService.GetAvailableCompanies();
                        var company = companies.FirstOrDefault(c => Convert.ToInt32(((dynamic)c).id) == companyId);
                        if (company != null)
                        {
                            companyName = ((dynamic)company).companyName;
                            companyCode = ((dynamic)company).companyCode;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "Could not retrieve company details for history record");
                    }
                    var uploadHistory = new UploadHistory
                    {
                        UploadTimestamp = DateTime.UtcNow,
                        CompanyId = companyId,
                        CompanyName = companyName ?? "Unknown Company",
                        CompanyCode = companyCode ?? "Unknown",
                        Frequency = frequency ?? "Unknown Frequency",
                        Run = run ?? "Unknown Run",
                        TotalEntries = _lastMappedEntries.Count,
                        SuccessfulEntries = apiResponse.SuccessfulEntries, // From PaySpace response
                        FailedEntries = apiResponse.FailedEntries,     // From PaySpace response
                        Success = apiResponse.Success,                 // Overall success from PaySpace
                        FileName = "Manual Upload"
                    };
                    _dbContext.UploadHistories.Add(uploadHistory);
                    await _dbContext.SaveChangesAsync();
                    _logger.LogInformation($"Successfully saved upload history with ID {uploadHistory.Id}");

                    int entriesSavedToDbCount = 0;
                    if (apiResponse.Results != null && apiResponse.Results.Any())
                    {
                        foreach (var result in apiResponse.Results.Take(50)) // Process up to 50 results for history details
                        {
                            try
                            {
                                // Find the original entry that corresponds to this result
                                var originalEntry = _lastMappedEntries.FirstOrDefault(e =>
                                    e.EmployeeId == result.EmployeeId && // Assuming result has EmployeeId
                                    e.PayElementId == result.PayElementId); // Assuming result has PayElementId

                                string historyInputType = "Amount"; // Default for history record
                                decimal historyInputValue = 0;

                                if (originalEntry != null)
                                {
                                    string entryUnitTypeLower = originalEntry.UnitType?.ToLowerInvariant();
                                    if (entryUnitTypeLower == "days")
                                    {
                                        historyInputType = "Days";
                                        historyInputValue = originalEntry.NumberOfUnits ?? 0;
                                    }
                                    else if (entryUnitTypeLower == "hours")
                                    {
                                        historyInputType = "Hours";
                                        historyInputValue = originalEntry.NumberOfUnits ?? 0;
                                    }
                                    else if (entryUnitTypeLower == "units")
                                    {
                                        historyInputType = "Units"; // Specific for history if "units"
                                        historyInputValue = originalEntry.NumberOfUnits ?? 0;
                                    }
                                    // If UnitType is something else (e.g. "pieces") and NumberOfUnits has a value,
                                    // and Amount is null, this logic defaults to Amount=0 for history.
                                    // This aligns with export logic: if not days/hours/units, primary value is Amount.
                                    else if (originalEntry.Amount.HasValue)
                                    {
                                        historyInputType = "Amount";
                                        historyInputValue = originalEntry.Amount.Value;
                                    }
                                    // Fallback if UnitType is e.g. "PIECES" and only NumberOfUnits is set
                                    else if (!string.IsNullOrEmpty(originalEntry.UnitType) && originalEntry.NumberOfUnits.HasValue)
                                    {
                                        historyInputType = originalEntry.UnitType; // Use original unit type if not standard
                                        historyInputValue = originalEntry.NumberOfUnits.Value;
                                    }
                                }

                                string empId = (result.EmployeeId ?? "Unknown").Substring(0, Math.Min(50, (result.EmployeeId ?? "Unknown").Length));
                                string empName = (originalEntry?.EmployeeName ?? "Unknown").Substring(0, Math.Min(100, (originalEntry?.EmployeeName ?? "Unknown").Length));
                                string payElementIdStr = (result.PayElementId ?? "Unknown").Substring(0, Math.Min(50, (result.PayElementId ?? "Unknown").Length));
                                string compCode = (result.ComponentCode ?? originalEntry?.PaySpaceCompCode ?? "").Substring(0, Math.Min(50, (result.ComponentCode ?? originalEntry?.PaySpaceCompCode ?? "").Length));
                                string errorMsg = result.ErrorMessage;

                                if (errorMsg != null && errorMsg.Length > 500) errorMsg = errorMsg.Substring(0, 500);
                                else if (errorMsg == null) errorMsg = "";

                                var historyEntry = new UploadHistoryEntry
                                {
                                    UploadHistoryId = uploadHistory.Id,
                                    EmployeeId = empId,
                                    EmployeeName = empName,
                                    PayElementId = payElementIdStr,
                                    ComponentCode = compCode,
                                    Success = result.Success, // Success of this specific entry from PaySpace
                                    ErrorMessage = errorMsg,
                                    Amount = historyInputValue,
                                    InputType = historyInputType,
                                    InputQuantity = historyInputValue
                                };
                                bool entrySaved = await _uploadHistoryService.SaveEntryAsync(historyEntry);
                                if (entrySaved) entriesSavedToDbCount++;
                            }
                            catch (Exception entryEx)
                            {
                                _logger.LogError(entryEx, $"Error saving history entry for employee {result.EmployeeId}, PayElement {result.PayElementId}");
                            }
                        }
                    }
                    _logger.LogInformation($"Saved {entriesSavedToDbCount} detailed entry results to history out of {apiResponse.Results?.Count ?? 0} results from PaySpace.");
                    return Ok(new
                    {
                        responseData.success,
                        responseData.message,
                        responseData.details,
                        responseData.failureDetails,
                        responseData.company,
                        responseData.frequency,
                        responseData.run,
                        responseData.entriesCount,
                        responseData.successfulEntries,
                        responseData.failedEntries,
                        responseData.errors,
                        responseData.results,
                        uploadHistorySaved = true,
                        uploadHistoryId = uploadHistory.Id,
                        entriesSavedToDb = entriesSavedToDbCount
                    });
                }
                catch (Exception dbEx)
                {
                    _logger.LogError(dbEx, "Database error saving upload history");
                    return Ok(responseData);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error uploading to PaySpace for company {companyId} and frequency {frequency}");
                return StatusCode(500, new { error = "Error uploading to PaySpace", details = ex.Message });
            }
        }


        [HttpGet("test-connection")]
        public IActionResult TestConnection()
        {
            try
            {
                var legalEntity = _dbContext.LegalEntities.FirstOrDefault();
                if (legalEntity == null)
                {
                    return BadRequest(new { success = false, message = "No legal entities found in database" });
                }

                _logger.LogInformation("Testing PaySpace API connection");
                var companies = _paySpaceApiService.GetAvailableFrequencies(legalEntity.Id); // This was GetAvailableCompanies before, assuming Frequencies is a valid test

                if (companies != null) // Check if the result itself is not null; content check might be needed
                {
                    _logger.LogInformation("Connection successful (API call returned data)");
                    return Ok(new { success = true, message = "Connection successful" });
                }

                _logger.LogWarning("Connection test returned null or empty result from PaySpace API.");
                return BadRequest(new { success = false, message = "Connection test returned null or empty result." });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error testing PaySpace API connection");
                return StatusCode(500, new { success = false, error = ex.Message, stackTrace = ex.StackTrace });
            }
        }

        // Unified ExportToExcel method
        private IActionResult ExportToExcel(List<PayElementEntry> entries, string fileNamePrefix)
        {
            try
            {
                using var workbook = new XSSFWorkbook();
                using var memoryStream = new MemoryStream();
                var sheet = workbook.CreateSheet("Entries");

                var headerStyle = workbook.CreateCellStyle();
                var headerFont = workbook.CreateFont();
                headerFont.IsBold = true;
                headerStyle.SetFont(headerFont);

                // MODIFIED headers: Removed "Unit Type", kept "Input Type" and "Value"
                var headers = new List<string>
                {
                    "Logical ID", "Record Number", "Person ID", "Employee ID", "Employee Name", "Event", "Action", "Category",
                    "Start Date", "End Date", "Cost Center", "Position Title",
                    "Termination Date", "Termination Reason",
                    "First Name", "Last Name", "Title", "Gender", "Language", "Citizenship Country", "Birth Date",
                    "Pay Element ID", "Pay Element Type", "PaySpace Component Code",
                    "Input Type", "Value" // Consolidated from "Unit Type", "Number of Units", "Amount"
                };

                var headerRow = sheet.CreateRow(0);
                for (int i = 0; i < headers.Count; i++)
                {
                    var cell = headerRow.CreateCell(i);
                    cell.SetCellValue(headers[i]);
                    cell.CellStyle = headerStyle;
                }

                int rowNum = 1;
                foreach (var entry in entries)
                {
                    var row = sheet.CreateRow(rowNum++);
                    int col = 0;

                    // Common Fields
                    row.CreateCell(col++).SetCellValue(entry.RowId ?? "");
                    row.CreateCell(col++).SetCellValue(entry.RecordNumber ?? "");
                    row.CreateCell(col++).SetCellValue(entry.PersonId ?? "");
                    row.CreateCell(col++).SetCellValue(entry.EmployeeId ?? "");
                    row.CreateCell(col++).SetCellValue(entry.EmployeeName ?? "");
                    row.CreateCell(col++).SetCellValue(entry.Event ?? "");
                    row.CreateCell(col++).SetCellValue(entry.Action ?? "");
                    row.CreateCell(col++).SetCellValue(entry.Category ?? "");
                    row.CreateCell(col++).SetCellValue(entry.StartDate != default ? entry.StartDate.ToString("yyyy-MM-dd") : "");
                    row.CreateCell(col++).SetCellValue(entry.EndDate != default ? entry.EndDate.ToString("yyyy-MM-dd") : "");
                    row.CreateCell(col++).SetCellValue(entry.CostCenter ?? "");
                    row.CreateCell(col++).SetCellValue(entry.PositionTitle ?? "");
                    row.CreateCell(col++).SetCellValue(entry.TerminationDate?.ToString("yyyy-MM-dd") ?? "");
                    row.CreateCell(col++).SetCellValue(entry.TerminationReason ?? "");
                    row.CreateCell(col++).SetCellValue(entry.EmployeeFirstName ?? "");
                    row.CreateCell(col++).SetCellValue(entry.EmployeeLastName ?? "");
                    row.CreateCell(col++).SetCellValue(entry.Title ?? "");
                    row.CreateCell(col++).SetCellValue(entry.Gender ?? "");
                    row.CreateCell(col++).SetCellValue(entry.Language ?? "");
                    row.CreateCell(col++).SetCellValue(entry.CitizenshipCountry ?? "");
                    row.CreateCell(col++).SetCellValue(entry.BirthDate.HasValue ? entry.BirthDate.Value.ToString("yyyy-MM-dd") : "");
                    row.CreateCell(col++).SetCellValue(entry.PayElementId ?? "");
                    row.CreateCell(col++).SetCellValue(entry.PayElementType ?? "");
                    row.CreateCell(col++).SetCellValue(entry.PaySpaceCompCode ?? ""); // This is the mapped one for MappedEntries export

                    // Derived Input Type and Value logic
                    string originalUnitType = entry.UnitType;
                    string derivedInputType;
                    string derivedValueString;

                    string unitTypeLower = originalUnitType?.ToLowerInvariant();

                    if (unitTypeLower == "days")
                    {
                        derivedInputType = "Days";
                        derivedValueString = entry.NumberOfUnits?.ToString() ?? "";
                    }
                    else if (unitTypeLower == "hours")
                    {
                        derivedInputType = "Hours";
                        derivedValueString = entry.NumberOfUnits?.ToString() ?? "";
                    }
                    else // Includes null, empty, "units", or any other UnitType
                    {
                        derivedInputType = "Amount";
                        // If UnitType was present (e.g., "units", "pieces") and NumberOfUnits has a value,
                        // prioritize NumberOfUnits. Otherwise, use Amount.
                        if (!string.IsNullOrEmpty(originalUnitType) && originalUnitType.ToLowerInvariant() != "days" && originalUnitType.ToLowerInvariant() != "hours" && entry.NumberOfUnits.HasValue)
                        {
                            derivedValueString = entry.NumberOfUnits.Value.ToString();
                        }
                        else
                        {
                            derivedValueString = entry.Amount?.ToString() ?? "";
                        }
                    }

                    row.CreateCell(col++).SetCellValue(derivedInputType);   // Input Type (derived)
                    row.CreateCell(col++).SetCellValue(derivedValueString); // Value (derived)
                }

                for (int i = 0; i < headers.Count; i++)
                {
                    sheet.AutoSizeColumn(i);
                }

                workbook.Write(memoryStream);
                var byteArray = memoryStream.ToArray();

                return File(
                    byteArray,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    $"{fileNamePrefix}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
                );
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error exporting to Excel");
                return StatusCode(500, new { error = "Error creating Excel file", details = ex.Message });
            }
        }
    }
}
