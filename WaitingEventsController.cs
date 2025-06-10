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
                        var mappedEntry = new PayElementEntry
                        {
                            PersonId = entry.PersonId,
                            EmployeeId = entry.EmployeeId,
                            EmployeeName = entry.EmployeeName,
                            Event = entry.Event,
                            Action = entry.Action,
                            RowId = entry.RowId,
                            StartDate = entry.StartDate,
                            EndDate = entry.EndDate,
                            Amount = entry.Amount,
                            NumberOfUnits = entry.NumberOfUnits,
                            UnitType = entry.UnitType,
                            PayElementId = entry.PayElementId,
                            PayElementType = entry.PayElementType,
                            PaySpaceCompCode = componentCode
                        };

                        mappedEntries.Add(mappedEntry);
                    }
                    else
                    {
                        unmappedElements.Add(entry.PayElementId);
                    }
                }

                _lastMappedEntries = mappedEntries;

                if (unmappedElements.Any())
                {
                    return BadRequest(new
                    {
                        Message = "Mapping failed. There are unmapped pay elements.",
                        TotalOriginalEntries = _lastProcessedEntries.Count,
                        UnmappedEntries = unmappedElements.Count,
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
                        TotalAmount = mappedEntries.Sum(e => e.Amount ?? 0),
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

            return ExportToExcel(_lastProcessedEntries, "OriginalEntries");
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

                // Group entries by category for processing
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
                        return "Data Change";
                    })
                    .ToDictionary(g => g.Key, g => g.ToList());

                // Step 1: Call PaySpace API for all categories
                var apiResponse = await _paySpaceApiService.SubmitAllCategoriesAsync(companyId, frequency, run, entriesByCategory);

                _logger.LogInformation($"PaySpace API call completed with success={apiResponse.Success}, " +
                                      $"successful entries={apiResponse.SuccessfulEntries}, " +
                                      $"failed entries={apiResponse.FailedEntries}");

                // Step 2: Create the initial response object
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
                    results = apiResponse.Results,
                    uploadHistorySaved = false,
                    uploadHistoryId = 0,
                    entriesSaved = 0
                };

                // Step 3: Try to save history, but continue even if it fails
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

                    // Create history record
                    var uploadHistory = new UploadHistory
                    {
                        UploadTimestamp = DateTime.UtcNow,
                        CompanyId = companyId,
                        CompanyName = companyName ?? "Unknown Company",
                        CompanyCode = companyCode ?? "Unknown",
                        Frequency = frequency ?? "Unknown Frequency",
                        Run = run ?? "Unknown Run",
                        TotalEntries = _lastMappedEntries.Count,
                        SuccessfulEntries = apiResponse.SuccessfulEntries,
                        FailedEntries = apiResponse.FailedEntries,
                        Success = apiResponse.Success,
                        FileName = "Manual Upload"
                    };

                    // Save 
                    _dbContext.UploadHistories.Add(uploadHistory);
                    await _dbContext.SaveChangesAsync();

                    _logger.LogInformation($"Successfully saved upload history with ID {uploadHistory.Id}");

                    // Now process entries with direct SQL
                    int entriesSaved = 0;

                    if (apiResponse.Results != null && apiResponse.Results.Any())
                    {
                        foreach (var result in apiResponse.Results.Take(50))
                        {
                            try
                            {
                                var originalEntry = _lastMappedEntries.FirstOrDefault(e =>
                                    e.EmployeeId == result.EmployeeId &&
                                    e.PayElementId == result.PayElementId);

                                // Determine input type and value
                                string inputType = "Amount"; // Default
                                decimal inputValue = 0;

                                if (originalEntry != null)
                                {
                                    if (!string.IsNullOrEmpty(originalEntry.UnitType))
                                    {
                                        if (originalEntry.UnitType.Equals("days", StringComparison.OrdinalIgnoreCase))
                                        {
                                            inputType = "Days";
                                            inputValue = originalEntry.NumberOfUnits ?? 0;
                                        }
                                        else if (originalEntry.UnitType.Equals("hours", StringComparison.OrdinalIgnoreCase))
                                        {
                                            inputType = "Hours";
                                            inputValue = originalEntry.NumberOfUnits ?? 0;
                                        }
                                        else if (originalEntry.UnitType.Equals("units", StringComparison.OrdinalIgnoreCase))
                                        {
                                            inputType = "Units";
                                            inputValue = originalEntry.NumberOfUnits ?? 0;
                                        }
                                        else
                                        {
                                            inputValue = originalEntry.Amount ?? 0;
                                        }
                                    }
                                    else if (originalEntry.Amount.HasValue)
                                    {
                                        inputValue = originalEntry.Amount.Value;
                                    }
                                }

                                // Prepare data with truncation
                                string empId = (result.EmployeeId ?? "Unknown").Substring(0, Math.Min(50, (result.EmployeeId ?? "Unknown").Length));
                                string empName = (originalEntry?.EmployeeName ?? "Unknown").Substring(0, Math.Min(100, (originalEntry?.EmployeeName ?? "Unknown").Length));
                                string payElementId = (result.PayElementId ?? "Unknown").Substring(0, Math.Min(50, (result.PayElementId ?? "Unknown").Length));
                                string compCode = (result.ComponentCode ?? "").Substring(0, Math.Min(50, (result.ComponentCode ?? "").Length));
                                string errorMsg = result.ErrorMessage;

                                if (errorMsg != null && errorMsg.Length > 500)
                                    errorMsg = errorMsg.Substring(0, 500);
                                else if (errorMsg == null)
                                    errorMsg = "";

                                // Create entry with new fields
                                var historyEntry = new UploadHistoryEntry
                                {
                                    UploadHistoryId = uploadHistory.Id,
                                    EmployeeId = empId,
                                    EmployeeName = empName,
                                    PayElementId = payElementId,
                                    ComponentCode = compCode,
                                    Success = result.Success,
                                    ErrorMessage = errorMsg,
                                    Amount = inputValue,
                                    InputType = inputType,
                                    InputQuantity = inputValue
                                };

                                // Use upload history service to save entry
                                bool entrySaved = await _uploadHistoryService.SaveEntryAsync(historyEntry);
                                if (entrySaved)
                                {
                                    entriesSaved++;
                                }
                            }
                            catch (Exception entryEx)
                            {
                                _logger.LogError(entryEx, $"Error saving entry for employee {result.EmployeeId}");
                            }
                        }
                    }

                    _logger.LogInformation($"Saved {entriesSaved} out of {apiResponse.Results?.Count ?? 0} entries");

                    // Create the updated response with history details
                    return Ok(new
                    {
                        success = responseData.success,
                        message = responseData.message,
                        details = responseData.details,
                        failureDetails = responseData.failureDetails,
                        company = responseData.company,
                        frequency = responseData.frequency,
                        run = responseData.run,
                        entriesCount = responseData.entriesCount,
                        successfulEntries = responseData.successfulEntries,
                        failedEntries = responseData.failedEntries,
                        errors = responseData.errors,
                        results = responseData.results,
                        uploadHistorySaved = true,
                        uploadHistoryId = uploadHistory.Id,
                        entriesSaved = entriesSaved
                    });
                }
                catch (Exception dbEx)
                {
                    _logger.LogError(dbEx, "Database error saving upload history");
                    // Return original response if database saving failed
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
                var companies = _paySpaceApiService.GetAvailableFrequencies(legalEntity.Id);

                if (companies != null)
                {
                    _logger.LogInformation("Connection successful");
                    return Ok(new { success = true, message = "Connection successful" });
                }

                _logger.LogWarning("Connection test returned null result");
                return BadRequest(new { success = false, message = "Connection test returned null result" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error testing PaySpace API connection");
                return StatusCode(500, new { success = false, error = ex.Message, stackTrace = ex.StackTrace });
            }
        }

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

                // master headers
                var headers = new List<string>
        {
            "Logical ID", "Record Number", "Person ID", "Employee ID", "Employee Name", "Event", "Action", "Category",
            "Start Date", "End Date", "Cost Center", "Position Title",
            "Termination Date", "Termination Reason",
            "First Name", "Last Name", "Title", "Gender", "Language", "Citizenship Country", "Birth Date",
            "Pay Element ID", "Pay Element Type", "PaySpace Component Code", "Unit Type", "Input Type", "Number of Units", "Amount"
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

                    row.CreateCell(col++).SetCellValue(entry.RowId ?? "");
                    row.CreateCell(col++).SetCellValue(entry.RecordNumber ?? "");
                    row.CreateCell(col++).SetCellValue(entry.PersonId ?? "");
                    row.CreateCell(col++).SetCellValue(entry.EmployeeId ?? "");
                    row.CreateCell(col++).SetCellValue(entry.EmployeeName ?? "");
                    row.CreateCell(col++).SetCellValue(entry.Event ?? "");
                    row.CreateCell(col++).SetCellValue(entry.Action ?? "");
                    row.CreateCell(col++).SetCellValue(entry.Category ?? "");

                    // Core values
                    row.CreateCell(col++).SetCellValue(entry.StartDate != default ? entry.StartDate.ToString("yyyy-MM-dd") : "");
                    row.CreateCell(col++).SetCellValue(entry.EndDate != default ? entry.EndDate.ToString("yyyy-MM-dd") : "");
                    row.CreateCell(col++).SetCellValue(entry.CostCenter ?? "");
                    row.CreateCell(col++).SetCellValue(entry.PositionTitle ?? "");

                    // Termination
                    row.CreateCell(col++).SetCellValue(entry.TerminationDate?.ToString("yyyy-MM-dd") ?? "");
                    row.CreateCell(col++).SetCellValue(entry.TerminationReason ?? "");

                    // Personal Data
                    row.CreateCell(col++).SetCellValue(entry.EmployeeFirstName ?? "");
                    row.CreateCell(col++).SetCellValue(entry.EmployeeLastName ?? "");
                    row.CreateCell(col++).SetCellValue(entry.Title ?? "");
                    row.CreateCell(col++).SetCellValue(entry.Gender ?? "");
                    row.CreateCell(col++).SetCellValue(entry.Language ?? "");
                    row.CreateCell(col++).SetCellValue(entry.CitizenshipCountry ?? "");
                    row.CreateCell(col++).SetCellValue(entry.BirthDate.HasValue ? entry.BirthDate.Value.ToString("yyyy-MM-dd") : "");

                    // Pay Element / Pay Rate
                    row.CreateCell(col++).SetCellValue(entry.PayElementId ?? "");
                    row.CreateCell(col++).SetCellValue(entry.PayElementType ?? "");
                    row.CreateCell(col++).SetCellValue(entry.PaySpaceCompCode ?? "");
                    row.CreateCell(col++).SetCellValue(entry.UnitType ?? "");
                    row.CreateCell(col++).SetCellValue(entry.UnitType?.ToLower() switch
                    {
                        "days" => "Days",
                        "hours" => "Hours",
                        "units" => "Units",
                        _ => "Amount"
                    });
                    row.CreateCell(col++).SetCellValue(entry.NumberOfUnits?.ToString() ?? "");
                    row.CreateCell(col++).SetCellValue(entry.Amount?.ToString() ?? "");
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