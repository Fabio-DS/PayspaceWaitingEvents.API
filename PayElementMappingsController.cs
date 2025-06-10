using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PaySpaceWaitingEvents.API.Data;
using PaySpaceWaitingEvents.API.Models;
using PaySpaceWaitingEvents.API.Models.PaySpaceWaitingEvents.API.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace PaySpaceWaitingEvents.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class PayElementMappingsController : ControllerBase
    {
        private readonly ApplicationDbContext _dbContext;
        private readonly ILogger<PayElementMappingsController> _logger;

        public PayElementMappingsController(
            ApplicationDbContext dbContext,
            ILogger<PayElementMappingsController> logger)
        {
            _dbContext = dbContext;
            _logger = logger;
        }

        [HttpGet("entity/{legalEntityId}")]
        public async Task<IActionResult> GetMappings(int legalEntityId)
        {
            var mappings = await _dbContext.PayElementMappings
                .Where(m => m.LegalEntityId == legalEntityId)
                .ToListAsync();

            return Ok(mappings);
        }

        [HttpPost]
        public async Task<IActionResult> CreateMapping([FromBody] CreateMappingRequest request)
        {
            try
            {
                _logger.LogInformation($"Creating new mapping. Request data: {System.Text.Json.JsonSerializer.Serialize(request)}");

                if (request == null)
                {
                    _logger.LogWarning("Request body was null for creating mapping");
                    return BadRequest(new { error = "No data provided" });
                }

                if (request.LegalEntityId <= 0)
                {
                    _logger.LogWarning($"Invalid LegalEntityId: {request.LegalEntityId}");
                    return BadRequest(new { error = "Invalid legal entity ID" });
                }

                if (string.IsNullOrWhiteSpace(request.PayElementId))
                {
                    _logger.LogWarning("Missing PayElementId in request");
                    return BadRequest(new { error = "Pay Element ID is required" });
                }

                if (string.IsNullOrWhiteSpace(request.ComponentCode))
                {
                    _logger.LogWarning("Missing ComponentCode in request");
                    return BadRequest(new { error = "Component Code is required" });
                }

                if (string.IsNullOrWhiteSpace(request.Frequency))
                {
                    _logger.LogWarning("Missing Frequency in request");
                    return BadRequest(new { error = "Frequency is required" });
                }

                var mapping = new PayElementMapping
                {
                    LegalEntityId = request.LegalEntityId,
                    PayElementId = request.PayElementId,
                    ComponentCode = request.ComponentCode,
                    Frequency = request.Frequency,
                    Description = request.Description ?? "",
                    IsActive = request.IsActive,
                    CreatedDate = DateTime.UtcNow
                };

                var legalEntity = await _dbContext.LegalEntities.FindAsync(mapping.LegalEntityId);
                if (legalEntity == null)
                {
                    _logger.LogWarning($"Legal entity ID {mapping.LegalEntityId} not found");
                    return BadRequest(new { error = $"Legal entity with id {mapping.LegalEntityId} not found" });
                }

                _logger.LogInformation($"Adding new mapping: {System.Text.Json.JsonSerializer.Serialize(mapping)}");
                await _dbContext.PayElementMappings.AddAsync(mapping);
                await _dbContext.SaveChangesAsync();

                return Ok(mapping);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating mapping");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> UpdateMapping(int id, [FromBody] UpdateMappingRequest request)
        {
            try
            {
                _logger.LogInformation($"Updating mapping ID {id}. Request data: {System.Text.Json.JsonSerializer.Serialize(request)}");

                if (request == null)
                {
                    _logger.LogWarning($"Request body was null for updating mapping ID {id}");
                    return BadRequest(new { error = "No data provided" });
                }

                var existingMapping = await _dbContext.PayElementMappings.FindAsync(id);
                if (existingMapping == null)
                {
                    _logger.LogWarning($"Mapping with ID {id} not found");
                    return NotFound(new { error = $"Mapping with id {id} not found" });
                }

                existingMapping.PayElementId = request.PayElementId;
                existingMapping.ComponentCode = request.ComponentCode;
                existingMapping.Frequency = request.Frequency;
                existingMapping.Description = request.Description ?? existingMapping.Description;
                existingMapping.IsActive = request.IsActive;
                existingMapping.LastModifiedDate = DateTime.UtcNow;

                _logger.LogInformation($"Saving updated mapping. New values: {System.Text.Json.JsonSerializer.Serialize(existingMapping)}");
                await _dbContext.SaveChangesAsync();

                return Ok(existingMapping);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error updating mapping ID {id}");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteMapping(int id)
        {
            try
            {
                var mapping = await _dbContext.PayElementMappings.FindAsync(id);

                if (mapping == null)
                    return NotFound($"Mapping with id {id} not found");

                _dbContext.PayElementMappings.Remove(mapping);
                await _dbContext.SaveChangesAsync();

                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting mapping");
                return StatusCode(500, new { error = ex.Message });
            }
        }

        [HttpGet("download-template")]
        public IActionResult DownloadTemplate()
        {
            try
            {
                using (var workbook = new XSSFWorkbook())
                using (var memoryStream = new MemoryStream())
                {
                    var sheet = workbook.CreateSheet("Mappings");

                    var headerStyle = workbook.CreateCellStyle();
                    var headerFont = workbook.CreateFont();
                    headerFont.IsBold = true;
                    headerStyle.SetFont(headerFont);

                    var headerRow = sheet.CreateRow(0);
                    var headers = new[] {
                        "LegalEntityId",
                        "PayElementId",
                        "ComponentCode",
                        "Frequency",
                        "Description",
                        "IsActive"
                    };

                    for (int i = 0; i < headers.Length; i++)
                    {
                        var cell = headerRow.CreateCell(i);
                        cell.SetCellValue(headers[i]);
                        cell.CellStyle = headerStyle;
                    }

                    var exampleRow = sheet.CreateRow(1);
                    exampleRow.CreateCell(0).SetCellValue("1"); // LegalEntityId
                    exampleRow.CreateCell(1).SetCellValue("BASICSALARY"); // PayElementId
                    exampleRow.CreateCell(2).SetCellValue("1000"); // ComponentCode
                    exampleRow.CreateCell(3).SetCellValue("Monthly"); // Frequency
                    exampleRow.CreateCell(4).SetCellValue("Basic Salary Component"); // Payspace Component Description
                    exampleRow.CreateCell(5).SetCellValue("TRUE"); // IsActive

                    for (int i = 0; i < headers.Length; i++)
                    {
                        sheet.AutoSizeColumn(i);
                    }

                    workbook.Write(memoryStream);
                    var byteArray = memoryStream.ToArray();

                    return File(
                        byteArray,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        $"MappingTemplate_{DateTime.Now:yyyyMMdd}.xlsx"
                    );
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating template file");
                return StatusCode(500, new { error = "Error creating template file", details = ex.Message });
            }
        }

        [HttpPost("bulk-upload")]
        public async Task<IActionResult> BulkUpload(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded");
            }

            try
            {
                _logger.LogInformation($"Starting mapping bulk upload: {file.FileName}");

                using var stream = file.OpenReadStream();

                var result = new BulkUploadResult
                {
                    FileName = file.FileName,
                    TotalRows = 0,
                    SuccessCount = 0,
                    FailureCount = 0,
                    Errors = new List<string>()
                };

                if (file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    await ProcessExcelFile(stream, result);
                }
                else if (file.FileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    await ProcessExcelFileHSSF(stream, result);
                }
                else if (file.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    await ProcessCsvFile(stream, result);
                }
                else
                {
                    return BadRequest("Unsupported file format. Please upload an Excel (.xlsx, .xls) or CSV file.");
                }

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing bulk upload file");
                return StatusCode(500, new
                {
                    error = "Error processing file",
                    details = ex.Message,
                    fileName = file.FileName
                });
            }
        }

        private async Task ProcessExcelFile(Stream fileStream, BulkUploadResult result)
        {
            var workbook = new XSSFWorkbook(fileStream);
            var sheet = workbook.GetSheetAt(0);

            await ProcessSheet(sheet, result);
        }

        private async Task ProcessExcelFileHSSF(Stream fileStream, BulkUploadResult result)
        {
            var workbook = new HSSFWorkbook(fileStream);
            var sheet = workbook.GetSheetAt(0);

            await ProcessSheet(sheet, result);
        }

        private async Task ProcessCsvFile(Stream fileStream, BulkUploadResult result)
        {
            using var reader = new StreamReader(fileStream);

            // Read header
            var headerLine = await reader.ReadLineAsync();
            if (headerLine == null)
            {
                result.Errors.Add("Empty CSV file");
                return;
            }

            var headers = headerLine.Split(',');

            // Map column indices
            int legalEntityIdIndex = Array.IndexOf(headers, "LegalEntityId");
            int payElementIdIndex = Array.IndexOf(headers, "PayElementId");
            int componentCodeIndex = Array.IndexOf(headers, "ComponentCode");
            int frequencyIndex = Array.IndexOf(headers, "Frequency");
            int descriptionIndex = Array.IndexOf(headers, "Description");
            int isActiveIndex = Array.IndexOf(headers, "IsActive");

            // Validate header
            if (legalEntityIdIndex == -1 || payElementIdIndex == -1 ||
                componentCodeIndex == -1 || frequencyIndex == -1)
            {
                result.Errors.Add("CSV file must contain LegalEntityId, PayElementId, ComponentCode, and Frequency columns");
                return;
            }

            // Process data rows
            int rowNumber = 1;
            string line;
            var newMappings = new List<PayElementMapping>();

            while ((line = await reader.ReadLineAsync()) != null)
            {
                rowNumber++;
                result.TotalRows++;

                try
                {
                    var values = line.Split(',');

                    if (values.Length <= Math.Max(legalEntityIdIndex, Math.Max(payElementIdIndex,
                        Math.Max(componentCodeIndex, frequencyIndex))))
                    {
                        result.Errors.Add($"Row {rowNumber}: Not enough columns");
                        result.FailureCount++;
                        continue;
                    }

                    if (!int.TryParse(values[legalEntityIdIndex], out int legalEntityId))
                    {
                        result.Errors.Add($"Row {rowNumber}: Invalid LegalEntityId '{values[legalEntityIdIndex]}'");
                        result.FailureCount++;
                        continue;
                    }

                    var mapping = new PayElementMapping
                    {
                        LegalEntityId = legalEntityId,
                        PayElementId = values[payElementIdIndex],
                        ComponentCode = values[componentCodeIndex],
                        Frequency = values[frequencyIndex],
                        Description = descriptionIndex >= 0 && descriptionIndex < values.Length
                            ? values[descriptionIndex]
                            : "",
                        IsActive = isActiveIndex >= 0 && isActiveIndex < values.Length
                            ? values[isActiveIndex].Trim().Equals("TRUE", StringComparison.OrdinalIgnoreCase)
                            : true,
                        CreatedDate = DateTime.UtcNow
                    };

                    newMappings.Add(mapping);
                    result.SuccessCount++;
                }
                catch (Exception ex)
                {
                    result.Errors.Add($"Row {rowNumber}: {ex.Message}");
                    result.FailureCount++;
                }
            }

            // Add valid mappings to database
            if (newMappings.Any())
            {
                await _dbContext.PayElementMappings.AddRangeAsync(newMappings);
                await _dbContext.SaveChangesAsync();
            }
        }

        private async Task ProcessSheet(ISheet sheet, BulkUploadResult result)
        {
            var headerRow = sheet.GetRow(0);
            if (headerRow == null)
            {
                result.Errors.Add("Empty sheet");
                return;
            }

            int legalEntityIdIndex = -1;
            int payElementIdIndex = -1;
            int componentCodeIndex = -1;
            int frequencyIndex = -1;
            int descriptionIndex = -1;
            int isActiveIndex = -1;

            for (int i = 0; i < headerRow.LastCellNum; i++)
            {
                var cell = headerRow.GetCell(i);
                if (cell == null) continue;

                string headerValue = cell.StringCellValue.Trim();

                if (headerValue.Equals("LegalEntityId", StringComparison.OrdinalIgnoreCase))
                    legalEntityIdIndex = i;
                else if (headerValue.Equals("PayElementId", StringComparison.OrdinalIgnoreCase))
                    payElementIdIndex = i;
                else if (headerValue.Equals("ComponentCode", StringComparison.OrdinalIgnoreCase))
                    componentCodeIndex = i;
                else if (headerValue.Equals("Frequency", StringComparison.OrdinalIgnoreCase))
                    frequencyIndex = i;
                else if (headerValue.Equals("Description", StringComparison.OrdinalIgnoreCase))
                    descriptionIndex = i;
                else if (headerValue.Equals("IsActive", StringComparison.OrdinalIgnoreCase))
                    isActiveIndex = i;
            }

            if (legalEntityIdIndex == -1 || payElementIdIndex == -1 ||
                componentCodeIndex == -1 || frequencyIndex == -1)
            {
                result.Errors.Add("Excel file must contain LegalEntityId, PayElementId, ComponentCode, and Frequency columns");
                return;
            }

            var newMappings = new List<PayElementMapping>();

            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null) continue;

                result.TotalRows++;

                try
                {
                    // Get legal entity ID
                    var legalEntityIdCell = row.GetCell(legalEntityIdIndex);
                    if (legalEntityIdCell == null)
                    {
                        result.Errors.Add($"Row {i + 1}: Missing LegalEntityId");
                        result.FailureCount++;
                        continue;
                    }

                    int legalEntityId;
                    if (legalEntityIdCell.CellType == CellType.Numeric)
                    {
                        legalEntityId = (int)legalEntityIdCell.NumericCellValue;
                    }
                    else
                    {
                        if (!int.TryParse(legalEntityIdCell.StringCellValue, out legalEntityId))
                        {
                            result.Errors.Add($"Row {i + 1}: Invalid LegalEntityId '{legalEntityIdCell.StringCellValue}'");
                            result.FailureCount++;
                            continue;
                        }
                    }

                    // Get PayElementId
                    var payElementIdCell = row.GetCell(payElementIdIndex);
                    if (payElementIdCell == null)
                    {
                        result.Errors.Add($"Row {i + 1}: Missing PayElementId");
                        result.FailureCount++;
                        continue;
                    }
                    string payElementId = GetCellValueAsString(payElementIdCell);

                    // Get ComponentCode
                    var componentCodeCell = row.GetCell(componentCodeIndex);
                    if (componentCodeCell == null)
                    {
                        result.Errors.Add($"Row {i + 1}: Missing ComponentCode");
                        result.FailureCount++;
                        continue;
                    }
                    string componentCode = GetCellValueAsString(componentCodeCell);

                    // Get Frequency
                    var frequencyCell = row.GetCell(frequencyIndex);
                    if (frequencyCell == null)
                    {
                        result.Errors.Add($"Row {i + 1}: Missing Frequency");
                        result.FailureCount++;
                        continue;
                    }
                    string frequency = GetCellValueAsString(frequencyCell);

                    // Get Description (optional)
                    string description = "";
                    if (descriptionIndex >= 0)
                    {
                        var descriptionCell = row.GetCell(descriptionIndex);
                        if (descriptionCell != null)
                        {
                            description = GetCellValueAsString(descriptionCell);
                        }
                    }

                    bool isActive = true;
                    if (isActiveIndex >= 0)
                    {
                        var isActiveCell = row.GetCell(isActiveIndex);
                        if (isActiveCell != null)
                        {
                            string isActiveStr = GetCellValueAsString(isActiveCell);
                            bool.TryParse(isActiveStr, out isActive);
                        }
                    }

                    var mapping = new PayElementMapping
                    {
                        LegalEntityId = legalEntityId,
                        PayElementId = payElementId,
                        ComponentCode = componentCode,
                        Frequency = frequency,
                        Description = description,
                        IsActive = isActive,
                        CreatedDate = DateTime.UtcNow
                    };

                    newMappings.Add(mapping);
                    result.SuccessCount++;
                }
                catch (Exception ex)
                {
                    result.Errors.Add($"Row {i + 1}: {ex.Message}");
                    result.FailureCount++;
                }
            }

            if (newMappings.Any())
            {
                await _dbContext.PayElementMappings.AddRangeAsync(newMappings);
                await _dbContext.SaveChangesAsync();
            }
        }

        private string GetCellValueAsString(ICell cell)
        {
            if (cell == null) return string.Empty;

            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue?.Trim() ?? string.Empty;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                        return string.Format("{0:yyyy-MM-dd}", cell.DateCellValue);
                    return cell.NumericCellValue.ToString();
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Formula:
                    switch (cell.CachedFormulaResultType)
                    {
                        case CellType.String:
                            return cell.StringCellValue?.Trim() ?? string.Empty;
                        case CellType.Numeric:
                            return cell.NumericCellValue.ToString();
                        default:
                            return string.Empty;
                    }
                default:
                    return string.Empty;
            }
        }
    }

    public class BulkUploadResult
    {
        public string FileName { get; set; }
        public int TotalRows { get; set; }
        public int SuccessCount { get; set; }
        public int FailureCount { get; set; }
        public List<string> Errors { get; set; }
    }

    public class CreateMappingRequest
    {
        public int LegalEntityId { get; set; }
        public string PayElementId { get; set; }
        public string ComponentCode { get; set; }
        public string Frequency { get; set; }
        public string Description { get; set; }
        public bool IsActive { get; set; } = true;
    }

    public class UpdateMappingRequest
    {
        public string PayElementId { get; set; }
        public string ComponentCode { get; set; }
        public string Frequency { get; set; }
        public string Description { get; set; }
        public bool IsActive { get; set; }
    }
}