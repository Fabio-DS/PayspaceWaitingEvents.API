using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using NPOI.SS.UserModel;
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
    public class UploadHistoryController : ControllerBase
    {
        private readonly ApplicationDbContext _dbContext;
        private readonly ILogger<UploadHistoryController> _logger;
        private readonly IUploadHistoryService _uploadHistoryService;

        public UploadHistoryController(
            ApplicationDbContext dbContext,
            ILogger<UploadHistoryController> logger,
            IUploadHistoryService uploadHistoryService)
        {
            _dbContext = dbContext;
            _logger = logger;
            _uploadHistoryService = uploadHistoryService;
        }

        [HttpGet]
        public async Task<IActionResult> GetUploadHistory()
        {
            try
            {
                _logger.LogInformation("Getting upload history");

                var history = await _dbContext.UploadHistories
                    .OrderByDescending(h => h.UploadTimestamp)
                    .Select(h => new
                    {
                        h.Id,
                        UploadDate = h.UploadTimestamp.ToString("yyyy-MM-dd HH:mm:ss"),
                        h.CompanyName,
                        h.CompanyCode,
                        h.Frequency,
                        h.Run,
                        h.TotalEntries,
                        h.SuccessfulEntries,
                        h.FailedEntries,
                        h.Success,
                        h.FileName
                    })
                    .ToListAsync();

                return Ok(history);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving upload history");
                return StatusCode(500, new { error = "Error retrieving upload history", details = ex.Message });
            }
        }

        [HttpGet("{id}")]
        public async Task<IActionResult> GetUploadHistoryDetails(int id)
        {
            try
            {
                _logger.LogInformation($"Getting details for upload history ID {id}");

                // Use the service to get history with entries
                var history = await _uploadHistoryService.GetUploadHistoryWithEntriesAsync(id);

                if (history == null)
                {
                    return NotFound($"Upload history with ID {id} not found");
                }

                _logger.LogInformation($"Retrieved history ID {id} with {history.Entries?.Count ?? 0} entries");

                var result = new
                {
                    history.Id,
                    UploadDate = history.UploadTimestamp.ToString("yyyy-MM-dd HH:mm:ss"),
                    history.CompanyName,
                    history.CompanyCode,
                    history.Frequency,
                    history.Run,
                    history.TotalEntries,
                    history.SuccessfulEntries,
                    history.FailedEntries,
                    history.Success,
                    Entries = (history.Entries?.Select(e => new
                    {
                        e.EmployeeId,
                        e.EmployeeName,
                        e.PayElementId,
                        e.ComponentCode,
                        e.Success,
                        e.ErrorMessage,
                        Amount = e.Amount.HasValue ? e.Amount.Value.ToString("0.00") : null,
                        e.InputType,
                        InputQuantity = e.InputQuantity.HasValue ? e.InputQuantity.Value.ToString("0.00") : null
                    }) ?? Enumerable.Empty<object>()).ToList()
                };

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error retrieving upload history details for ID {id}");
                return StatusCode(500, new { error = "Error retrieving upload history details", details = ex.Message });
            }
        }

        [HttpGet("download/{id}")]
        public async Task<IActionResult> DownloadUploadHistory(int id)
        {
            try
            {
                _logger.LogInformation($"Preparing download for upload history ID {id}");

                var history = await _dbContext.UploadHistories
                    .Include(h => h.Entries)
                    .FirstOrDefaultAsync(h => h.Id == id);

                if (history == null)
                {
                    return NotFound($"Upload history with ID {id} not found");
                }

                // Create Excel workbook
                using (var workbook = new XSSFWorkbook())
                using (var memoryStream = new MemoryStream())
                {
                    // Create Summary sheet
                    var summarySheet = workbook.CreateSheet("Summary");
                    var headerStyle = workbook.CreateCellStyle();
                    var headerFont = workbook.CreateFont();
                    headerFont.IsBold = true;
                    headerStyle.SetFont(headerFont);

                    // Summary headers
                    var summaryHeaderRow = summarySheet.CreateRow(0);
                    var summaryHeaders = new[] {
                        "Property", "Value"
                    };

                    for (int i = 0; i < summaryHeaders.Length; i++)
                    {
                        var cell = summaryHeaderRow.CreateCell(i);
                        cell.SetCellValue(summaryHeaders[i]);
                        cell.CellStyle = headerStyle;
                    }

                    // Summary data
                    var summaryData = new Dictionary<string, string>
                    {
                        { "Upload Date", history.UploadTimestamp.ToString("yyyy-MM-dd HH:mm:ss") },
                        { "Company ID", history.CompanyId.ToString() },
                        { "Company Name", history.CompanyName },
                        { "Company Code", history.CompanyCode },
                        { "Frequency", history.Frequency },
                        { "Run", history.Run },
                        { "Total Entries", history.TotalEntries.ToString() },
                        { "Successful Entries", history.SuccessfulEntries.ToString() },
                        { "Failed Entries", history.FailedEntries.ToString() },
                        { "Status", history.Success ? "Success" : "Failed" },
                        { "File Name", history.FileName ?? "N/A" }
                    };

                    int rowNum = 1;
                    foreach (var kvp in summaryData)
                    {
                        var row = summarySheet.CreateRow(rowNum++);
                        row.CreateCell(0).SetCellValue(kvp.Key);
                        row.CreateCell(1).SetCellValue(kvp.Value);
                    }

                    for (int i = 0; i < summaryHeaders.Length; i++)
                    {
                        summarySheet.AutoSizeColumn(i);
                    }

                    // Create Details sheet
                    var detailsSheet = workbook.CreateSheet("Entry Details");

                    // Details headers
                    var detailsHeaderRow = detailsSheet.CreateRow(0);
                    var detailsHeaders = new[] {
                        "EmployeeID",
                        "EmployeeName",
                        "Pay Element ID",
                        "Component Code",
                        "Amount",
                        "Status",
                        "Error Message",
                        "Input Type",
                        "Input Quantity"
                    };

                    for (int i = 0; i < detailsHeaders.Length; i++)
                    {
                        var cell = detailsHeaderRow.CreateCell(i);
                        cell.SetCellValue(detailsHeaders[i]);
                        cell.CellStyle = headerStyle;
                    }

                    // Details data
                    rowNum = 1;
                    foreach (var entry in history.Entries)
                    {
                        var row = detailsSheet.CreateRow(rowNum++);

                        row.CreateCell(0).SetCellValue(entry.EmployeeId ?? "");
                        row.CreateCell(1).SetCellValue(entry.EmployeeName ?? "");
                        row.CreateCell(2).SetCellValue(entry.PayElementId ?? "");
                        row.CreateCell(3).SetCellValue(entry.ComponentCode ?? "");

                        var amountCell = row.CreateCell(4);
                        if (entry.Amount.HasValue)
                        {
                            amountCell.SetCellValue((double)entry.Amount.Value);
                        }

                        row.CreateCell(5).SetCellValue(entry.Success ? "Success" : "Failed");
                        row.CreateCell(6).SetCellValue(entry.ErrorMessage ?? "");
                        row.CreateCell(7).SetCellValue(entry.InputType ?? "Amount");

                        var quantityCell = row.CreateCell(8);
                        if (entry.InputQuantity.HasValue)
                        {
                            quantityCell.SetCellValue((double)entry.InputQuantity.Value);
                        }
                    }

                    for (int i = 0; i < detailsHeaders.Length; i++)
                    {
                        detailsSheet.AutoSizeColumn(i);
                    }

                    // Write workbook to memory stream
                    workbook.Write(memoryStream);
                    var byteArray = memoryStream.ToArray();

                    string fileName = $"UploadHistory_{history.CompanyCode}_{history.UploadTimestamp:yyyyMMdd_HHmmss}.xlsx";

                    return File(
                        byteArray,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        fileName
                    );
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error downloading upload history for ID {id}");
                return StatusCode(500, new { error = "Error downloading upload history", details = ex.Message });
            }
        }
    }
}