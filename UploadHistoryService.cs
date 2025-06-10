using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using PaySpaceWaitingEvents.API.Data;
using PaySpaceWaitingEvents.API.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PaySpaceWaitingEvents.API.Services
{
    public interface IUploadHistoryService
    {
        Task<int> SaveUploadHistoryAsync(UploadHistory history);
        Task<bool> SaveEntryAsync(UploadHistoryEntry entry);
        Task<UploadHistory> GetUploadHistoryWithEntriesAsync(int id);
    }

    public class UploadHistoryService : IUploadHistoryService
    {
        private readonly ApplicationDbContext _dbContext;
        private readonly ILogger<UploadHistoryService> _logger;

        public UploadHistoryService(
            ApplicationDbContext dbContext,
            ILogger<UploadHistoryService> logger)
        {
            _dbContext = dbContext;
            _logger = logger;
        }

        public async Task<int> SaveUploadHistoryAsync(UploadHistory history)
        {
            try
            {
                // Validate and clean data
                history.CompanyName = (history.CompanyName ?? "Unknown").Substring(0, Math.Min(100, (history.CompanyName ?? "Unknown").Length));
                history.CompanyCode = (history.CompanyCode ?? "Unknown").Substring(0, Math.Min(50, (history.CompanyCode ?? "Unknown").Length));
                history.Frequency = (history.Frequency ?? "Unknown").Substring(0, Math.Min(50, (history.Frequency ?? "Unknown").Length));
                history.Run = (history.Run ?? "Unknown").Substring(0, Math.Min(50, (history.Run ?? "Unknown").Length));
                history.FileName = (history.FileName ?? "Unknown").Substring(0, Math.Min(255, (history.FileName ?? "Unknown").Length));

                // use existing context
                _dbContext.UploadHistories.Add(history);
                await _dbContext.SaveChangesAsync();

                _logger.LogInformation($"Successfully saved upload history with ID {history.Id}");
                return history.Id;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving upload history");
                throw;
            }
        }

        public async Task<bool> SaveEntryAsync(UploadHistoryEntry entry)
        {
            try
            {
                // Validate and clean data
                entry.EmployeeId = (entry.EmployeeId ?? "Unknown").Substring(0, Math.Min(50, (entry.EmployeeId ?? "Unknown").Length));
                entry.EmployeeName = (entry.EmployeeName ?? "Unknown").Substring(0, Math.Min(100, (entry.EmployeeName ?? "Unknown").Length));
                entry.PayElementId = (entry.PayElementId ?? "Unknown").Substring(0, Math.Min(50, (entry.PayElementId ?? "Unknown").Length));
                entry.ComponentCode = (entry.ComponentCode ?? "").Substring(0, Math.Min(50, (entry.ComponentCode ?? "").Length));
                entry.ErrorMessage = entry.ErrorMessage ?? "";
                entry.InputType = entry.InputType ?? "Amount"; // Default to Amount if not provided

                if (entry.ErrorMessage.Length > 500)
                    entry.ErrorMessage = entry.ErrorMessage.Substring(0, 500);

                // Use standard Entity Framework 
                _dbContext.UploadHistoryEntries.Add(entry);
                await _dbContext.SaveChangesAsync();

                _logger.LogInformation($"Successfully saved entry for employee {entry.EmployeeId} with InputType: {entry.InputType}, InputQuantity: {entry.InputQuantity}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error saving entry for employee {entry.EmployeeId}");
                return false;
            }
        }

        public async Task<UploadHistory> GetUploadHistoryWithEntriesAsync(int id)
        {
            try
            {
                // Use existing context to get history
                var history = await _dbContext.UploadHistories
                    .FirstOrDefaultAsync(h => h.Id == id);

                if (history == null)
                    return null;

                // Fetch entries separately
                var entries = await _dbContext.UploadHistoryEntries
                    .Where(e => e.UploadHistoryId == id)
                    .ToListAsync();

                history.Entries = entries;
                return history;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error retrieving upload history with ID {id}");
                throw;
            }
        }
    }
}