using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using PaySpaceWaitingEvents.API.Data;
using PaySpaceWaitingEvents.API.Models;
using PaySpaceWaitingEvents.API.Models.PaySpaceWaitingEvents.API.Models;

namespace PaySpaceWaitingEvents.API.Services
{

    public class PayElementMappingService : IPayElementMappingService
    {
        private readonly ApplicationDbContext _context;
        private readonly ILogger<PayElementMappingService> _logger;

        public PayElementMappingService(
            ApplicationDbContext context,
            ILogger<PayElementMappingService> logger)
        {
            _context = context;
            _logger = logger;
        }

        public async Task<List<PayElementMapping>> GetMappingsForLegalEntity(int legalEntityId)
        {
            return await _context.PayElementMappings
                .Where(m => m.LegalEntityId == legalEntityId && m.IsActive)
                .Include(m => m.LegalEntity)
                .ToListAsync();
        }

        public async Task<PayElementMapping> CreateMapping(PayElementMapping mapping)
        {
            mapping.CreatedDate = DateTime.UtcNow;
            mapping.IsActive = true;

            await _context.PayElementMappings.AddAsync(mapping);
            await _context.SaveChangesAsync();

            return mapping;
        }

        public async Task<PayElementMapping> UpdateMapping(int id, PayElementMapping mapping)
        {
            var existingMapping = await _context.PayElementMappings
                .FindAsync(id);

            if (existingMapping == null)
                throw new MappingException($"Mapping with id {id} not found");

            existingMapping.PayElementId = mapping.PayElementId;
            existingMapping.ComponentCode = mapping.ComponentCode;
            existingMapping.Frequency = mapping.Frequency;
            existingMapping.Description = mapping.Description;
            existingMapping.IsActive = mapping.IsActive;
            existingMapping.LastModifiedDate = DateTime.UtcNow;

            await _context.SaveChangesAsync();

            return existingMapping;
        }

        public async Task DeleteMapping(int id)
        {
            var mapping = await _context.PayElementMappings
                .FindAsync(id);

            if (mapping == null)
                throw new MappingException($"Mapping with id {id} not found");

            _context.PayElementMappings.Remove(mapping);
            await _context.SaveChangesAsync();
        }

        public async Task<bool> ValidateMapping(PayElementMapping mapping)
        {
            var legalEntity = await _context.LegalEntities.FindAsync(mapping.LegalEntityId);
            return legalEntity != null;
        }

        public async Task<List<PayElementEntry>> MapPayElementsForEntity(List<PayElementEntry> entries, int legalEntityId)
        {
            var entityMappings = await _context.PayElementMappings
                .Where(m => m.LegalEntityId == legalEntityId && m.IsActive)
                .ToDictionaryAsync(m => m.PayElementId, m => m.ComponentCode);

            var mappedEntries = new List<PayElementEntry>();
            var unmappedElements = new HashSet<string>();

            foreach (var entry in entries)
            {
                if (entry.Event == "Pay Element" && !string.IsNullOrEmpty(entry.PayElementId))
                {
                    if (entityMappings.TryGetValue(entry.PayElementId, out string componentCode))
                    {
                        entry.PaySpaceCompCode = componentCode;
                        mappedEntries.Add(entry);
                    }
                    else
                    {
                        unmappedElements.Add(entry.PayElementId);
                        _logger.LogWarning($"No mapping found for Pay Element ID {entry.PayElementId} in legal entity {legalEntityId}");
                    }
                }
                else
                {
                    mappedEntries.Add(entry);
                }
            }

            if (unmappedElements.Any())
            {
                throw new MappingException($"Unmapped pay elements found for legal entity {legalEntityId}: {string.Join(", ", unmappedElements)}");
            }

            return mappedEntries;
        }
    }
}