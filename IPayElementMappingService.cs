using PaySpaceWaitingEvents.API.Models;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace PaySpaceWaitingEvents.API.Services
{
    public interface IPayElementMappingService
    {
        Task<List<PayElementMapping>> GetMappingsForLegalEntity(int legalEntityId);
        Task<PayElementMapping> CreateMapping(PayElementMapping mapping);
        Task<PayElementMapping> UpdateMapping(int id, PayElementMapping mapping);
        Task DeleteMapping(int id);

    }
}