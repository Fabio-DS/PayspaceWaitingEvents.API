using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace PaySpaceWaitingEvents.API.Services
{
    public interface IWaitingEventsService
    {
        Task<(List<PayElementEntry> Entries, string LogicalIdPrefix)> ProcessWaitingEventsFile(Stream fileStream);
    }
}