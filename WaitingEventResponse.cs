namespace PaySpaceWaitingEvents.API.Models
{
    public class WaitingEventResponse
    {
        public int TotalEntriesFound { get; set; }
        public int ValidEntriesCount { get; set; }
        public List<PayElementEntry> ValidEntries { get; set; }
        public string Message { get; set; }
    }
}
