namespace PaySpaceWaitingEvents.API.Models
{
    public class PayElementMapping
    {
        public int Id { get; set; }
        public int LegalEntityId { get; set; }

        public string PayElementId { get; set; }
        public string ComponentCode { get; set; }
        public string Frequency { get; set; }
        public string Description { get; set; }
        public bool IsActive { get; set; }
        public DateTime CreatedDate { get; set; }
        public DateTime? LastModifiedDate { get; set; }

        public string Category { get; set; }
        public string ApiEndpoint { get; set; }

        public virtual LegalEntity LegalEntity { get; set; }
    }
}
