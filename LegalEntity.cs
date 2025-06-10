using System.ComponentModel.DataAnnotations;

namespace PaySpaceWaitingEvents.API.Models
{
    public class LegalEntity
    {
        public int Id { get; set; }
        public string CompanyCode { get; set; }
        public string CompanyName { get; set; }
        public bool IsActive { get; set; }

        [MaxLength(20)]
        public string LogicalIdPrefix { get; set; }

        public virtual ICollection<PayElementMapping> PayElementMappings { get; set; }
    }
}
