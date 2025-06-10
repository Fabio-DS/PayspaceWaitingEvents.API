namespace PaySpaceWaitingEvents.API.Models
{
    namespace PaySpaceWaitingEvents.API.Models
    {
        public class MappingException : Exception
        {
            public MappingException() { }
            public MappingException(string message) : base(message) { }
            public MappingException(string message, Exception inner) : base(message, inner) { }
        }
    }
}
