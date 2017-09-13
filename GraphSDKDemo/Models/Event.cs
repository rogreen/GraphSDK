using System;

namespace GraphSDKDemo.Models
{
    public class Event
    {
        public string Id { get; set; }
        public string Subject { get; set; }
        public string Location { get; set; }
        public string Organizer { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public bool IsAllDay { get; set; }
        public string Categories { get; set; }
    }


}
