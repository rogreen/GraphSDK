namespace GraphSDKDemo.Models
{
    public class Message
    {
        public string Id { get; set; }
        public string From { get; set; }
        public string Sender { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string ReceivedDateTime { get; set; }
        public string Importance { get; set; }
        public bool IsRead { get; set; }
        public bool HasAttachments { get; set; }
    }
}
