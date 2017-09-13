using System;

namespace GraphSDKDemo.Models
{
    public class File
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public long Size { get; set; }
        public DateTime Created { get; set; }
        public DateTime LastModified { get; set; }
        public bool Shared { get; set; }
        public string Url { get; set; }
    }
}
