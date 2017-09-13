using System;

namespace GraphSDKDemo.Models
{
    public class Folder
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public int FileCount { get; set; }
        public DateTime Created { get; set; }
        public DateTime LastModified { get; set; }
        public bool Shared { get; set; }
        public string Url { get; set; }
    }
}
