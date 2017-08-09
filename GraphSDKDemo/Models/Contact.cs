using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GraphSDKDemo.Models
{
    public class Contact
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string EmailAddress { get; set; }
        public string CompanyName { get; set; }
        public string JobTitle { get; set; }
        public string HomePhone { get; set; }
        public string BusinessPhone { get; set; }
        public string MobilePhone { get; set; }
        public string PersonalNotes { get; set; }
    }
}
