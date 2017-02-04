using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bussiness
{
    public class LDAP
    {
        public int ID { get; set; }
        public string DomainName { get; set; }
        public string LDAPDomain { get; set; }
        public string ADPath { get; set; }
        public string ADUser { get; set; }
        public string ADPassword { get; set; }
    }
}
