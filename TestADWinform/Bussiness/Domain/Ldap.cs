using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bussiness
{
    public struct Ldap
    {
        public int ID { get; set; }
        public string DomainName { get; set; }
        public string LDAPDomain { get; set; }
        public string ADPath { get; set; }
        public string LogonName { get; set; }
        public string LogonPassword { get; set; }
    }
}
