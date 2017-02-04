using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bussiness.Domain
{
    public static class LdapStore2
    {
        const string LOGONNAME = @"extest\Administrator";
        const string LOGONPASSWORD = "123.com";

        const string Testdc_Extest_com = "Testdc_Extest_com";

        public static Dictionary<string, Ldap> LdapList { get; set; }

        static LdapStore2()
        {
            Ldap TestdcExtestCom = new Ldap();
            TestdcExtestCom.ID = 1;
            TestdcExtestCom.DomainName = "testdc.extest.com";
            TestdcExtestCom.LDAPDomain = "DC=extest,DC=com";
            TestdcExtestCom.ADPath = "LDAP://172.16.253.234";
            TestdcExtestCom.LogonName = LOGONNAME;
            TestdcExtestCom.LogonPassword = LOGONPASSWORD;

            LdapList = new Dictionary<string, Ldap>();
            LdapList.Add(Testdc_Extest_com, TestdcExtestCom);
        }

        #region 使用静态属性创建

        //static Ldap TestdcExtestCom = new Ldap
        //{
        //    ID = 1,
        //    DomainName = "testdc.extest.com",
        //    LDAPDomain = "DC=extest,DC=com",
        //    ADPath = "LDAP://172.16.253.234",
        //    LogonName = LOGONNAME,
        //    LogonPassword = LOGONPASSWORD
        //};

        #endregion

    }
}
