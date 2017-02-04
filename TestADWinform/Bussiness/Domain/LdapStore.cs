using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Bussiness
{
    public static class LdapStore
    {
        const string LOGONNAME = "extooladmin@cn-bj.pactera.com";
        const string LOGONPASSWORD = "4%RDXAsd";

        const string Beijing_Hisoft_com = "Beijing_Hisoft_com";
        const string Cnbj_Pactera_com = "Cnbj_Pactera_com";
        const string Dalian_Hisoft_com = "Dalian_Hisoft_com";
        const string Wuxi_Hisoft_com = "Wuxi_Hisoft_com";
        const string Shanghai_Hisoft_com = "Shanghai_Hisoft_com";
        const string Pactera_com = "Pactera_com";
        const string Cnsh_Pactera_com = "Cnsh_Pactera_com";
        const string Cnsz_Pactera_com = "Cnsz_Pactera_com";
        const string Cnbj_Pactera_com_183 = "Cnbj_Pactera_com_183";

        public static Dictionary<string, Ldap> LdapList { get; private set; }

        static LdapStore()
        {
            Ldap BeijingHisoftCom = new Ldap
            {
                ID = 1,
                DomainName = "beijing.hisoft.com",
                LDAPDomain = "DC=beijing,DC=hisoft,DC=com",
                ADPath = "LDAP://192.168.18.10",
                LogonName = LOGONNAME,
                LogonPassword = LOGONPASSWORD
            };
            

            Ldap CnbjPacteraCom = new Ldap();
            CnbjPacteraCom.ID = 2;
            CnbjPacteraCom.DomainName = "cn-bj.pactera.com";
            CnbjPacteraCom.LDAPDomain = "DC=cn-bj,DC=pactera,DC=com";
            CnbjPacteraCom.ADPath = "LDAP://172.16.254.25";
            CnbjPacteraCom.LogonName = LOGONNAME;
            CnbjPacteraCom.LogonPassword = LOGONPASSWORD;

            Ldap DalianHisoftCom = new Ldap
            {
                ID = 3,
                DomainName = "dalian.hisoft.com",
                LDAPDomain = "DC=cn-bj,DC=pactera,DC=com",
                ADPath = "LDAP://192.168.88.64",
                LogonName = LOGONNAME,
                LogonPassword = LOGONPASSWORD
            };

            Ldap WuxiHisoftCom = new Ldap
            {
                ID = 4,
                DomainName = "wuxi.hisoft.com",
                LDAPDomain = "DC=wuxi,DC=hisoft,DC=com",
                ADPath = "LDAP://192.168.88.76",
                LogonName = LOGONNAME,
                LogonPassword = LOGONPASSWORD
            };

            Ldap ShanghaiHisoftCom = new Ldap
            {
                ID = 5,
                DomainName = "shanghai.hisoft.com",
                LDAPDomain = "DC=shanghai,DC=hisoft,DC=com",
                ADPath = "LDAP://192.168.88.77",
                LogonName = LOGONNAME,
                LogonPassword = LOGONPASSWORD
            };

            Ldap PacteraCom = new Ldap
            {
                ID = 6,
                DomainName = "pactera.com",
                LDAPDomain = "DC=pactera,DC=com",
                ADPath = "LDAP://172.16.254.21",
                LogonName = LOGONNAME,
                LogonPassword = LOGONPASSWORD
            };

            Ldap CnshPacteraCom = new Ldap
            {
                ID = 7,
                DomainName = "cn-sh.pactera.com",
                LDAPDomain = "DC=cn-sh,DC=pactera,DC=com",
                ADPath = "LDAP://172.16.254.28",
                LogonName = LOGONNAME,
                LogonPassword = LOGONPASSWORD
            };

            Ldap CnszPacteraCom = new Ldap();
            CnszPacteraCom.ID = 8;
            CnszPacteraCom.DomainName = "cn-sh.pactera.com";
            CnszPacteraCom.LDAPDomain = "CN-SZ.pactera.com";
            CnszPacteraCom.ADPath = "LDAP://172.18.10.200";
            CnszPacteraCom.LogonName = LOGONNAME;
            CnszPacteraCom.LogonPassword = LOGONPASSWORD;

            Ldap CnbjPacteraCom183 = new Ldap
            {
                ID = 9,
                DomainName = "cn-bj.pactera.com",
                LDAPDomain = "DC=cn-bj,DC=pactera,DC=com",
                ADPath = "LDAP://172.16.254.183",
                LogonName = LOGONNAME,
                LogonPassword = LOGONPASSWORD
            };
            

            LdapList = new Dictionary<string, Ldap>();
            LdapList.Add(Beijing_Hisoft_com, BeijingHisoftCom);
            LdapList.Add(Cnbj_Pactera_com, CnbjPacteraCom);
            LdapList.Add(Dalian_Hisoft_com, DalianHisoftCom);
            LdapList.Add(Wuxi_Hisoft_com, WuxiHisoftCom);
            LdapList.Add(Shanghai_Hisoft_com, ShanghaiHisoftCom);
            LdapList.Add(Pactera_com, PacteraCom);
            LdapList.Add(Cnsh_Pactera_com, CnshPacteraCom);
            LdapList.Add(Cnsz_Pactera_com, CnszPacteraCom);
            LdapList.Add(Cnbj_Pactera_com_183, CnbjPacteraCom183);
        }

        #region 使用静态属性创建

        //public static Dictionary<string, Ldap> LdapList
        //{
        //    get
        //    {
        //        Dictionary<string, Ldap> dic = new Dictionary<string, Ldap>();
        //        dic.Add(Beijing_Hisoft_com, BeijingHisoftCom);
        //        dic.Add(Cnbj_Pactera_com, CnbjPacteraCom);
        //        dic.Add(Dalian_Hisoft_com, DalianHisoftCom);
        //        dic.Add(Wuxi_Hisoft_com, WuxiHisoftCom);
        //        dic.Add(Shanghai_Hisoft_com, ShanghaiHisoftCom);
        //        dic.Add(Pactera_com, PacteraCom);
        //        dic.Add(Cnsh_Pactera_com, CnshPacteraCom);
        //        dic.Add(Cnsz_Pactera_com, CnszPacteraCom);
        //        dic.Add(Cnbj_Pactera_com_183, CnbjPacteraCom183);
        //        return dic;
        //    }
        //    set;
        //}

        //static Ldap BeijingHisoftCom = new Ldap 
        //{ 
        //    ID = 1,
        //    DomainName = "beijing.hisoft.com",
        //    LDAPDomain = "DC=beijing,DC=hisoft,DC=com",
        //    ADPath = "LDAP://192.168.18.10",
        //    LogonName = LOGONNAME,
        //    LogonPassword = LOGONPASSWORD 
        //};

        //static Ldap CnbjPacteraCom = new Ldap
        //{
        //    ID = 2,
        //    DomainName = "cn-bj.pactera.com",
        //    LDAPDomain = "DC=cn-bj,DC=pactera,DC=com",
        //    ADPath = "LDAP://172.16.254.25",
        //    LogonName = LOGONNAME,
        //    LogonPassword = LOGONPASSWORD
        //};

        //static Ldap DalianHisoftCom = new Ldap
        //{
        //    ID = 3,
        //    DomainName = "dalian.hisoft.com",
        //    LDAPDomain = "DC=cn-bj,DC=pactera,DC=com",
        //    ADPath = "LDAP://192.168.88.64",
        //    LogonName = LOGONNAME,
        //    LogonPassword = LOGONPASSWORD
        //};

        //static Ldap WuxiHisoftCom = new Ldap
        //{
        //    ID = 4,
        //    DomainName = "wuxi.hisoft.com",
        //    LDAPDomain = "DC=wuxi,DC=hisoft,DC=com",
        //    ADPath = "LDAP://192.168.88.76",
        //    LogonName = LOGONNAME,
        //    LogonPassword = LOGONPASSWORD
        //};

        //static Ldap ShanghaiHisoftCom = new Ldap
        //{
        //    ID = 5,
        //    DomainName = "shanghai.hisoft.com",
        //    LDAPDomain = "DC=shanghai,DC=hisoft,DC=com",
        //    ADPath = "LDAP://192.168.88.77",
        //    LogonName = LOGONNAME,
        //    LogonPassword = LOGONPASSWORD
        //};

        //static Ldap PacteraCom = new Ldap
        //{
        //    ID = 6,
        //    DomainName = "pactera.com",
        //    LDAPDomain = "DC=pactera,DC=com",
        //    ADPath = "LDAP://172.16.254.21",
        //    LogonName = LOGONNAME,
        //    LogonPassword = LOGONPASSWORD
        //};

        //static Ldap CnshPacteraCom = new Ldap
        //{
        //    ID = 7,
        //    DomainName = "cn-sh.pactera.com",
        //    LDAPDomain = "DC=cn-sh,DC=pactera,DC=com",
        //    ADPath = "LDAP://172.16.254.28",
        //    LogonName = LOGONNAME,
        //    LogonPassword = LOGONPASSWORD
        //};

        //static Ldap CnszPacteraCom = new Ldap
        //{
        //    ID = 8,
        //    DomainName = "cn-sh.pactera.com",
        //    LDAPDomain = "CN-SZ.pactera.com",
        //    ADPath = "LDAP://172.18.10.200",
        //    LogonName = LOGONNAME,
        //    LogonPassword = LOGONPASSWORD
        //};

        
        //static Ldap CnbjPacteraCom183 = new Ldap
        //{
        //    ID = 9,
        //    DomainName = "cn-bj.pactera.com",
        //    LDAPDomain = "DC=cn-bj,DC=pactera,DC=com",
        //    ADPath = "LDAP://172.16.254.183",
        //    LogonName = LOGONNAME,
        //    LogonPassword = LOGONPASSWORD
        //};

        #endregion

    }
}
