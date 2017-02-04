using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Bussiness
{
    public class BaseBusiness
    {
        public string Conn { get; set; }
        public string ADConn { get; set; }
        public List<LDAP> LDAPStore { get; set; }
        public List<LDAPLocationMapping> LDAPLocationMappingStore { get; set; }
        public List<LocationMapping> LocationMappingStore { get; set; }

        public BaseBusiness()
        {
            Conn = ConfigurationManager.AppSettings["Conn"];
            ADConn = ConfigurationManager.ConnectionStrings["ADDBConnection"].ConnectionString;
            LDAPStore = GetLADPStore(ConfigurationManager.AppSettings["LDAPPath"]);
            LDAPLocationMappingStore = GetLDAPLocationMappingStore(ConfigurationManager.AppSettings["LDAPLocationMappingPath"]);
        }

        /// <summary>
        /// 从XML获取LDAP数据
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static List<LDAP> GetLADPStore(string path)
        {
            XElement xe = XElement.Load(path);
            IEnumerable<XElement> elements = from ele in xe.Elements("LDAP")
                                             select ele;
            List<LDAP> ldaps = new List<LDAP>();
            foreach (var ele in elements)
            {
                LDAP ldap = new LDAP();
                var id = ele.Element("ID");
                if (id != null)
                    ldap.ID = Convert.ToInt32(id.Value);
                var ladpDomain = ele.Element("LDAPDomain");
                if (ladpDomain != null)
                    ldap.LDAPDomain = ladpDomain.Value;
                var domainName = ele.Element("DomainName");
                if (domainName != null)
                    ldap.DomainName = domainName.Value;
                var adPath = ele.Element("ADPath");
                if (adPath != null)
                    ldap.ADPath = adPath.Value;
                var adUser = ele.Element("ADUser");
                if (adUser != null)
                    ldap.ADUser = adUser.Value;
                var adPassword = ele.Element("ADPassword");
                if (adPassword != null)
                    ldap.ADPassword = adPassword.Value;
                ldaps.Add(ldap);
            }
            return ldaps;
        }

        /// <summary>
        /// 从XML获取LDAPLocationMapping数据
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static List<LDAPLocationMapping> GetLDAPLocationMappingStore(string path)
        {
            XElement xe = XElement.Load(path);
            IEnumerable<XElement> elements = from ele in xe.Elements("Mapping")
                                             select ele;
            List<LDAPLocationMapping> mappings = new List<LDAPLocationMapping>();
            foreach (var ele in elements)
            {
                LDAPLocationMapping mapping = new LDAPLocationMapping();
                var ldapID = ele.Element("LDAPID");
                if (ldapID != null)
                    mapping.LDAPID = Convert.ToInt32(ldapID.Value);
                var ou = ele.Element("OU");
                if (ou != null)
                    mapping.OU = ou.Value;
                var location = ele.Element("Location");
                if (location != null)
                    mapping.Location = location.Value;
                mappings.Add(mapping);
            }
            return mappings;
        }

        public bool IsADUserExists(string empno)
        {
            //存在则登录当前AD Server
            try
            {
                foreach (var ldap in LDAPStore)
                {
                    ADHelper.DomainName = ldap.DomainName;
                    ADHelper.LDAPDomain = ldap.LDAPDomain; //ADHelper.DomainName = ldap.LDAPDomain;
                    ADHelper.ADPath = ldap.ADPath;
                    ADHelper.ADUser = ldap.ADUser;
                    ADHelper.ADPassword = ldap.ADPassword;
                    if (ADHelper.IsAccExists(empno))
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        public bool IsADUserExistsByUPN(string upn)
        {
            //存在则登录当前AD Server
            try
            {
                foreach (var ldap in LDAPStore)
                {
                    ADHelper.DomainName = ldap.DomainName;
                    ADHelper.LDAPDomain = ldap.LDAPDomain; //ADHelper.DomainName = ldap.LDAPDomain;
                    ADHelper.ADPath = ldap.ADPath;
                    ADHelper.ADUser = ldap.ADUser;
                    ADHelper.ADPassword = ldap.ADPassword;
                    if (ADHelper.IsAccExists2(upn))
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        public bool IsADUserExistsByDisplayName(string displayName)
        {
            bool isExist = false;
            //存在则登录当前AD Server
            foreach (var ldap in LDAPStore)
            {
                ADHelper.DomainName = ldap.DomainName;
                ADHelper.LDAPDomain = ldap.LDAPDomain; //ADHelper.DomainName = ldap.LDAPDomain;
                ADHelper.ADPath = ldap.ADPath;
                ADHelper.ADUser = ldap.ADUser;
                ADHelper.ADPassword = ldap.ADPassword;

                DirectoryEntry de = ADHelper.GetDirectoryObject();
                DirectorySearcher deSearch = new DirectorySearcher(de);
                deSearch.Filter = "(&(&(objectCategory=person)(objectClass=user))(displayname=" + displayName + "))";       // LDAP 查询串
                SearchResultCollection results = deSearch.FindAll();
                if (results.Count > 0)
                {
                    isExist = true;
                    break;
                }
            }
            return isExist;
        }
        /// <summary>
        /// AD中是否存在该用户
        /// 0:不存在该账户和公共名
        /// 1:存在该账户
        /// 2:存在公共名
        /// </summary>
        /// <param name="empno">例如：P0023143</param>
        /// <param name="commonName">例如：San Zhang</param>
        /// <returns></returns>
        public int IsADUserExists(string empno, string commonName)
        {
            //存在则登录当前AD Server
            try
            {
                foreach (var ldap in LDAPStore)
                {
                    ADHelper.DomainName = ldap.DomainName;
                    ADHelper.LDAPDomain = ldap.LDAPDomain; //ADHelper.DomainName = ldap.LDAPDomain;
                    ADHelper.ADPath = ldap.ADPath;
                    ADHelper.ADUser = ldap.ADUser;
                    ADHelper.ADPassword = ldap.ADPassword;
                    if (ADHelper.IsAccExists(empno))
                    {
                        return 1;
                    }
                    if (ADHelper.IsUserExists(commonName))
                    {
                        return 2;
                    }
                }
                return 0;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        protected void LogonPrimaryAD()
        {
            //登录AD
            var ldap = LDAPStore.FirstOrDefault(l => l.ID == Convert.ToInt32(ConfigurationManager.AppSettings["GroupLDAPID"]));
            if (ldap != null)
            {
                ADHelper.DomainName = ldap.DomainName;
                ADHelper.LDAPDomain = ldap.LDAPDomain;
                ADHelper.ADPath = ldap.ADPath;
                ADHelper.ADUser = ldap.ADUser;
                ADHelper.ADPassword = ldap.ADPassword;
            }
        }
        protected void LogonDomainComputer(string domain)
        {
            if (LDAPStore == null)
                LDAPStore = new List<LDAP>();

            LDAP ldap = LDAPStore.FirstOrDefault(e => e.DomainName.ToLower() == domain.ToLower());
            if (ldap == null)
                ldap = LDAPStore.FirstOrDefault(l => l.ID == Convert.ToInt32(ConfigurationManager.AppSettings["GroupLDAPID"]));

            ADHelper.DomainName = ldap.DomainName;
            ADHelper.LDAPDomain = ldap.LDAPDomain;
            ADHelper.ADPath = ldap.ADPath;
            ADHelper.ADUser = ldap.ADUser;
            ADHelper.ADPassword = ldap.ADPassword;
        }

    }
}
