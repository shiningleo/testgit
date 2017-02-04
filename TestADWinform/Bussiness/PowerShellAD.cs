using ADModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.DirectoryServices;
using System.Linq;
using System.Text;

namespace Bussiness
{
    public class PowerShellAD : BaseBusiness
    {
        static readonly Object lockthis = new object();
        public ExecResult PostADUserInfo(List<ADUser> adUserList)
        {
            ExecResult result = ExecResult.Failure;
            if (adUserList != null)
            {
                foreach (ADUser adUser in adUserList)
                {
                    result = PostADUserInfo(adUser);
                }
                if (result == (int)ExecResult.Failure)
                {
                    foreach (ADUser adUser in adUserList)
                    {
                        DeleteADUserInfo(adUser.SamAccountName);
                    }
                }
            }
            return result;
        }
        public ExecResult PostADUserInfo(ADUser adUser)
        {
            ExecResult result = ExecResult.Failure;
            try
            {
                if (IsADUserExists(adUser.SamAccountName))
                {
                    LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, adUser.SamAccountName + "已存在。"));
                    return ExecResult.Existing;
                }
                if (string.IsNullOrEmpty(adUser.FirstName))
                    return ExecResult.FirstNameEmpty;
                if (string.IsNullOrEmpty(adUser.LastName))
                    return ExecResult.LastNameEmpty;

                //补全用户信息
                result = CompleteADUserInfo2(adUser);
                if (result == ExecResult.Success)
                {
                    lock (lockthis)
                    {
                        //插入用户信息
                        result = InsertADUserInfoToDB(adUser);
                        if (result == ExecResult.Success)
                        {
                            result = CreateADUser(adUser);  //创建ADUser
                            if (result == ExecResult.Success)
                                UpdateADUserInfo(adUser);
                        }
                    }
                    #region old
                    ////插入用户信息
                    //result = InsertADUserInfoToDB(adUser);
                    //////如果是正式员工，创建AD账号
                    ////if (result == ExecResult.Success)
                    ////{
                    ////    //LogonPrimaryAD(); //登录到主域控
                    ////    result = CreateADUser(adUser);  //创建ADUser
                    ////    if (result == ExecResult.Success)
                    ////        UpdateADUserInfo(adUser);
                    ////}
                    //System.Threading.Thread.Sleep(1000);
                    //List<ADUser> adUserList = ReadADUserInfoFromDB();
                    //foreach (ADUser user in adUserList)
                    //{
                    //    result = CreateADUser(adUser);  //创建ADUser
                    //    if (result == ExecResult.Success)
                    //        UpdateADUserInfo(adUser);
                    //}
                    #endregion
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "CreateADUser," + adUser.SamAccountName + "," + ex.Message));
            }

            LogHelper.WriteLog(new LogModel(Level.Info, DateTime.Now, "Create AD Account:" + "SamAccountName: " + adUser.SamAccountName + ", FirstName:" + adUser.FirstName + ", LastName:" + adUser.LastName + ", Result:" + result.ToString()));
            return result;
        }

        public ExecResult CreateADAccountTest(ADUser adUser)
        {
            return CreateADUser(adUser); 
        }
        public List<EmailInfo> GetEmailInfo(List<string> employeeIds)
        {
            List<EmailInfo> emailInfoList = new List<EmailInfo>();
            try
            {
                emailInfoList = GetEmailInfosByEmployeeId(employeeIds);
            }
            catch(Exception ex)
            {
                LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "GetEmailInfosByEmployeeId：" + ex.Message));
            }
            return emailInfoList;
        }
        public EmailInfo GetEmailInfo(string employeeId)
        {
            try
            {
                EmailInfo emailInfo = GetEmailInfoByEmployeeId(employeeId);
                return emailInfo;
            }
            catch
            {
                return null;
            }
        }
        public string GetEmailProperty(string sAmAccountName)
        {
            LogonPrimaryAD();
            DirectoryEntry de = ADHelper.GetDirectoryObject();
            DirectorySearcher deSearch = new DirectorySearcher(de);
            deSearch.Filter = "(&(&(objectCategory=person)(objectClass=user))(sAMAccountName=" +
                sAmAccountName + "))";       // LDAP 查询串
            SearchResult result = deSearch.FindOne();
            return ADHelper.GetProperty(result, "mail");
        }
        public void Execute_minus1to0()
        {
            string sql = "select * from T_ADUserInfo where RFlag = -1";
            System.Data.SqlClient.SqlDataReader reader = SqlHelper.ExecuteReader(ADConn, CommandType.Text, sql);

            List<ADUser> adUserList = new List<ADUser>();
            while (reader.Read())
            {
                ADUser adUser = new ADUser();
                adUser.SamAccountName = DBNullToString(reader["SamAccountName"]);
                adUser.CommonName = DBNullToString(reader["CommonName"]);
                adUser.Password = DBNullToString(reader["Password"]);
                adUser.UserPrincipalName = DBNullToString(reader["UserPrincipalName"]);
                adUser.FirstName = DBNullToString(reader["FirstName"]);
                adUser.MiddleName = DBNullToString(reader["MiddleName"]);
                adUser.LastName = DBNullToString(reader["LastName"]);
                adUser.GivenName = DBNullToString(reader["GivenName"]);
                adUser.SN = DBNullToString(reader["SN"]);
                adUser.DisplayName = DBNullToString(reader["DisplayName"]);
                adUser.Company = DBNullToString(reader["Company"]);
                adUser.DepartmentNo = DBNullToString(reader["DepartmentNo"]);
                adUser.DisplayName = DBNullToString(reader["DisplayName"]);
                adUser.RootDepartmentNo = DBNullToString(reader["RootDepartmentNo"]);
                adUser.RootDepartmentName = DBNullToString(reader["RootDepartmentName"]);
                adUser.Country = DBNullToString(reader["Country"]);
                adUser.State = DBNullToString(reader["State"]);
                adUser.City = DBNullToString(reader["City"]);
                adUser.Location = DBNullToString(reader["Location"]);
                adUser.OfficeLocation = DBNullToString(reader["OfficeLocation"]);
                adUser.StreetAddress = DBNullToString(reader["StreetAddress"]);
                adUser.Domain = DBNullToString(reader["Domain"]);
                adUser.OU = DBNullToString(reader["OU"]);
                adUser.Grade = DBNullToString(reader["Grade"]);
                adUser.Type = DBNullToString(reader["Type"]);
                adUser.PositionName = DBNullToString(reader["PositionName"]);
                adUser.ManagerNumber = DBNullToString(reader["ManagerNumber"]);
                adUser.ManagerName = DBNullToString(reader["ManagerName"]);
                adUser.PhoneNumber = DBNullToString(reader["PhoneNumber"]);
                adUser.MobileNumber = DBNullToString(reader["MobileNumber"]);
                adUserList.Add(adUser);
            }

            foreach (ADUser adUser in adUserList)
            {
                //LogonDomainComputer(adUser.Domain); //登录到相应的域控服务器
                //LogonPrimaryAD(); //登录到主域控
                ExecResult result = CreateADUser(adUser);  //创建ADUser
                if (result == ExecResult.Success)
                    UpdateADUserInfo(adUser);
            }
        }
        public void Execute_minus1to0(string samAccountName)
        {
            string sql = "select * from T_ADUserInfo where SamAccountName = N'" + samAccountName + "' and RFlag = -1";
            System.Data.SqlClient.SqlDataReader reader = SqlHelper.ExecuteReader(ADConn, CommandType.Text, sql);
            ADUser adUser = new ADUser();
            if (reader.Read())
            {
                adUser.SamAccountName = samAccountName;
                adUser.CommonName = DBNullToString(reader["CommonName"]);
                adUser.Password = DBNullToString(reader["Password"]);
                adUser.UserPrincipalName = DBNullToString(reader["UserPrincipalName"]);
                adUser.FirstName = DBNullToString(reader["FirstName"]);
                adUser.MiddleName = DBNullToString(reader["MiddleName"]);
                adUser.LastName = DBNullToString(reader["LastName"]);
                adUser.GivenName = DBNullToString(reader["GivenName"]);
                adUser.SN = DBNullToString(reader["SN"]);
                adUser.DisplayName = DBNullToString(reader["DisplayName"]);
                adUser.Company = DBNullToString(reader["Company"]);
                adUser.DepartmentNo = DBNullToString(reader["DepartmentNo"]);
                adUser.DisplayName = DBNullToString(reader["DisplayName"]);
                adUser.RootDepartmentNo = DBNullToString(reader["RootDepartmentNo"]);
                adUser.RootDepartmentName = DBNullToString(reader["RootDepartmentName"]);
                adUser.Country = DBNullToString(reader["Country"]);
                adUser.State = DBNullToString(reader["State"]);
                adUser.City = DBNullToString(reader["City"]);
                adUser.Location = DBNullToString(reader["Location"]);
                adUser.OfficeLocation = DBNullToString(reader["OfficeLocation"]);
                adUser.StreetAddress = DBNullToString(reader["StreetAddress"]);
                adUser.Domain = DBNullToString(reader["Domain"]);
                adUser.OU = DBNullToString(reader["OU"]);
                adUser.Grade = DBNullToString(reader["Grade"]);
                adUser.Type = DBNullToString(reader["Type"]);
                adUser.PositionName = DBNullToString(reader["PositionName"]);
                adUser.ManagerNumber = DBNullToString(reader["ManagerNumber"]);
                adUser.ManagerName = DBNullToString(reader["ManagerName"]);
                adUser.PhoneNumber = DBNullToString(reader["PhoneNumber"]);
                adUser.MobileNumber = DBNullToString(reader["MobileNumber"]);
            }

            //LogonDomainComputer(adUser.Domain); //登录到相应的域控服务器
            //LogonPrimaryAD(); //登录到主域控
            ExecResult result = CreateADUser(adUser);  //创建ADUser
            if (result == ExecResult.Success)
                UpdateADUserInfo(adUser);

            ////如果是正式员工，创建AD账号
            //if (IsRegularEmployee(adUser.Grade))
            //{
                
            //}
        }
        private string DBNullToString(object obj)
        {
            if (obj == DBNull.Value || obj == null)
                return string.Empty;
            else
                return obj.ToString();
        }

        /// <summary>
        /// 获取完整的用户信息
        /// </summary>
        /// <param name="adUser"></param>
        /// <param name="fromAD">是否从AD账号获取</param>
        protected void CompleteADUserInfo(ADUser adUser, bool fromAD = false)
        {
            adUser.SamAccountName = adUser.SamAccountName ?? "";
            adUser.FirstName = adUser.FirstName ?? "";
            adUser.MiddleName = adUser.MiddleName ?? "";
            adUser.LastName = adUser.LastName ?? "";
            adUser.Company = adUser.Company ?? "";
            adUser.DepartmentNo = adUser.DepartmentNo ?? "";
            adUser.Department = adUser.Department ?? "";
            adUser.RootDepartmentNo = adUser.RootDepartmentNo ?? "";
            adUser.RootDepartmentName = adUser.RootDepartmentName ?? "";
            adUser.Country = adUser.Country ?? "";
            adUser.State = adUser.State ?? "";
            adUser.City = adUser.City ?? "";
            adUser.Location = adUser.Location ?? "";
            adUser.OfficeLocation = adUser.OfficeLocation ?? "";
            adUser.StreetAddress = adUser.StreetAddress ?? "";
            adUser.Grade = adUser.Grade ?? "";
            adUser.ManagerNumber = adUser.ManagerNumber ?? "";
            adUser.ManagerName = adUser.ManagerName ?? "";
            adUser.PhoneNumber = adUser.PhoneNumber ?? "";
            adUser.MobileNumber = adUser.MobileNumber ?? "";

            DealwithRepeateCommonName(adUser);
            adUser.UserPrincipalName = CreateUserPrincipalName(adUser);
            adUser.DisplayName = GetDisplayName(adUser.FirstName, adUser.LastName, adUser.Department, adUser.Location);

            adUser.Domain = GetDomainByLocation(adUser.Location);
            adUser.OU = GetOUByLocation(adUser.Location);
            //adUser.Domain = string.IsNullOrEmpty(adUser.Domain) ? ADHelper.DomainName : adUser.Domain;
            //adUser.OU = GetPacteraUsersOU(adUser.Domain); //GetServerAndOUByEmployee(adUser); //因为用户都会创建到172.16.254.183服务器上，而不是创建到对应的子域上。

            adUser.Type = GetEmployeeType(adUser.Grade);
            int passwordLen = 8;
            if (adUser.Type.ToUpper() == "E3")
                passwordLen = 10;
            adUser.Password = Utilities.GetPassword(passwordLen);
            int loop = 0;
            bool isStrong = Utilities.GetPasswordStrong(adUser.Password) < 4;
            while (isStrong)
            {
                if (loop > 4)
                {
                    adUser.Password = "Pactera@2015";
                    isStrong = true;
                }
                else
                {
                    adUser.Password = Utilities.GetPassword(passwordLen);
                    isStrong = Utilities.GetPasswordStrong(adUser.Password) < 4;
                }
                loop++;
            }
            adUser.PositionName = GetEmployeePositionName(adUser.SamAccountName);

            //从数据库获取其他信息
            string strSql = @"select * from V_Employee_Info where EMPLOYEE_NUMBER = '" + adUser.SamAccountName + "'";
            System.Data.SqlClient.SqlDataReader reader = SqlHelper.ExecuteReader(Conn, CommandType.Text, strSql);
            if (reader.Read())
            {
                adUser.Country = (reader["COUNTRY"] == null || reader["COUNTRY"] == DBNull.Value) ? "" : reader["COUNTRY"].ToString();
                adUser.StreetAddress = (reader["OFFICE_LOCATION"] == null || reader["OFFICE_LOCATION"] == DBNull.Value) ? "" : reader["OFFICE_LOCATION"].ToString();
                adUser.City = (reader["HWS_HUKOU_CITY"] == null || reader["HWS_HUKOU_CITY"] == DBNull.Value) ? "" : reader["HWS_HUKOU_CITY"].ToString();
            }
            if (string.IsNullOrEmpty(adUser.Country) || adUser.Country.ToUpper() == "CHN")
                adUser.Country = "CN";
        }
        protected ExecResult CompleteADUserInfo2(ADUser adUser)
        {
            ExecResult result = ExecResult.Success;

            result = CommonNameUPNDisplayName(adUser);
            if (result == ExecResult.Success)
            {
                adUser.Domain = GetDomainByLocation(adUser.Location);
                adUser.OU = GetOUByLocation(adUser.Location);

                adUser.Type = GetEmployeeType(adUser.Grade);
                int passwordLen = 8;
                if (adUser.Type.ToUpper() == "E4")
                    passwordLen = 10;
                adUser.Password = Utilities.GetPassword(passwordLen);
                int loop = 0;
                bool isStrong = Utilities.GetPasswordStrong(adUser.Password) < 4;
                while (isStrong)
                {
                    if (loop > 4)
                    {
                        adUser.Password = "Pactera@2015";
                        isStrong = true;
                    }
                    else
                    {
                        adUser.Password = Utilities.GetPassword(passwordLen);
                        isStrong = Utilities.GetPasswordStrong(adUser.Password) < 4;
                    }
                    loop++;
                }
                adUser.PositionName = GetEmployeePositionName(adUser.SamAccountName);

                //从数据库获取其他信息
                string strSql = @"select * from V_Employee_Info where EMPLOYEE_NUMBER = '" + adUser.SamAccountName + "'";
                System.Data.SqlClient.SqlDataReader reader = SqlHelper.ExecuteReader(Conn, CommandType.Text, strSql);
                if (reader.Read())
                {
                    adUser.Country = (reader["COUNTRY"] == null || reader["COUNTRY"] == DBNull.Value) ? "" : reader["COUNTRY"].ToString();
                    adUser.StreetAddress = (reader["OFFICE_LOCATION"] == null || reader["OFFICE_LOCATION"] == DBNull.Value) ? "" : reader["OFFICE_LOCATION"].ToString();
                    adUser.City = (reader["HWS_HUKOU_CITY"] == null || reader["HWS_HUKOU_CITY"] == DBNull.Value) ? "" : reader["HWS_HUKOU_CITY"].ToString();
                }
                if (string.IsNullOrEmpty(adUser.Country) || adUser.Country.ToUpper() == "CHN")
                    adUser.Country = "CN";
            }

            return result;
        }

        protected string GetCommonName(string firstName, string lastName)
        {
            return firstName + "." + lastName;
        }
        /// <summary>
        /// 拼凑DisplayName
        /// </summary>
        /// <param name="firstName"></param>
        /// <param name="lastName"></param>
        /// <returns></returns>
        protected string GetDisplayName(string firstName, string lastName, string department, string workLocation)
        {
            string displayName = string.Empty;
            if (string.IsNullOrEmpty(firstName) && string.IsNullOrEmpty(lastName))
                return displayName;

            if (string.IsNullOrEmpty(department)) department = string.Empty;
            string[] departmentNames = department.Split(new char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);

            string region = GetRegionByLocation(workLocation);
            displayName = ToTitleCase(lastName.Trim()) + "," + ToTitleCase(firstName.Trim());
            
            bool isexist = IsADUserExistsByDisplayName(displayName);
            int departmenti = 0;
            int i = 1;
            while (isexist)
            {
                if (departmentNames.Length > departmenti)
                {
                    displayName += "(" + departmentNames[departmenti].Trim() + "/" + region + ")";
                    isexist = IsADUserExistsByDisplayName(displayName);
                }
                else
                {
                    string tempName = displayName + i.ToString();
                    isexist = IsADUserExistsByDisplayName(tempName);
                    if (!isexist)
                    {
                        displayName = tempName;
                    }
                    else
                    {
                        i++;
                    }
                }
            }
            
            //if (isexist)
            //{
            //    displayName += "(" + department + "/" + region + ")";
            //    isexist = IsADUserExistsByDisplayName(displayName);
            //    int i = 1;
            //    string tempName = string.Empty;
            //    while (isexist)
            //    {
            //        tempName = displayName;
            //        tempName += i.ToString();
            //        isexist = IsADUserExistsByDisplayName(tempName);
            //        i++;
            //    }
            //    if (!string.IsNullOrEmpty(tempName))
            //        displayName = tempName;
            //}

            return displayName;
        }
        protected string GetDomainByLocation(string location)
        {
            if (string.IsNullOrEmpty(location))
                return string.Empty;

            if (LocationMappingStore == null)
                LocationMappingStore = GetLocationMappingStore();

            LocationMapping entity = LocationMappingStore.FirstOrDefault(s => s.WorkLocation.ToUpper() == location.ToUpper());
            if (entity == null)
                return string.Empty;
            else
                return entity.Domain;
        }
        /// <summary>
        /// 根据工作地，获取OU
        /// </summary>
        /// <param name="location">BEIJING</param>
        /// <returns>OU=Others,OU=Office365Users,DC=pactera,DC=com</returns>
        protected string GetOUByLocation(string location)
        {
            if (string.IsNullOrEmpty(location))
                return string.Empty;

            if (LocationMappingStore == null)
                LocationMappingStore = GetLocationMappingStore();

            LocationMapping entity = LocationMappingStore.FirstOrDefault(s => s.WorkLocation.ToUpper() == location.ToUpper());
            if (entity == null)
                return string.Empty;
            else
                return entity.OU;
        }
        /// <summary>
        /// 根据工作地，获取地区简写
        /// </summary>
        /// <param name="location"></param>
        /// <returns></returns>
        protected string GetRegionByLocation(string location)
        {
            if (string.IsNullOrEmpty(location))
                return string.Empty;

            if (LocationMappingStore == null)
                LocationMappingStore = GetLocationMappingStore();

            LocationMapping entity = LocationMappingStore.FirstOrDefault(s => s.WorkLocation.ToUpper() == location.ToUpper());
            if (entity == null)
                return string.Empty;
            else
                return entity.Region;
        }

        /// <summary>
        /// 获取OU名称（Office365Users/Others）
        /// </summary>
        /// <param name="ou">OU=Others,OU=Office365Users,DC=pactera,DC=com</param>
        /// <returns></returns>
        protected string GetLocateOUName(string ou, string adUser, string adPassword, string adPath, string ldapDomain)
        {
            List<string> ous = new List<string>();
            string[] ouArr = ou.Split(new char[] { ',' });
            if (ouArr.Length > 0)
            {
                for (int i = 0; i < ouArr.Length; i++)
                {
                    if (ouArr[i].Trim().ToUpper().IndexOf("OU=") >= 0 || ouArr[i].Trim().ToUpper().IndexOf("OU =") >= 0)
                    {
                        string[] ouName = ouArr[i].Split(new char[] { '=' });
                        if (ouName.Length > 1 && !string.IsNullOrEmpty(ouName[1]))
                            ous.Add(ouName[1]);
                    }
                }
            }
            else
            {
                ous.Add("PacteraUsers");
            }

            string ouPath = string.Empty;
            for (int i = ous.Count - 1; i >= 0; i--)
            {
                CheckOUPath(ous[i], ouPath, adUser, adPassword, adPath, ldapDomain);
                ouPath = string.IsNullOrEmpty(ouPath) ? ous[i] : (ouPath + "/" + ous[i]);
            }

            return ouPath;
        }
        /// <summary>
        /// 检测OU路径是否存在，不存在则创建
        /// </summary>
        /// <param name="ouName">Overseas</param>
        /// <param name="parentOuName">PacteraUsers</param>
        protected void CheckOUPath(string ouName, string parentOuName, string adUser, string adPassword, string adPath, string ldapDomain)
        {
            string ouPath = ouName;
            if (!string.IsNullOrEmpty(parentOuName))
                ouPath = parentOuName + "/" + ouName;

            if (!CheckOU2(ouPath, adUser, adPassword, adPath, ldapDomain))
            {
                CreateOrganizeUnit(adUser, adPassword, ouName, parentOuName, adPath, ldapDomain);
            }
            //if (!ADHelper.CheckOU(ouPath))
            //{
            //    ADHelper.CreateOrganizeUnit(ouName, parentOuName);
            //}
        }
        /// <summary>
        /// 检查组织单位（OU）是否存在
        /// </summary>
        /// <param name="sOUPath">组织单位名称，示例：PacteraUsers 或 PacteraUsers/Others</param>
        /// <param name="adUser"></param>
        /// <param name="adPassword"></param>
        /// <param name="adPath">示例：LDAP://192.168.88.64</param>
        /// <param name="ldapDomain">dc=testcode,dc=com</param>
        /// <returns></returns>
        protected bool CheckOU2(string sOUPath,string adUser,string adPassword,string adPath,string ldapDomain)
        {
            DirectoryEntry objOU = null;
            try
            {
                string path = GetOrganizeNamePath(sOUPath, adPath, ldapDomain);
                objOU = new DirectoryEntry(path, adUser, adPassword);
                string OUName = objOU.Name;
                //if (OUName == sOU)
                objOU.Close();
                objOU = null;
                return true;
            }
            catch
            {
                objOU.Close();
                objOU = null;
                return false;
            }
        }
        protected string GetOrganizeNamePath(string organizeUnit, string adPath,string ldapDomain)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(adPath);//示例：LDAP://192.168.88.64
            sb.Append("/");
            string path = sb.Append(SplitOrganizeNameToDN(organizeUnit, ldapDomain)).ToString();
            return path;
        }
        public string SplitOrganizeNameToDN(string organizeName,string ldapDomain)
        {
            StringBuilder sb = new StringBuilder();

            if (organizeName != null && organizeName.Length > 0)
            {
                string[] allOu = organizeName.Split(new char[] { '/', '\\' });
                for (int i = allOu.Length - 1; i >= 0; i--)
                {
                    string ou = allOu[i];
                    if (sb.Length > 0)
                    {
                        sb.Append(",");
                    }
                    sb.Append("OU=").Append(ou);
                }
            }
            //如果传入了组织名称，则添加,
            if (sb.Length > 0)
            {
                sb.Append(",");
            }
            sb.Append(ldapDomain);
            return sb.ToString();
        }
        /// <summary>
        /// 创建OU，需要指定连接到AD的授权信息
        /// </summary>
        /// <param name="adminName"></param>
        /// <param name="adminPassword"></param>
        /// <param name="ouName"></param>
        /// <param name="parentOrganizeUnit"></param>
        /// <param name="adPath"></param>
        /// <param name="ldapDomain"></param>
        /// <returns></returns>
        public DirectoryEntry CreateOrganizeUnit(string adminName, string adminPassword, string ouName, string parentOrganizeUnit, string adPath, string ldapDomain)
        {
            DirectoryEntry parentEntry = null;
            if (adminName == null || adminPassword == null)
            {
                parentEntry = new DirectoryEntry(GetOrganizeNamePath(parentOrganizeUnit, adPath, ldapDomain));
            }
            else
            {
                if (parentOrganizeUnit != "")
                {
                    parentEntry = new DirectoryEntry(GetOrganizeNamePath(parentOrganizeUnit, adPath, ldapDomain),
                       adminName, adminPassword,
                       AuthenticationTypes.Secure);
                }
                else
                {
                    string path = GetOrganizeNamePath(parentOrganizeUnit, adPath, ldapDomain);
                    parentEntry = new DirectoryEntry(path, adminName, adminPassword, AuthenticationTypes.Secure);
                }
            }
            DirectoryEntry organizeEntry = parentEntry.Children.Add("OU=" + ouName, "organizationalUnit");
            organizeEntry.CommitChanges();
            return organizeEntry;
        }

        protected string GetPacteraUsersOU(string domain)
        {
            var pacteraUsersOU = ConfigurationManager.AppSettings["PacteraUsersOU"];
            if (!ADHelper.ObjectExists(pacteraUsersOU, "OU"))
            {
                ADHelper.CreateOrganizeUnit(pacteraUsersOU, "");
            }
            //var ldap = LDAPStore.FirstOrDefault(l => l.LDAPDomain.ToLower() == domain.ToLower());
            return "OU=" + pacteraUsersOU + ",DC=cn-bj,DC=pactera,DC=com";
        }
        protected List<LocationMapping> GetLocationMappingStore()
        {
            List<LocationMapping> LocationMappingList = new List<LocationMapping>();

            try
            {
                using (System.Data.SqlClient.SqlDataReader reader = SqlHelper.ExecuteReader(ADConn, System.Data.CommandType.Text, "select * from T_LocationMapping"))
                {
                    while (reader.Read())
                    {
                        LocationMapping entity = new LocationMapping();
                        entity.Id = Convert.ToInt32(reader["Id"]);
                        entity.WorkLocation = Convert.ToString(reader["WorkLocation"]);
                        entity.Region = Convert.ToString(reader["Region"]);
                        entity.City = Convert.ToString(reader["City"]);
                        entity.State = Convert.ToString(reader["State"]);
                        entity.Country = Convert.ToString(reader["Country"]);
                        entity.Domain = Convert.ToString(reader["Domain"]);
                        entity.OU = Convert.ToString(reader["OU"]);
                        entity.Remarks = Convert.ToString(reader["Remarks"]);
                        LocationMappingList.Add(entity);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return LocationMappingList;
        }

        /// <summary>
        /// 临时使用
        /// </summary>
        /// <param name="adUser"></param>
        protected void ReadADUserInfo(ADUser adUser)
        {
            string conn = "server=172.16.253.172;database=ADDB;uid=sa;pwd=1234-abcd";
            string sql = "select * from T_ADUserInfo where SamAccountName = N'" + adUser.SamAccountName + "'";
            System.Data.SqlClient.SqlDataReader reader = SqlHelper.ExecuteReader(conn, CommandType.Text, sql);
            if (reader.Read())
            {
                adUser.CommonName = reader["CommonName"].ToString();
                adUser.Password = reader["Password"].ToString();
                adUser.UserPrincipalName = reader["UserPrincipalName"].ToString();
                adUser.DisplayName = reader["DisplayName"].ToString();
                adUser.Company = reader["Company"].ToString();
                adUser.Department = reader["Department"].ToString();
                adUser.Country = reader["Country"].ToString();
                adUser.State = reader["State"].ToString();
                adUser.OfficeLocation = reader["OfficeLocation"].ToString();
                adUser.StreetAddress = reader["StreetAddress"].ToString();
                adUser.Domain = reader["Domain"].ToString();
                adUser.OU = reader["OU"].ToString();
                adUser.Type = reader["Type"].ToString();
                adUser.ManagerNumber = reader["ManagerNumber"].ToString();
            }
        }

        ///// <summary>
        ///// 创建邮箱地址
        ///// </summary>
        ///// <param name="firstName"></param>
        ///// <param name="lastName"></param>
        ///// <returns></returns>
        //protected string GetUPNByEmailAddress(ADUser adUser)
        //{
        //    string emailAddress = string.Empty;
        //    if (!string.IsNullOrEmpty(adUser.GivenName) && !string.IsNullOrEmpty(adUser.SN))
        //    {
        //        string emailDomain = ConfigurationManager.AppSettings["EmailDomain"]; //"@pactera.com"; //"@extest.com";
        //        emailAddress = adUser.GivenName + "." + adUser.SN + emailDomain;

        //        while (IsExistEmailAddress(emailAddress))
        //        {
        //            System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(@"\d+$");
        //            if (reg.IsMatch(adUser.SN))
        //            {
        //                string num = reg.Match(adUser.SN).Value;
        //                string subLastName = adUser.SN.Substring(0, adUser.SN.LastIndexOf(num));
        //                adUser.SN = subLastName + (int.Parse(num) + 1).ToString();
        //            }
        //            else
        //            {
        //                adUser.SN = adUser.SN + "0";
        //            }
        //            emailAddress = adUser.GivenName + "." + adUser.SN + emailDomain;
        //        }
        //        adUser.CommonName = GetCommonName(adUser.GivenName, adUser.SN);
        //    }
        //    return emailAddress;
        //}
        protected string CreateUserPrincipalName(ADUser adUser)
        {
            string emailDomain = ConfigurationManager.AppSettings["EmailDomain"];
            string userPrincipalName = adUser.GivenName + "." + adUser.SN + emailDomain; //san.zhang@pactera.com

            bool isExist = IsExistEmailAddress(userPrincipalName) || IsADUserExistsByUPN(userPrincipalName);
            System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(@"\d+$");
            while (isExist)
            {
                if (reg.IsMatch(adUser.SN))
                {
                    string matchVal = reg.Match(adUser.SN).Value; //10
                    string subLastName = adUser.SN.Substring(0, adUser.SN.LastIndexOf(matchVal)); //zhang
                    adUser.SN = subLastName + (int.Parse(matchVal) + 1).ToString(); //zhang11
                }
                else
                {
                    adUser.SN = adUser.SN + "1"; //zhang1
                }

                userPrincipalName = adUser.GivenName + "." + adUser.SN + emailDomain;
                isExist = IsExistEmailAddress(userPrincipalName) || IsADUserExistsByUPN(userPrincipalName);
            }
            return userPrincipalName;
        }
        /// <summary>
        /// 设置CommonName(cn),givenName,sn
        /// </summary>
        /// <param name="adUser"></param>
        protected void DealwithRepeateCommonName(ADUser adUser)
        {
            adUser.GivenName = adUser.FirstName == null ? "" : adUser.FirstName.Trim().Replace(" ", "_");
            adUser.SN = adUser.LastName == null ? "" : adUser.LastName.Trim().Replace(" ", "_");
            if (!string.IsNullOrEmpty(adUser.GivenName) && !string.IsNullOrEmpty(adUser.SN))
            {
                //string emailDomain = ConfigurationManager.AppSettings["EmailDomain"]; //"@pactera.com"; //"@extest.com";
                adUser.CommonName = GetCommonName(adUser.GivenName, adUser.SN);
                try
                {
                    bool isRepeat = FindCommonNameInDomains(adUser.CommonName);
                    while (isRepeat)
                    {
                        System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(@"\d+"); //\d+$
                        //zhang10
                        if (reg.IsMatch(adUser.SN))
                        {
                            string matchVal = reg.Match(adUser.SN).Value; //10
                            string subLastName = adUser.SN.Substring(0, adUser.SN.LastIndexOf(matchVal)); //zhang
                            adUser.SN = subLastName + (int.Parse(matchVal) + 1).ToString(); //zhang11
                        }
                        else
                        {
                            adUser.SN = adUser.SN + "1"; //zhang1
                        }
                        adUser.CommonName = GetCommonName(adUser.GivenName, adUser.SN);
                        isRepeat = FindCommonNameInDomains(adUser.CommonName);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
            }
            //adUser.GivenName = adUser.CommonName; //san.zhang0
            //adUser.SN = string.Empty;
        }
        private bool FindCommonNameInDomains(string commonName)
        {
            bool result = false;
            try
            {
                foreach (var ldap in LDAPStore)
                {
                    ADHelper.DomainName = ldap.DomainName;
                    ADHelper.LDAPDomain = ldap.LDAPDomain; //ADHelper.DomainName = ldap.LDAPDomain;
                    ADHelper.ADPath = ldap.ADPath;
                    ADHelper.ADUser = ldap.ADUser;
                    ADHelper.ADPassword = ldap.ADPassword;
                    if (ADHelper.IsUserExists(commonName))
                    {
                        result = true;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("查找CommonName是否重复时出错：" + ex.Message);
            }
            return result;
        }

        protected ExecResult CommonNameUPNDisplayName(ADUser adUser)
        {
            #region UserPrincipalName

            if (string.IsNullOrEmpty(adUser.FirstName.Trim()))
                return ExecResult.FirstNameEmpty;
            if (string.IsNullOrEmpty(adUser.LastName.Trim()))
                return ExecResult.LastNameEmpty;

            adUser.GivenName = adUser.FirstName.Trim().Replace(" ", "");
            adUser.SN = adUser.LastName.Trim().Replace(" ", "");

            string emailDomain = ConfigurationManager.AppSettings["EmailDomain"];
            string emailAddressName = adUser.GivenName + "." + adUser.SN;
            string userPrincipalName = emailAddressName + emailDomain; //san.zhang@pactera.com
            string curTime = string.Empty;

            int i = 0;
            while (IsExistEmailAddress(emailAddressName) || IsADUserExistsByUPN(userPrincipalName))
            {
                i++;
                emailAddressName = adUser.GivenName + "." + adUser.SN + i.ToString();
                userPrincipalName = emailAddressName + emailDomain;
            }
            adUser.UserPrincipalName = userPrincipalName;

            #endregion

            #region DisplayName & CommonName

            if (string.IsNullOrEmpty(adUser.Department))
                return ExecResult.DepartmentEmpty;

            string[] departmentNames = adUser.Department.Split(new char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);

            string region = GetRegionByLocation(adUser.Location);
            string displayName = ToTitleCase(adUser.SN.Trim()) + "," + ToTitleCase(adUser.GivenName.Trim()); //+ (i > 0 ? i.ToString() : "");
            adUser.DisplayName = displayName;

            string tempName = string.Empty;
            bool isexist = IsADUserExistsByDisplayName(displayName);
            if (!isexist)
                isexist = IsExistsCommonName(adUser);

            int num = 0;
            int departmenti = 0; //部门级（1,2,...）
            while (isexist)
            {
                if (departmentNames.Length > departmenti)
                {
                    tempName = displayName + "(" + departmentNames[departmenti].Trim() + "/" + region + ")";
                    adUser.DisplayName = tempName;
                    isexist = IsADUserExistsByDisplayName(tempName);
                    if (!isexist)
                        isexist = IsExistsCommonName(adUser);
                    departmenti++;
                }
                else
                {
                    num++;
                    if (departmentNames.Length > 0)
                    {
                        tempName = displayName + num.ToString() + "(" + departmentNames[departmentNames.Length - 1].Trim() + "/" + region + ")";
                    }
                    else
                    {
                        tempName = displayName + num.ToString();
                    }
                    adUser.DisplayName = tempName;
                    isexist = IsADUserExistsByDisplayName(tempName);
                    if (!isexist)
                        isexist = IsExistsCommonName(adUser);

                    //string suffix = string.Empty;
                    //if (i > 0)
                    //    suffix = i.ToString();

                    //tempName = displayName + suffix;
                    //isexist = IsADUserExistsByDisplayName(tempName);
                    //i++;
                }
            }

            #endregion

            #region CommonName

            //if (adUser.DisplayName.Length > 200)
            //{
            //    adUser.CommonName = adUser.DisplayName.Substring(0, 200);
            //}
            //else
            //{
            //    adUser.CommonName = adUser.DisplayName;
            //}

            //string comName = adUser.CommonName;
            //int commoni = 0;
            //while (FindCommonNameInDomains(comName))
            //{
            //    comName = adUser.CommonName + "_" + (++commoni).ToString();
            //}
            //adUser.CommonName = comName;

            //if (adUser.DisplayName.Length > 200)
            //{
            //    adUser.CommonName = adUser.GivenName + "." + adUser.SN;
            //}
            //else
            //{
            //    adUser.CommonName = adUser.DisplayName;
            //}
            //System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(@"\d+$");
            //while (FindCommonNameInDomains(adUser.CommonName))
            //{
            //    if (reg.IsMatch(adUser.CommonName))
            //    {
            //        string matchVal = reg.Match(adUser.CommonName).Value; //10
            //        string subLastName = adUser.CommonName.Substring(0, adUser.CommonName.LastIndexOf(matchVal));
            //        adUser.CommonName = subLastName + (int.Parse(matchVal) + 1).ToString();
            //    }
            //    else
            //    {
            //        if (i > 0)
            //            adUser.CommonName += i.ToString();
            //        else
            //            adUser.CommonName += "1";
            //    }
            //}

            #endregion

            if (i > 0)
            {
                adUser.SN += i.ToString();
            }

            return ExecResult.Success;
        }
        protected bool IsExistsCommonName(ADUser adUser)
        {
            if (adUser.DisplayName.Length > 200)
            {
                adUser.CommonName = adUser.DisplayName.Substring(0, 200);
            }
            else
            {
                adUser.CommonName = adUser.DisplayName;
            }

            return FindCommonNameInDomains(adUser.CommonName);
        }

        /// <summary>
        /// 查询是否存在给定的邮箱地址
        /// </summary>
        /// <param name="emailAddress">邮箱地址</param>
        /// <returns></returns>
        protected bool IsExistEmailAddress(string emailAddress)
        {
            string strSql = @"select Count(*) from V_Employee_Info where E_MAIL like N'" + emailAddress + "%'";
            object count = SqlHelper.ExecuteScalar(Conn, CommandType.Text, strSql);
            if ((int)count > 0)
                return true;
            else
                return false;
        }
        
        protected string ToTitleCase(string str)
        {
            if (string.IsNullOrEmpty(str))
                return "";
            else
                return str.Substring(0, 1).ToUpper() + str.Substring(1);
        }

        /// <summary>
        /// 获取员工类型（小于B9是K1，大于等于B9或其他是E4）
        /// </summary>
        /// <param name="gradeName">职级</param>
        /// <returns></returns>
        protected string GetEmployeeType(string gradeName)
        {
            string type = string.Empty;

            if (!string.IsNullOrEmpty(gradeName))
            {
                System.Text.RegularExpressions.Regex reg1 = new System.Text.RegularExpressions.Regex("[A-Za-z]+");
                System.Text.RegularExpressions.Regex reg2 = new System.Text.RegularExpressions.Regex("\\d+");
                if (reg1.IsMatch(gradeName) && reg2.IsMatch(gradeName))
                {
                    if (reg1.Match(gradeName).Value.ToUpper() == "B" && int.Parse(reg2.Match(gradeName).Value) < 9)
                    {
                        type = "E1";
                    }
                    else
                    {
                        type = "E4";
                    }
                }
            }

            return type;
        }

        protected string GetEmployeePositionName(string employeeId)
        {
            string positionName = string.Empty;
            string sql = "select isnull(POSITION_EN_NAME,'') as POSITION_EN_NAME from V_Employee_Position where EMPLOYEE_NUMBER = N'" + employeeId + "'";
            try
            {
                System.Data.SqlClient.SqlDataReader reader = SqlHelper.ExecuteReader(Conn, CommandType.Text, sql);
                if (reader.Read())
                {
                    positionName = reader["POSITION_EN_NAME"].ToString();
                }
            }
            catch
            {
                positionName = string.Empty;
            }
            return positionName;
        }

        /// <summary>
        /// 将AD用户信息插入到AD数据库中
        /// </summary>
        /// <param name="adUser"></param>
        /// <returns></returns>
        protected ExecResult InsertADUserInfoToDB(ADUser adUser)
        {
            ExecResult result = ExecResult.Failure;
            int flagVal = -1; //IsRegularEmployee(adUser.Grade) ? -1 : 1; //第一次插入，标识位为-1，成功常见账号之后，标志位更新为0
            try
            {
                string sql = "select Count(*) from T_ADUserInfo where SamAccountName = N'" + adUser.SamAccountName + "'";
                object obj = SqlHelper.ExecuteScalar(ADConn, CommandType.Text, sql);

                sql = string.Empty;
                //ADDB中不存在该员工号信息
                if (obj != null && int.Parse(obj.ToString()) < 1)
                {
                    sql = "Insert into T_ADUserInfo values(N'" + adUser.SamAccountName +
                        "',N'" + adUser.CommonName +
                        "',N'" + adUser.Password +
                        "',N'" + adUser.UserPrincipalName +
                        "',N'" + adUser.FirstName +
                        "',N'" + adUser.MiddleName +
                        "',N'" + adUser.LastName +
                        "',N'" + adUser.GivenName +
                        "',N'" + adUser.SN +
                        "',N'" + adUser.DisplayName +
                        "',N'" + adUser.Company +
                        "',N'" + adUser.DepartmentNo +
                        "',N'" + adUser.Department +
                        "',N'" + adUser.RootDepartmentNo +
                        "',N'" + adUser.RootDepartmentName +
                        "',N'" + adUser.Country +
                        "',N'" + adUser.State +
                        "',N'" + adUser.City +
                        "',N'" + adUser.Location +
                        "',N'" + adUser.OfficeLocation +
                        "',N'" + adUser.StreetAddress +
                        "',N'" + adUser.Domain +
                        "',N'" + adUser.OU +
                        "',N'" + adUser.Grade +
                        "',N'" + adUser.Type +
                        "',N'" + adUser.PositionName +
                        "',N'" + adUser.ManagerNumber +
                        "',N'" + adUser.ManagerName +
                        "',N'" + adUser.PhoneNumber +
                        "',N'" + adUser.MobileNumber +
                        "','" + adUser.IsRelocate +
                        "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                        "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                        "'," + flagVal + ")";
                }
                else
                {
                    RecordADUserInfo(adUser, "Repeat insert");
                    result = ExecResult.Failure;
                    #region Update
                    ////ADDB中已存在该员工号信息
                    //sql = "Update T_ADUserInfo Set CommonName = N'" + adUser.CommonName +
                    //    "',Password = N'" + adUser.Password +
                    //    "',UserPrincipalName = N'" + adUser.UserPrincipalName +
                    //    "',FirstName = N'" + adUser.FirstName +
                    //    "',MiddleName = N'" + adUser.MiddleName +
                    //    "',LastName = N'" + adUser.LastName +
                    //    "',GivenName = N'" + adUser.GivenName +
                    //    "',SN = N'" + adUser.SN +
                    //    "',DisplayName = N'" + adUser.DisplayName +
                    //    "',Company = N'" + adUser.Company +
                    //    "',DepartmentNo = N'" + adUser.DepartmentNo +
                    //    "',Department = N'" + adUser.Department +
                    //    "',RootDepartmentNo = N'" + adUser.RootDepartmentNo +
                    //    "',RootDepartmentName = N'" + adUser.RootDepartmentName +
                    //    "',Country = N'" + adUser.Country +
                    //    "',State = N'" + adUser.State +
                    //    "',City = N'" + adUser.City +
                    //    "',Location = N'" + adUser.Location +
                    //    "',OfficeLocation = N'" + adUser.OfficeLocation +
                    //    "',StreetAddress = N'" + adUser.StreetAddress +
                    //    "',Domain = N'" + adUser.Domain +
                    //    "',OU = N'" + adUser.OU +
                    //    "',Grade = N'" + adUser.Grade +
                    //    "',Type = N'" + adUser.Type +
                    //    "',PositionName = N'" + adUser.PositionName +
                    //    "',ManagerNumber = N'" + adUser.ManagerNumber +
                    //    "',ManagerName = N'" + adUser.ManagerName +
                    //    "',PhoneNumber = N'" + adUser.PhoneNumber +
                    //    "',MobileNumber = N'" + adUser.MobileNumber +
                    //    "',IsRelocate = '" + adUser.IsRelocate +
                    //    "',RUpdateTime = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                    //    "',RFlag = " + flagVal + " where SamAccountName = N'" + adUser.SamAccountName + "'";
                    #endregion
                }

                if (!string.IsNullOrEmpty(sql))
                {
                    SqlHelper.ExecuteNonQuery(ADConn, CommandType.Text, sql);
                    result = ExecResult.Success;
                }
            }
            catch(Exception ex)
            {
                result = ExecResult.Failure;
                LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "Insert AD Account (" + adUser.SamAccountName + ") to ADDB Failure: " + ex.Message));
                //RecordADUserInfo(adUser, ex.Message);
            }

            return result;
        }
        protected List<ADUser> ReadADUserInfoFromDB()
        {
            string sql = "select * from T_ADUserInfo where RFlag = -1";
            System.Data.SqlClient.SqlDataReader reader = SqlHelper.ExecuteReader(ADConn, CommandType.Text, sql);

            List<ADUser> adUserList = new List<ADUser>();
            while (reader.Read())
            {
                ADUser adUser = new ADUser();
                adUser.SamAccountName = DBNullToString(reader["SamAccountName"]);
                adUser.CommonName = DBNullToString(reader["CommonName"]);
                adUser.Password = DBNullToString(reader["Password"]);
                adUser.UserPrincipalName = DBNullToString(reader["UserPrincipalName"]);
                adUser.FirstName = DBNullToString(reader["FirstName"]);
                adUser.MiddleName = DBNullToString(reader["MiddleName"]);
                adUser.LastName = DBNullToString(reader["LastName"]);
                adUser.GivenName = DBNullToString(reader["GivenName"]);
                adUser.SN = DBNullToString(reader["SN"]);
                adUser.DisplayName = DBNullToString(reader["DisplayName"]);
                adUser.Company = DBNullToString(reader["Company"]);
                adUser.DepartmentNo = DBNullToString(reader["DepartmentNo"]);
                adUser.DisplayName = DBNullToString(reader["DisplayName"]);
                adUser.RootDepartmentNo = DBNullToString(reader["RootDepartmentNo"]);
                adUser.RootDepartmentName = DBNullToString(reader["RootDepartmentName"]);
                adUser.Country = DBNullToString(reader["Country"]);
                adUser.State = DBNullToString(reader["State"]);
                adUser.City = DBNullToString(reader["City"]);
                adUser.Location = DBNullToString(reader["Location"]);
                adUser.OfficeLocation = DBNullToString(reader["OfficeLocation"]);
                adUser.StreetAddress = DBNullToString(reader["StreetAddress"]);
                adUser.Domain = DBNullToString(reader["Domain"]);
                adUser.OU = DBNullToString(reader["OU"]);
                adUser.Grade = DBNullToString(reader["Grade"]);
                adUser.Type = DBNullToString(reader["Type"]);
                adUser.PositionName = DBNullToString(reader["PositionName"]);
                adUser.ManagerNumber = DBNullToString(reader["ManagerNumber"]);
                adUser.ManagerName = DBNullToString(reader["ManagerName"]);
                adUser.PhoneNumber = DBNullToString(reader["PhoneNumber"]);
                adUser.MobileNumber = DBNullToString(reader["MobileNumber"]);
                adUserList.Add(adUser);
            }

            return adUserList;
        }
        protected ExecResult UpdateADUserInfo(ADUser adUser)
        {
            ExecResult result = ExecResult.Failure;
            try
            {
                string sql = "Update T_ADUserInfo Set RFlag = 0,RUpdateTime = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' where SamAccountName = N'" + adUser.SamAccountName + "'";
                
                #region UpdateAll
                //string sql = "Update T_ADUserInfo Set CommonName = N'" + adUser.CommonName +
                //    "',Password = N'" + adUser.Password +
                //    "',UserPrincipalName = N'" + adUser.UserPrincipalName +
                //    "',FirstName = N'" + adUser.FirstName +
                //    "',MiddleName = N'" + adUser.MiddleName +
                //    "',LastName = N'" + adUser.LastName +
                //    "',GivenName = N'" + adUser.GivenName +
                //    "',SN = N'" + adUser.SN +
                //    "',DisplayName = N'" + adUser.DisplayName +
                //    "',Company = N'" + adUser.Company +
                //    "',DepartmentNo = N'" + adUser.DepartmentNo +
                //    "',Department = N'" + adUser.Department +
                //    "',RootDepartmentNo = N'" + adUser.RootDepartmentNo +
                //    "',RootDepartmentName = N'" + adUser.RootDepartmentName +
                //    "',Country = N'" + adUser.Country +
                //    "',State = N'" + adUser.State +
                //    "',City = N'" + adUser.City +
                //    "',Location = N'" + adUser.Location +
                //    "',OfficeLocation = N'" + adUser.OfficeLocation +
                //    "',StreetAddress = N'" + adUser.StreetAddress +
                //    "',Domain = N'" + adUser.Domain +
                //    "',OU = N'" + adUser.OU +
                //    "',Grade = N'" + adUser.Grade +
                //    "',Type = N'" + adUser.Type +
                //    "',PositionName = N'" + adUser.PositionName +
                //    "',ManagerNumber = N'" + adUser.ManagerNumber +
                //    "',ManagerName = N'" + adUser.ManagerName +
                //    "',PhoneNumber = N'" + adUser.PhoneNumber +
                //    "',MobileNumber = N'" + adUser.MobileNumber +
                //    "',IsRelocate = '" + adUser.IsRelocate +
                //    "',RUpdateTime = '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                //    "',RFlag = 0 where SamAccountName = N'" + adUser.SamAccountName + "'";
                #endregion

                SqlHelper.ExecuteNonQuery(ADConn, CommandType.Text, sql);
                result = ExecResult.Success;
            }
            catch
            {
                RecordADUserInfo(adUser, "更新失败（-1 to 0）");
            }

            return result;
        }

        protected void RecordADUserInfo(ADUser adUser, string msg)
        {
            LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "Create AD Account(" + msg + "): SamAccountName: " + adUser.SamAccountName + ",FirstName:" + adUser.FirstName + ",GivenName:" + adUser.GivenName + ",LastName:" + adUser.LastName + ",sn:" + adUser.SN + ",Location:" + adUser.Location + ",Grade:" + adUser.Grade));
        }
        /// <summary>
        /// 是否为正式员工
        /// </summary>
        /// <param name="grade"></param>
        /// <returns>True:正式员工</returns>
        protected bool IsRegularEmployee(string grade)
        {
            string gradeFix = grade.Substring(0, 1).ToUpper();
            return !(gradeFix == "C" || gradeFix == "I");
        }

        /// <summary>
        /// 从我们自己的AD数据表中删除AD用户信息
        /// </summary>
        /// <param name="employeeId"></param>
        /// <returns></returns>
        protected int DeleteADUserInfo(string employeeId)
        {
            string sql = "delete from T_ADUserInfo where SamAccountName = N'" + employeeId + "'";
            return SqlHelper.ExecuteNonQuery(ADConn, CommandType.Text, sql);
        }

        protected List<EmailInfo> GetEmailInfosByEmployeeId(List<string> employeeIds)
        {
            if (employeeIds == null || employeeIds.Count == 0)
                return new List<EmailInfo>();

            string sql = "select * from T_ADUserInfo where SamAccountName in (";
            foreach (string employeeId in employeeIds)
            {
                sql += "'" + employeeId + "',";
            }
            sql = sql.Substring(0, sql.Length - 1) + ")";
            sql += " and RFlag > 0";
            List<EmailInfo> emailInfos = new List<EmailInfo>();
            DataSet dataSet = SqlHelper.ExecuteDataset(ADConn, CommandType.Text, sql);
            if (dataSet != null && dataSet.Tables.Count > 0)
            {
                DataTable dt = dataSet.Tables[0];
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    EmailInfo emailInfo = new EmailInfo();
                    emailInfo.EmployeeId = dt.Rows[i]["SamAccountName"].ToString();
                    emailInfo.Email = dt.Rows[i]["UserPrincipalName"].ToString();
                    emailInfo.Password = dt.Rows[i]["Password"].ToString();
                    emailInfos.Add(emailInfo);
                }
            }
            //System.Data.SqlClient.SqlDataReader reader = SqlHelper.ExecuteReader(ADConn, CommandType.Text, sql);
            //while (reader.Read())
            //{
            //    EmailInfo emailInfo = new EmailInfo();
            //    emailInfo.EmployeeId = reader["SamAccountName"].ToString();
            //    emailInfo.Email = reader["UserPrincipalName"].ToString();
            //    emailInfo.Password = reader["Password"].ToString();
            //    emailInfos.Add(emailInfo);
            //}
            return emailInfos;
        }
        protected EmailInfo GetEmailInfoByEmployeeId(string employeeId)
        {
            if (string.IsNullOrEmpty(employeeId))
                return new EmailInfo();

            string sql = "select * from T_ADUserInfo where SamAccountName = '" + employeeId + "' and RFlag > 3";
            System.Data.SqlClient.SqlDataReader reader = SqlHelper.ExecuteReader(ADConn, CommandType.Text, sql);
            EmailInfo emailInfo = null;
            if (reader.Read())
            {
                emailInfo = new EmailInfo();
                emailInfo.EmployeeId = reader["SamAccountName"].ToString();
                emailInfo.Email = reader["UserPrincipalName"].ToString();
                emailInfo.Password = reader["Password"].ToString();
            }
            return emailInfo;
        }

        #region 创建ADUser

        /// <summary>
        /// 创建AD用户
        /// </summary>
        /// <param name="e"></param>
        protected ExecResult CreateADUser(ADUser adUser)
        {
            //登录到相应的域控服务器 //LogonDomainComputer(adUser.Domain); 
            if (LDAPStore == null)
                LDAPStore = new List<LDAP>();

            LDAP ldap = LDAPStore.FirstOrDefault(e => e.DomainName.ToLower() == adUser.Domain.ToLower());
            if (ldap == null)
                ldap = LDAPStore.FirstOrDefault(l => l.ID == Convert.ToInt32(ConfigurationManager.AppSettings["GroupLDAPID"]));

            ADHelper.DomainName = ldap.DomainName;
            ADHelper.LDAPDomain = ldap.LDAPDomain;
            ADHelper.ADPath = ldap.ADPath;
            ADHelper.ADUser = ldap.ADUser;
            ADHelper.ADPassword = ldap.ADPassword;

            ExecResult result = ExecResult.Failure;
            if (adUser != null)
            {
                try
                {
                    string ldapDN = GetLocateOUName(adUser.OU, ldap.ADUser, ldap.ADPassword, ldap.ADPath, ldap.LDAPDomain); //老的程序："OU=PacteraUsers"; //都在主域（cn-bj.pactera.com）的PacteraUsers下
                    //创建用户
                    DirectoryEntry de = CreateNewADAccount(ldapDN, adUser.CommonName, adUser.SamAccountName, adUser.Password, ldap.ADPath, ldap.ADUser, ldap.ADPassword);
                    //DirectoryEntry de = ADHelper.CreateNewUser(ldapDN, adUser.CommonName, adUser.SamAccountName, adUser.Password, ldap.ADPath, ldap.ADUser, ldap.ADPassword);
                    //设置属性
                    if (de != null)
                    {
                        SetUserProperty(de, adUser);
                        result = ExecResult.Success;
                    }
                    else
                    {
                        LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "CreateADUser, " + adUser.SamAccountName + ", Failure."));
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "CreateADUser Exception: " + adUser.SamAccountName + "," + ex.Message));
                }
            }
            return result;
        }

        /// <summary>
        /// 创建AD账号
        /// </summary>
        /// <param name="ldapDN">PacteraUsers 或 PacteraUsers/Others（旧的：OU=PacteraUsers 或 CN=Users）</param>
        /// <param name="commonName"></param>
        /// <param name="sAMAccountName"></param>
        /// <param name="password"></param>
        /// <param name="adPath">示例：LDAP://192.168.88.64</param>
        /// <param name="adUser"></param>
        /// <param name="adUserPwd"></param>
        /// <returns></returns>
        public DirectoryEntry CreateNewADAccount(string ldapDN, string commonName, string sAMAccountName, string password, string adPath, string adUser, string adUserPwd)
        {
            DirectoryEntry deUser = null;
            string path = ADHelper.GetOrganizeNamePath(ldapDN, adPath); //示例path：LDAP://172.16.254.21/OU=Beijing,OU=Office365Users,DC=pactera,DC=com
            DirectoryEntry entry = new DirectoryEntry(path, adUser, adUserPwd, AuthenticationTypes.Secure);

            if (entry != null)
            {
                try
                {
                    deUser = entry.Children.Add("CN=" + commonName.Replace(",", "\\,"), "user");
                    deUser.Properties["sAMAccountName"].Value = sAMAccountName;
                    deUser.CommitChanges();
                }
                catch (Exception ex)
                {
                    deUser = null;
                    LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "创建账号失败：" + sAMAccountName + "," + ex.Message));
                }

                if (deUser != null)
                {
                    try
                    {
                        deUser.Invoke("SetPassword", new object[] { password });
                        //deUser.CommitChanges();
                        deUser.Properties["userAccountControl"][0] = ADHelper.ADS_USER_FLAG_ENUM.ADS_UF_NORMAL_ACCOUNT | ADHelper.ADS_USER_FLAG_ENUM.ADS_UF_DONT_EXPIRE_PASSWD;
                        deUser.CommitChanges();
                    }
                    catch (Exception ex)
                    {
                        LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "设置密码和权限失败：" + sAMAccountName + "," + ex.Message));
                    }
                    finally
                    {
                        deUser.Close();
                    }
                }
            }
            return deUser;
        }

        /// <summary>
        /// 更新AD用户的属性和组
        /// </summary>
        /// <param name="de"></param>
        /// <param name="e"></param>
        protected void SetUserProperty(DirectoryEntry de, ADUser adUser)
        {
            try
            {
                //更新属性
                if (!string.IsNullOrEmpty(adUser.ManagerNumber) && ADHelper.IsAccExists(adUser.ManagerNumber))
                {
                    var manager = ADHelper.FindObject("user", adUser.ManagerNumber);
                    if (manager != null)
                        ADHelper.SetProperty(de, "manager", manager.Properties["distinguishedName"][0].ToString());
                }
                ADHelper.SetProperty(de, "department", adUser.Department);
                ADHelper.SetProperty(de, "company", adUser.Company);
                ADHelper.SetProperty(de, "title", adUser.PositionName);
                ADHelper.SetProperty(de, "physicalDeliveryOfficeName", adUser.OfficeLocation);
                ADHelper.SetProperty(de, "givenName", adUser.GivenName); //名,First name
                ADHelper.SetProperty(de, "sn", adUser.SN); //姓,Last name
                ADHelper.SetProperty(de, "userprincipalname", adUser.UserPrincipalName);
                ADHelper.SetProperty(de, "displayname", adUser.DisplayName);
                //ADHelper.SetProperty(de, "telephoneNumber", e.PHONE_NUMBER);
                de.CommitChanges();
                de.Close();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            DirectoryEntry oUser = ADHelper.GetDirectoryEntryByAccount(adUser.SamAccountName);
            //LogonPrimaryAD();
            DirectoryEntry oGroup = null;
            //更新组
            var groupPrefix = ConfigurationManager.AppSettings["GroupPrefix"];
            if (!string.IsNullOrEmpty(adUser.Company))
            {
                oGroup = ADHelper.GetDirectoryEntryOfGroup(adUser.Company);
                ADHelper.AddUserToGroup1(oGroup, oUser);
            }
            if (!string.IsNullOrEmpty(adUser.Country))
            {
                oGroup = ADHelper.GetDirectoryEntryOfGroup(groupPrefix + adUser.Country);
                ADHelper.AddUserToGroup1(oGroup, oUser);
            }
            if (!string.IsNullOrEmpty(adUser.Country) && !string.IsNullOrEmpty(adUser.Location))
            {
                oGroup = ADHelper.GetDirectoryEntryOfGroup(groupPrefix + adUser.Country + "_" + adUser.Location);
                ADHelper.AddUserToGroup1(oGroup, oUser);
            }
            if (!string.IsNullOrEmpty(adUser.OfficeLocation))
            {
                oGroup = ADHelper.GetDirectoryEntryOfGroup(groupPrefix + adUser.OfficeLocation);
                ADHelper.AddUserToGroup1(oGroup, oUser);
            }
        }

        protected string GetServerAndOUByEmployee(ADUser adUser)
        {
            var location = GetEmployeeLocation(adUser);
            var mapping = LDAPLocationMappingStore.FirstOrDefault(m => m.Location == location);
            if (mapping != null)
            {
                var ldap = LDAPStore.FirstOrDefault(l => l.ID == mapping.LDAPID);

                ADHelper.DomainName = ldap.DomainName;
                ADHelper.LDAPDomain = ldap.LDAPDomain; //ADHelper.DomainName = ldap.LDAPDomain;
                ADHelper.ADPath = ldap.ADPath;
                ADHelper.ADUser = ldap.ADUser;
                ADHelper.ADPassword = ldap.ADPassword;

                return mapping.OU;
            }
            else
            {
                return string.Empty;
            }
        }

        protected string GetEmployeeLocation(ADUser adUser)
        {
            string location = string.Empty;
            if (adUser.Country == "CHN" || adUser.Country == "HKG" || adUser.Country == "MAC" || adUser.Country == "TWN")
            {
                //国内
                var list = new List<string>() { "CN-BJ", "BEIJING", "DALIAN", "WUXI", "SHENZHEN", "SHANGHAI" };
                location = list.Contains(adUser.Location.ToUpper()) ? adUser.Location.ToUpper() : "OTHERS";
            }
            else
            {
                //海外
                location = "OVERSEA";
            }
            return location;
        }

        #endregion

        //public void ReName(string samAccountName)
        //{
        //    DirectoryEntry accDe = GetAccount(samAccountName);
        //    if (accDe != null)
        //    {
        //        if (accDe.Properties.Contains("name"))
        //        {
        //            string name = accDe.Properties["name"][0].ToString();
        //        }
        //        if (accDe.Properties.Contains("cn"))
        //        {
        //            string cn = accDe.Properties["cn"][0].ToString();
        //        }

        //        //accDe.Rename("hai.lin");
        //        //accDe.CommitChanges();
        //    }
        //}

        //public DirectoryEntry GetAccount(string samAccountName)
        //{
        //    DirectoryEntry deAcc = null;
        //    try
        //    {
        //        foreach (var ldap in LDAPStore)
        //        {
        //            ADHelper.DomainName = ldap.DomainName;
        //            ADHelper.LDAPDomain = ldap.LDAPDomain; //ADHelper.DomainName = ldap.LDAPDomain;
        //            ADHelper.ADPath = ldap.ADPath;
        //            ADHelper.ADUser = ldap.ADUser;
        //            ADHelper.ADPassword = ldap.ADPassword;

        //            DirectoryEntry de = ADHelper.GetDirectoryObject();
        //            DirectorySearcher deSearch = new DirectorySearcher(de);
        //            deSearch.Filter = "(&(&(objectCategory=person)(objectClass=user))(cn=" + samAccountName + "))";
        //            //deSearch.Filter = "(&(&(objectCategory=person)(objectClass=user))(sAMAccountName=" + samAccountName + "))";
        //            SearchResult oneAcc = deSearch.FindOne();

        //            if (oneAcc != null)
        //            {
        //                //deAcc = new DirectoryEntry(oneAcc.Path);
        //                deAcc = oneAcc.GetDirectoryEntry();
        //                break;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception("查找samAccountName是否重复时出错：" + ex.Message);
        //    }
        //    return deAcc;
        //}

        //public void CreateTestADAccount(ADUser adUser)
        //{
        //    LogonDomainComputer(adUser.Domain);
        //    string ldapDN = GetLocateOUName(adUser.OU); //老的程序："OU=PacteraUsers"; //都在主域（cn-bj.pactera.com）的PacteraUsers下
        //    //创建用户
        //    DirectoryEntry de = ADHelper.CreateNewUser(ldapDN, adUser.CommonName, adUser.SamAccountName, adUser.Password);

        //    //DirectoryEntry deAcc = null;
        //    //DirectoryEntry de = ADHelper.GetDirectoryObject();
        //    //DirectorySearcher deSearch = new DirectorySearcher(de);
        //    ////deSearch.Filter = "(&(&(objectCategory=person)(objectClass=user))(sAMAccountName=" + adUser.SamAccountName + "))";
        //    //deSearch.Filter = "(&(&(objectCategory=person)(objectClass=user))(cn=" + adUser.CommonName + "))";
        //    //SearchResult oneAcc = deSearch.FindOne();

        //    //if (oneAcc != null)
        //    //{
        //    //    //deAcc = new DirectoryEntry(oneAcc.Path);
        //    //    deAcc = oneAcc.GetDirectoryEntry();

        //    //    if (deAcc.Properties.Contains("name"))
        //    //    {
        //    //        string name = deAcc.Properties["name"][0].ToString();
        //    //    }
        //    //    if (deAcc.Properties.Contains("cn"))
        //    //    {
        //    //        string cn = deAcc.Properties["cn"][0].ToString();
        //    //    }
        //    //}
        //}
    }
}
