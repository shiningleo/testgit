using ADModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Text;
using System.Threading;
using System.Xml.Linq;

namespace Bussiness
{
    public class Business : BaseBusiness
    {
        private static object syncroot = new object();
        AutoResetEvent[] events;

        public Business()
        {

        }

        /// <summary>
        /// 业务类入口
        /// </summary>
        /// <param name="ldap"></param>
        public void Start(int flag, bool createOrUpdateGroup)
        {
            if (flag == 0)
            {
                if (createOrUpdateGroup)
                {
                    //1、创建组
                    CreateADGroup();
                }

                //2、从DataHub获取用户
                var eList = GetEmployeeInfoList();
                //var eList = GetEmployeeInfoList().Where(e => e.EMPLOYEE_NUMBER == "P0015448");

                List<string> whiteNameList = ReadWhiteNameList();
                foreach (EmployeeInfo e in eList)
                {
                    //3、修改AD用户属性
                    UpdateADUser(e, whiteNameList);
                }
            }
            else if (flag == 1)
            {
                //var r = DeleteDepartmentGroup();
            }

            #region 多线程
            //if (flag == 0)
            //{
            //    if (createOrUpdateGroup)
            //    {
            //        //1、创建组
            //        CreateADGroup();
            //    }

            //    //2、从DataHub获取用户
            //    var eList = GetEmployeeInfoList();
            //    //var eList = GetEmployeeInfoList().Where(e => e.EMPLOYEE_NUMBER == "P0015448");

            //    int threadCount = 10;
            //    ThreadPool.SetMaxThreads(threadCount, threadCount);
            //    events = new AutoResetEvent[threadCount + 1];
            //    for (int i = 0; i < threadCount + 1; i++)
            //    {
            //        events[i] = new AutoResetEvent(false);
            //    }

            //    for (int i = 0; i < eList.Count; i++)
            //    {
            //        MyParameter myParam;
            //        if (i == eList.Count - 1)
            //            myParam = new MyParameter(eList[i], events[threadCount]);
            //        else
            //            myParam = new MyParameter(eList[i], events[i % threadCount]);
            //        //3、修改AD用户属性
            //        ThreadPool.QueueUserWorkItem(UpdateADUser, myParam);
            //    }
            //    //foreach (EmployeeInfo e in eList)
            //    //{
            //    //    //3、修改AD用户属性
            //    //    ThreadPool.QueueUserWorkItem(UpdateADUser, e);
            //    //    //UpdateADUser(e);
            //    //}
            //}
            //else if (flag == 1)
            //{
            //    var r = DeleteDepartmentGroup();
            //}

            //WaitHandle.WaitAll(events);
            //Thread.Sleep(1000);
            #endregion
        }

        /// <summary>
        /// 多线程调用
        /// </summary>
        /// <param name="o"></param>
        public void UpdateADUser(object o)
        {
            //MyParameter myParam = new MyParameter();
            //try
            //{
            //    myParam = (MyParameter)o;
            //}
            //catch
            //{
            //    Monitor.Enter(syncroot);
            //    LogHelper.WriteLog(new LogModel(Level.Info, DateTime.Now, "object -> MyParameter，转换失败。"));
            //    Monitor.Exit(syncroot);
            //    myParam.Are.Set();
            //    return;
            //}
            //UpdateADUser(myParam.EmployeeInfo);
            //myParam.Are.Set();
        }
        /// <summary>
        /// 更新AD用户
        /// </summary>
        /// <param name="emo"></param>
        public void UpdateADUser(EmployeeInfo e, List<string> whiteNames)
        {
            if (e == null) return;
            try
            {
                ChangeUserProperty(e, whiteNames);
            }
            catch (Exception ex)
            {
                //Monitor.Enter(syncroot);
                LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "UpdateADUser," + e.EMPLOYEE_NUMBER + "," + ex.Message));
            }
            finally
            {
                //Monitor.Exit(syncroot);
            }
        }

        /// <summary>
        /// 创建AD用户（不再使用）
        /// </summary>
        /// <param name="e"></param>
        public void CreateADUser(EmployeeInfo e)
        {
            try
            {
                string commonName = e.EMPLOYEE_FIRSTNAME + " " + e.EMPLOYEE_LASTNAME;
                if (IsADUserExists(e.EMPLOYEE_NUMBER, commonName) == 0)
                {
                    //创建用户
                    var ouName = GetServerAndOUByEmployee(e);
                    DirectoryEntry de = ADHelper.CreateNewUser(ouName, e.EMPLOYEE_NUMBER, e.EMPLOYEE_NUMBER, "Password01");
                    //修改属性
                    ChangeUserProperty(de, e);
                }
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "CreateADUser," + e.EMPLOYEE_NUMBER + "," + ex.Message));
            }
        }

        public bool CreateADUser(ADUser adUser)
        {
            bool result = true;
            try
            {
                adUser.CommonName = adUser.FirstName + " " + adUser.LastName;
                if (IsADUserExists(adUser.SamAccountName, adUser.CommonName) == 0)
                {
                    //补全用户信息
                    List<ADUser> adUserList = new List<ADUser>() { adUser };
                    ////导入数据表
                    //InsertADUserInfoToDB(adUser);
                    //调用创建AD用户的ps脚本
                    string psFile = ConfigurationManager.AppSettings["ScriptDir"] + "\\" + ConfigurationManager.AppSettings["NewAD1"];
                    //Utilities.ExecuteShell(psFile, adUserList);
                    ////调用创建AD用户邮箱的ps脚本

                    //////创建用户
                    ////var ouName = GetServerAndOUByEmployee(e);
                    ////DirectoryEntry de = ADHelper.CreateNewUser(ouName, e.EMPLOYEE_NUMBER, e.EMPLOYEE_NUMBER, "Password01");
                    //////修改属性
                    ////ChangeUserProperty(de, e);
                }
            }
            catch (Exception ex)
            {
                result = false;
                LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "CreateADUser," + adUser.SamAccountName + "," + ex.Message));
            }
            return result;
        }


        /// <summary>
        /// 创建AD组
        /// </summary>
        public void CreateADGroup()
        {
            //ADHelper.CreateNewGroup("CN=Users", "TestGroup2", "TestGroup2");
            //ADHelper.CreateOrganizeUnit("TestOU2", "");

            //存在则删除所有成员?

            var groupPrefix = ConfigurationManager.AppSettings["GroupPrefix"];
            //1、国家
            var countryGroupOU = ConfigurationManager.AppSettings["CountryGroupOU"];
            if (!ADHelper.ObjectExists(countryGroupOU, "OU"))
            {
                ADHelper.CreateOrganizeUnit(countryGroupOU, "");
            }
            var countrys = GetCountry();
            foreach (var country in countrys)
            {
                var name = groupPrefix + country.Value;
                try
                {
                    if (!ADHelper.ObjectExists(name, "Group"))
                    {
                        ADHelper.CreateNewGroup("OU=" + countryGroupOU, name, name);
                    }
                    else
                    {
                        //
                        //ADHelper.RemoveAllUserFromGroup(name);
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "CreateADGroup,CountryGroup," + name + "," + ex.Message));
                }
            }
            //2、工作城市
            var cityGroup = ConfigurationManager.AppSettings["CityGroupOU"];
            if (!ADHelper.ObjectExists(cityGroup, "OU"))
            {
                ADHelper.CreateOrganizeUnit(cityGroup, "");
            }
            var citys = GetCity();
            foreach (var city in citys)
            {
                var name = groupPrefix + city.Value;
                try
                {
                    if (!ADHelper.ObjectExists(name, "Group"))
                    {
                        ADHelper.CreateNewGroup("OU=" + cityGroup, name, name);
                    }
                    else
                    {
                        //
                        //ADHelper.RemoveAllUserFromGroup(name);
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "CreateADGroup,CityGroup," + name + "," + ex.Message));
                }
            }
            //3、办公地点
            var officeLocationGroupOU = ConfigurationManager.AppSettings["OfficeLocationGroupOU"];
            if (!ADHelper.ObjectExists(officeLocationGroupOU, "OU"))
            {
                ADHelper.CreateOrganizeUnit(officeLocationGroupOU, "");
            }
            var officeLocations = GetOfficeLocation();
            foreach (var officeLocation in officeLocations)
            {
                var name = groupPrefix + officeLocation.Value;
                try
                {
                    if (!ADHelper.ObjectExists(name, "Group"))
                    {
                        ADHelper.CreateNewGroup("OU=" + officeLocationGroupOU, name, name);
                    }
                    else
                    {
                        //
                        //ADHelper.RemoveAllUserFromGroup(name);
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "CreateADGroup,OfficeLocationGroup," + name + "," + ex.Message));
                }
            }
            //4、部门
            CreateDepartmentGroup();
            #region 以前只创建第一级
            //var departmentGroupOU = ConfigurationManager.AppSettings["DepartmentGroupOU"];
            //if (!ADHelper.ObjectExists(departmentGroupOU, "OU"))
            //{
            //    ADHelper.CreateOrganizeUnit(departmentGroupOU, "");
            //}
            //var deptments = GetDepartment();
            //foreach (var deptment in deptments)
            //{
            //    var name = deptment.Value;
            //    try
            //    {
            //        if (!ADHelper.ObjectExists(name, "Group"))
            //        {
            //            ADHelper.CreateNewGroup("OU=" + departmentGroupOU, name, name);
            //        }
            //        else
            //        {
            //            //
            //            //ADHelper.RemoveAllUserFromGroup(name);
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "CreateADGroup,DepartmentGroup," + name + "," + ex.Message));
            //    }
            //}
            #endregion
        }

        public void CreateDepartmentGroup()
        {
            var departmentGroupOU = ConfigurationManager.AppSettings["DepartmentGroupOU"];
            if (!ADHelper.ObjectExists(departmentGroupOU, "OU"))
            {
                ADHelper.CreateOrganizeUnit(departmentGroupOU, "");
            }
            CreateGroup(departmentGroupOU);
        }

        /// <summary>
        /// 组织架构变更，删除现有的DepartmentGroup
        /// </summary>
        /// <returns></returns>
        public bool DeleteDepartmentGroup()
        {
            //登录AD
            var ldap = LDAPStore.FirstOrDefault(l => l.ID == Convert.ToInt32(ConfigurationManager.AppSettings["GroupLDAPID"]));
            if (ldap != null)
            {
                ADHelper.DomainName = ldap.DomainName;
                ADHelper.DomainName = ldap.LDAPDomain;
                ADHelper.ADPath = ldap.ADPath;
                ADHelper.ADUser = ldap.ADUser;
                ADHelper.ADPassword = ldap.ADPassword;
            }

            var departmentGroupOU = ConfigurationManager.AppSettings["DepartmentGroupOU"];
            return ADHelper.DeleteOU(departmentGroupOU, true);
        }


        /// <summary>
        /// 更新AD用户的属性和组（不再使用）
        /// </summary>
        /// <param name="de"></param>
        /// <param name="e"></param>
        public void ChangeUserProperty(DirectoryEntry de, EmployeeInfo e)
        {
            //更新属性
            //ADHelper.SetProperty(de, "displayName", e.EMPLOYEE_LASTNAME + "," + e.EMPLOYEE_FIRSTNAME);//Zhang,San
            if (ADHelper.IsAccExists(e.MANAGER_NUMBER))
            {
                var manager = ADHelper.FindObject("user", e.MANAGER_NUMBER);
                string managerName = manager.Properties["distinguishedName"][0].ToString();
                ADHelper.SetProperty(de, "manager", managerName);
            }
            ADHelper.SetProperty(de, "department", e.DEPARTMENT_NAME);
            ADHelper.SetProperty(de, "company", e.ROOT_DEPARTMENT_NAME);
            ADHelper.SetProperty(de, "title", e.POSITION_EN_NAME);
            ADHelper.SetProperty(de, "physicalDeliveryOfficeName", e.OFFICE_LOCATION);
            ADHelper.SetProperty(de, "telephoneNumber", e.PHONE_NUMBER);
            de.CommitChanges();
            de.Close();

            DirectoryEntry oUser = ADHelper.GetDirectoryEntryByAccount(e.EMPLOYEE_NUMBER);
            LogonPrimaryAD();
            DirectoryEntry oGroup = null;

            //更新组
            var groupPrefix = ConfigurationManager.AppSettings["GroupPrefix"];
            if (!string.IsNullOrEmpty(e.DEPARTMENT_NAME))
            {
                oGroup = ADHelper.GetDirectoryEntryOfGroup(e.DEPARTMENT_NAME);
                ADHelper.AddUserToGroup1(oGroup, oUser);
            }
            //if (!string.IsNullOrEmpty(e.ROOT_DEPARTMENT_NAME))
            //{
            //    oGroup = ADHelper.GetDirectoryEntryOfGroup(e.ROOT_DEPARTMENT_NAME);
            //    ADHelper.AddUserToGroup1(oGroup, oUser);
            //    //ADHelper.AddUserToGroup2(e.EMPLOYEE_NUMBER, e.ROOT_DEPARTMENT_NAME);
            //}
            if (!string.IsNullOrEmpty(e.COUNTRY))
            {
                oGroup = ADHelper.GetDirectoryEntryOfGroup(groupPrefix + e.COUNTRY);
                ADHelper.AddUserToGroup1(oGroup, oUser);
                //ADHelper.AddUserToGroup2(e.EMPLOYEE_NUMBER, groupPrefix + e.COUNTRY);
            }
            if (!string.IsNullOrEmpty(e.COUNTRY) && !string.IsNullOrEmpty(e.LOCATION))
            {
                oGroup = ADHelper.GetDirectoryEntryOfGroup(groupPrefix + e.COUNTRY + "_" + e.LOCATION);
                ADHelper.AddUserToGroup1(oGroup, oUser);
                //ADHelper.AddUserToGroup2(e.EMPLOYEE_NUMBER, groupPrefix + e.COUNTRY + "_" + e.LOCATION);
            }
            if (!string.IsNullOrEmpty(e.OFFICE_LOCATION))
            {
                oGroup = ADHelper.GetDirectoryEntryOfGroup(groupPrefix + e.OFFICE_LOCATION);
                ADHelper.AddUserToGroup1(oGroup, oUser);
                //ADHelper.AddUserToGroup2(e.EMPLOYEE_NUMBER, groupPrefix + e.OFFICE_LOCATION);
            }
        }

        public void ChangeUserProperty(EmployeeInfo e, List<string> whiteNames)
        {
            //更新属性
            string managerDistinguishedName = string.Empty;
            if (!string.IsNullOrEmpty(e.MANAGER_NUMBER) && IsADUserExists(e.MANAGER_NUMBER))
            {
                var manager = ADHelper.FindObject("user", e.MANAGER_NUMBER);
                managerDistinguishedName = manager.Properties["distinguishedName"][0].ToString();
            }
            DirectoryEntry de = new DirectoryEntry();
            if (!IsADUserExists(e.EMPLOYEE_NUMBER))
            {
                //Monitor.Enter(syncroot);
                LogHelper.WriteLog(new LogModel(Level.Info, DateTime.Now, e.EMPLOYEE_NUMBER + ",AD用户不存在"));
                //Monitor.Exit(syncroot);
                return;
            }

            string logDesc = e.EMPLOYEE_NUMBER + " Update log: {";
            de = ADHelper.FindObject("user", e.EMPLOYEE_NUMBER);
            bool isUpdate = false;

            string oldValue = ADHelper.GetProperty(de, "manager");
            if (oldValue != managerDistinguishedName)
            {
                isUpdate = true;
                logDesc += "Manager: " + (string.IsNullOrEmpty(oldValue) ? "-" : oldValue + "(" + managerDistinguishedName + "),");
                ADHelper.SetProperty(de, "manager", managerDistinguishedName);
            }

            oldValue = ADHelper.GetProperty(de, "department");
            if (oldValue != e.DEPARTMENT_NAME)
            {
                isUpdate = true;
                logDesc += "Department: " + (string.IsNullOrEmpty(oldValue) ? "-" : oldValue + "(" + e.DEPARTMENT_NAME + "),");
                ADHelper.SetProperty(de, "department", e.DEPARTMENT_NAME);
            }

            oldValue = ADHelper.GetProperty(de, "company");
            if (oldValue != e.WS_WRK_LOC)
            {
                isUpdate = true;
                logDesc += "Company: " + (string.IsNullOrEmpty(oldValue) ? "-" : oldValue + "(" + e.WS_WRK_LOC + "),");
                ADHelper.SetProperty(de, "company", e.WS_WRK_LOC);
            }
            //ADHelper.SetProperty(de, "company", e.ROOT_DEPARTMENT_NAME);

            oldValue = ADHelper.GetProperty(de, "title");
            if (oldValue != e.POSITION_EN_NAME)
            {
                isUpdate = true;
                logDesc += "Title: " + (string.IsNullOrEmpty(oldValue) ? "-" : oldValue + "(" + e.POSITION_EN_NAME + "),");
                ADHelper.SetProperty(de, "title", e.POSITION_EN_NAME);
            }

            oldValue = ADHelper.GetProperty(de, "physicalDeliveryOfficeName");
            if (oldValue != e.OFFICE_LOCATION)
            {
                isUpdate = true;
                logDesc += "PhysicalDeliveryOfficeName: " + (string.IsNullOrEmpty(oldValue) ? "-" : oldValue + "(" + e.OFFICE_LOCATION + "),");
                ADHelper.SetProperty(de, "physicalDeliveryOfficeName", e.OFFICE_LOCATION);
            }

            oldValue = ADHelper.GetProperty(de, "mobile");
            if (oldValue != e.PHONE_NUMBER)
            {
                isUpdate = true;
                logDesc += "Mobile: " + (string.IsNullOrEmpty(oldValue) ? "-" : oldValue + "(" + e.PHONE_NUMBER + "),");
                ADHelper.SetProperty(de, "mobile", e.PHONE_NUMBER);
            }
            //ADHelper.SetProperty(de, "telephoneNumber", e.PHONE_NUMBER);

            ////如果不是白名单中的用户，则更新显示名
            //if (!whiteNames.Contains(e.EMPLOYEE_NUMBER.ToLower()))
            //{
            //    string oldDisplayName = ADHelper.GetProperty(de, "displayName");
            //    oldValue = oldDisplayName;
            //    string pattern = @"\(.+/.+\)";
            //    //如果是带括号的显示名，则更新
            //    if (System.Text.RegularExpressions.Regex.IsMatch(oldDisplayName, pattern))
            //    {
            //        if (LocationMappingStore == null)
            //            LocationMappingStore = GetLocationMappingStore();
            //        LocationMapping entity = LocationMappingStore.FirstOrDefault(s => s.WorkLocation.ToUpper() == e.LOCATION.ToUpper());
            //        string workLocation = string.Empty;
            //        if (entity != null)
            //            workLocation = entity.Region;

            //        if (!string.IsNullOrEmpty(workLocation))
            //        {
            //            string oldMatchDepart = System.Text.RegularExpressions.Regex.Match(oldDisplayName, pattern).Value;
            //            string[] departmentNames = e.DEPARTMENT_NAME.Split(new char[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            //            bool isChange = true;
            //            string newMatchDepart = string.Empty;
            //            //循环查看，只要有相同的，就不用更新
            //            for (int i = 0; i < departmentNames.Length; i++)
            //            {
            //                newMatchDepart = "(" + departmentNames[i] + "/" + workLocation + ")";
            //                if (oldMatchDepart.Replace(" ", "").ToLower() == newMatchDepart.Replace(" ", "").ToLower())
            //                {
            //                    isChange = false;
            //                    break;
            //                }
            //            }
            //            //部门或者工作地是否更改，如果更改，则更新
            //            if (isChange)
            //            {
            //                string firstNameLastName = oldDisplayName.Substring(0, oldDisplayName.IndexOf("("));
            //                string newDisplayName = string.Empty;
            //                bool isexist = false;
            //                for (int i = 0; i < departmentNames.Length; i++)
            //                {
            //                    newDisplayName = firstNameLastName + "(" + departmentNames[i] + "/" + workLocation + ")";
            //                    isexist = IsADUserExistsByDisplayName(newDisplayName);
            //                    if (!isexist)
            //                        break;
            //                }
            //                int num = 0;
            //                string newDisplayNameNum = newDisplayName;
            //                while (isexist)
            //                {
            //                    num++;
            //                    newDisplayNameNum = newDisplayName + num.ToString();
            //                    isexist = IsADUserExistsByDisplayName(newDisplayNameNum);
            //                }
            //                logDesc += "Displayname: " + (string.IsNullOrEmpty(oldValue) ? "-" : oldValue + "(" + newDisplayNameNum + ")");
            //                ADHelper.SetProperty(de, "displayname", newDisplayNameNum);
            //                isUpdate = true;
            //            }
            //        }
            //    }
            //}

            if (isUpdate)
            {
                de.CommitChanges();
                logDesc = (logDesc.Substring(logDesc.Length - 1, 1).Equals(",") ? logDesc.Substring(0, logDesc.Length - 1) : logDesc) + "}";
                LogHelper.WriteLog(new LogModel(Level.Info, DateTime.Now, logDesc));
            }
            de.Close();

            DirectoryEntry oUser = ADHelper.GetDirectoryEntryByAccount(e.EMPLOYEE_NUMBER);
            LogonPrimaryAD();
            DirectoryEntry oGroup = null;

            //更新组
            var groupPrefix = ConfigurationManager.AppSettings["GroupPrefix"];
            if (!string.IsNullOrEmpty(e.DEPARTMENT_NAME))
            {
                string depName = string.Empty;
                string[] depNames = e.DEPARTMENT_NAME.Split(new char[] { '_' });
                for (int i = 0; i < depNames.Length; i++)
                {
                    if (i == 0)
                        depName = depNames[i];
                    else
                        depName = depName + "_" + depName[i];
                    oGroup = ADHelper.GetDirectoryEntryOfGroup(depName);
                    ADHelper.AddUserToGroup1(oGroup, oUser);
                }
            }
            //if (!string.IsNullOrEmpty(e.ROOT_DEPARTMENT_NAME))
            //{
            //    oGroup = ADHelper.GetDirectoryEntryOfGroup(e.ROOT_DEPARTMENT_NAME);
            //    ADHelper.AddUserToGroup1(oGroup, oUser);
            //    //ADHelper.AddUserToGroup2(e.EMPLOYEE_NUMBER, e.ROOT_DEPARTMENT_NAME);
            //}
            if (!string.IsNullOrEmpty(e.COUNTRY))
            {
                oGroup = ADHelper.GetDirectoryEntryOfGroup(groupPrefix + e.COUNTRY);
                ADHelper.AddUserToGroup1(oGroup, oUser);
                //ADHelper.AddUserToGroup2(e.EMPLOYEE_NUMBER, groupPrefix + e.COUNTRY);
            }
            if (!string.IsNullOrEmpty(e.COUNTRY) && !string.IsNullOrEmpty(e.LOCATION))
            {
                oGroup = ADHelper.GetDirectoryEntryOfGroup(groupPrefix + e.COUNTRY + "_" + e.LOCATION);
                ADHelper.AddUserToGroup1(oGroup, oUser);
                //ADHelper.AddUserToGroup2(e.EMPLOYEE_NUMBER, groupPrefix + e.COUNTRY + "_" + e.LOCATION);
            }
            if (!string.IsNullOrEmpty(e.OFFICE_LOCATION))
            {
                oGroup = ADHelper.GetDirectoryEntryOfGroup(groupPrefix + e.OFFICE_LOCATION);
                ADHelper.AddUserToGroup1(oGroup, oUser);
                //ADHelper.AddUserToGroup2(e.EMPLOYEE_NUMBER, groupPrefix + e.OFFICE_LOCATION);
            }
        }

        /// <summary>
        /// 2016.3.16
        /// 同 PowerShellAD.cs 中的 GetLocationMappingStore() 方法。
        /// 从116的ADDB数据库中的T_LocationMapping表获取WorkLocation字段
        /// </summary>
        /// <returns></returns>
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

        private List<string> ReadWhiteNameList()
        {
            List<string> whiteNames = new List<string>();
            string txtNames = string.Empty;
            try
            {
                txtNames = System.IO.File.ReadAllText("WhiteNameList.txt");
            }
            catch
            { }

            if (!string.IsNullOrEmpty(txtNames))
            {
                whiteNames = txtNames.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                whiteNames.ForEach(c => c.ToLower());
            }

            return whiteNames;
        }

        /// <summary>
        /// 获取员工所属的AD Server(并登陆)和OU
        /// AD中不存在该用户
        /// </summary>
        /// <param name="empno"></param>
        /// <returns></returns>
        public string GetServerAndOUByEmployee(EmployeeInfo e)
        {
            var location = GetEmployeeLocation(e);
            var mapping = LDAPLocationMappingStore.FirstOrDefault(m => m.Location == location);
            var ldap = LDAPStore.FirstOrDefault(l => l.ID == mapping.LDAPID);

            ADHelper.DomainName = ldap.DomainName;
            ADHelper.LDAPDomain = ldap.LDAPDomain; //ADHelper.DomainName = ldap.LDAPDomain;
            ADHelper.ADPath = ldap.ADPath;
            ADHelper.ADUser = ldap.ADUser;
            ADHelper.ADPassword = ldap.ADPassword;

            return mapping.OU;
        }

        /// <summary>
        /// 获取员工Location
        /// </summary>
        /// <param name="e"></param>
        /// <returns></returns>
        public string GetEmployeeLocation(EmployeeInfo e)
        {
            string location = string.Empty;
            if (e.COUNTRY == "CHN" || e.COUNTRY == "HKG" || e.COUNTRY == "MAC" || e.COUNTRY == "TWN")
            {
                //国内
                var list = new List<string>() { "BEIJING", "DALIAN", "WUXI", "SHENZHEN", "SHANGHAI" };
                location = list.Contains(e.LOCATION) ? e.LOCATION : "Others";
            }
            else
            {
                //海外
                location = "Oversea";
            }
            return location;
        }

        #region XML Method

        ///// <summary>
        ///// 从XML获取LDAP数据
        ///// </summary>
        ///// <param name="path"></param>
        ///// <returns></returns>
        //public static List<LDAP> GetLADPStore(string path)
        //{
        //    XElement xe = XElement.Load(path);
        //    IEnumerable<XElement> elements = from ele in xe.Elements("LDAP")
        //                                     select ele;
        //    List<LDAP> ldaps = new List<LDAP>();
        //    foreach (var ele in elements)
        //    {
        //        LDAP ldap = new LDAP();
        //        var id = ele.Element("ID");
        //        if (id != null)
        //            ldap.ID = Convert.ToInt32(id.Value);
        //        var ladpDomain = ele.Element("LDAPDomain");
        //        if (ladpDomain != null)
        //            ldap.LDAPDomain = ladpDomain.Value;
        //        var domainName = ele.Element("DomainName");
        //        if (domainName != null)
        //            ldap.DomainName = domainName.Value;
        //        var adPath = ele.Element("ADPath");
        //        if (adPath != null)
        //            ldap.ADPath = adPath.Value;
        //        var adUser = ele.Element("ADUser");
        //        if (adUser != null)
        //            ldap.ADUser = adUser.Value;
        //        var adPassword = ele.Element("ADPassword");
        //        if (adPassword != null)
        //            ldap.ADPassword = adPassword.Value;
        //        ldaps.Add(ldap);
        //    }
        //    return ldaps;
        //}

        ///// <summary>
        ///// 从XML获取LDAPLocationMapping数据
        ///// </summary>
        ///// <param name="path"></param>
        ///// <returns></returns>
        //public static List<LDAPLocationMapping> GetLDAPLocationMappingStore(string path)
        //{
        //    XElement xe = XElement.Load(path);
        //    IEnumerable<XElement> elements = from ele in xe.Elements("Mapping")
        //                                     select ele;
        //    List<LDAPLocationMapping> mappings = new List<LDAPLocationMapping>();
        //    foreach (var ele in elements)
        //    {
        //        LDAPLocationMapping mapping = new LDAPLocationMapping();
        //        var ldapID = ele.Element("LDAPID");
        //        if (ldapID != null)
        //            mapping.LDAPID = Convert.ToInt32(ldapID.Value);
        //        var ou = ele.Element("OU");
        //        if (ou != null)
        //            mapping.OU = ou.Value;
        //        var location = ele.Element("Location");
        //        if (location != null)
        //            mapping.Location = location.Value;
        //        mappings.Add(mapping);
        //    }
        //    return mappings;
        //}

        #endregion

        #region DB Method

        /// <summary>
        /// 获取员工列表
        /// </summary>
        /// <returns></returns>
        public List<EmployeeInfo> GetEmployeeInfoList()
        {
            var list = new List<EmployeeInfo>();
            //26517
            string strSql = @"select isnull(a.EMPLOYEE_NUMBER,'') as EMPLOYEE_NUMBER
                            ,isnull(a.EMPLOYEE_FIRSTNAME,'') as EMPLOYEE_FIRSTNAME
                            ,isnull(a.EMPLOYEE_MIDDLENAME,'') as EMPLOYEE_MIDDLENAME
                            ,isnull(a.EMPLOYEE_LASTNAME,'') as EMPLOYEE_LASTNAME
                            ,isnull(a.E_MAIL,'') as E_MAIL
                            ,isnull(a.PHONE_NUMBER,'') as PHONE_NUMBER
                            ,isnull(a.MOBILE_NUMBER,'') as MOBILE_NUMBER
                            ,isnull(a.OFFICE_LOCATION,'') as OFFICE_LOCATION
                            ,isnull(a.COUNTRY,'') as COUNTRY
                            ,isnull(c.DEPARTMENT_NO,'') as DEPARTMENT_NO
                            ,isnull(c.DEPARTMENT_CN_NAME,'') as DEPARTMENT_CN_NAME
                            ,isnull(c.ROOT_DEPARTMENT_NUMBER,'') as ROOT_DEPARTMENT_NUMBER
                            ,isnull(c.ROOT_DEPARTMENT_NAME ,'') as ROOT_DEPARTMENT_NAME
                            ,isnull(d.LOCATION,'') as LOCATION
                            ,isnull(e.MANAGER_NUMBER,'') as MANAGER_NUMBER
                            ,isnull(e.MANAGER_NAME,'') as MANAGER_NAME
                            ,isnull(f.GRADE_NAME,'') as GRADE_NAME 
                            ,isnull(f.POSITION_EN_NAME,'') as POSITION_EN_NAME 
                            ,isnull(g.WS_WRK_LOC,'') as WS_WRK_LOC 
                            from [dbo].[V_Employee_Info] a WITH(NOLOCK)
                            left join [dbo].[V_Employee_Department] b WITH(NOLOCK) on a.EMPLOYEE_NUMBER Collate Database_Default =b.EMPLOYEE_NUMBER
                            left join [dbo].[V_Department_Info] c WITH(NOLOCK) on b.DEPARTMENT_ID Collate Database_Default =c.DEPARTMENT_ID
                            left join [dbo].[V_Employee_Location] d WITH(NOLOCK) on a.EMPLOYEE_NUMBER Collate Database_Default =d.EMPLOYEE_NUMBER
                            left join [dbo].[V_Employee_Manager] e WITH(NOLOCK) on a.EMPLOYEE_NUMBER Collate Database_Default =e.EMPLOYEE_NUMBER
                            left join [dbo].[V_Employee_Position] f WITH(NOLOCK) on a.EMPLOYEE_NUMBER Collate Database_Default =f.EMPLOYEE_NUMBER 
                            left join [dbo].[V_Employee_Cnt_Location] g WITH(NOLOCK) on a.EMPLOYEE_NUMBER Collate Database_Default =g.EMPLOYEE_NUMBER
                            where a.EMPLOYEE_STATUS='A'";
            var departmentFilter = ConfigurationManager.AppSettings["DepartmentFilter"];
            if (!string.IsNullOrEmpty(departmentFilter))
            {
                string[] departments = departmentFilter.Split(new char[] { ',' });
                strSql += "and (";
                foreach (string department in departments)
                {
                    strSql += " c.DEPARTMENT_CN_NAME like '%" + department + "%' or";
                }
                strSql = strSql.Substring(0, strSql.Length - 2) + ")";
            }
            var dr = SqlHelper.ExecuteReader(Conn, CommandType.Text, strSql);
            while (dr.Read())
            {
                var e = new EmployeeInfo();
                e.EMPLOYEE_NUMBER = dr["EMPLOYEE_NUMBER"].ToString().Trim();
                e.EMPLOYEE_FIRSTNAME = dr["EMPLOYEE_FIRSTNAME"].ToString();
                e.EMPLOYEE_MIDDLENAME = dr["EMPLOYEE_MIDDLENAME"].ToString();
                e.EMPLOYEE_LASTNAME = dr["EMPLOYEE_LASTNAME"].ToString();
                e.E_MAIL = dr["E_MAIL"].ToString();
                e.OFFICE_LOCATION = dr["OFFICE_LOCATION"].ToString();
                e.COUNTRY = dr["COUNTRY"].ToString();
                e.PHONE_NUMBER = dr["PHONE_NUMBER"].ToString();
                e.MOBILE_NUMBER = dr["MOBILE_NUMBER"].ToString();
                e.DEPARTMENT_NO = dr["DEPARTMENT_NO"].ToString().Trim();
                e.DEPARTMENT_NAME = dr["DEPARTMENT_CN_NAME"].ToString();
                e.ROOT_DEPARTMENT_NUMBER = dr["ROOT_DEPARTMENT_NUMBER"].ToString().Trim();
                e.ROOT_DEPARTMENT_NAME = dr["ROOT_DEPARTMENT_NAME"].ToString();
                e.LOCATION = dr["LOCATION"].ToString();
                e.MANAGER_NUMBER = dr["MANAGER_NUMBER"].ToString().Trim();
                e.MANAGER_NAME = dr["MANAGER_NAME"].ToString();
                e.GRADE_NAME = dr["GRADE_NAME"].ToString();
                e.POSITION_EN_NAME = dr["POSITION_EN_NAME"].ToString();
                e.WS_WRK_LOC = dr["WS_WRK_LOC"].ToString();
                list.Add(e);
            }
            return list;
        }

        /// <summary>
        /// 获取员工信息
        /// </summary>
        /// <param name="eno"></param>
        /// <returns></returns>
        public EmployeeInfo GetEmployeeInfoByID(string emo)
        {
            string strSql = @"select isnull(a.EMPLOYEE_NUMBER,'') as EMPLOYEE_NUMBER
                            ,isnull(a.EMPLOYEE_FIRSTNAME,'') as EMPLOYEE_FIRSTNAME
                            ,isnull(a.EMPLOYEE_MIDDLENAME,'') as EMPLOYEE_MIDDLENAME
                            ,isnull(a.EMPLOYEE_LASTNAME,'') as EMPLOYEE_LASTNAME
                            ,isnull(a.E_MAIL,'') as E_MAIL
                            ,isnull(a.PHONE_NUMBER,'') as PHONE_NUMBER
                            ,isnull(a.MOBILE_NUMBER,'') as MOBILE_NUMBER
                            ,isnull(a.OFFICE_LOCATION,'') as OFFICE_LOCATION
                            ,isnull(a.COUNTRY,'') as COUNTRY
                            ,isnull(c.DEPARTMENT_NO,'') as DEPARTMENT_NO
                            ,isnull(c.DEPARTMENT_CN_NAME,'') as DEPARTMENT_CN_NAME
                            ,isnull(c.ROOT_DEPARTMENT_NUMBER,'') as ROOT_DEPARTMENT_NUMBER
                            ,isnull(c.ROOT_DEPARTMENT_NAME ,'') as ROOT_DEPARTMENT_NAME
                            ,isnull(d.LOCATION,'') as LOCATION
                            ,isnull(e.MANAGER_NUMBER,'') as MANAGER_NUMBER
                            ,isnull(e.MANAGER_NAME,'') as MANAGER_NAME
                            ,isnull(f.GRADE_NAME,'') as GRADE_NAME 
                            ,isnull(g.WS_WRK_LOC,'') as WS_WRK_LOC 
                            from [dbo].[V_Employee_Info] a WITH(NOLOCK)
                            left join [dbo].[V_Employee_Department] b WITH(NOLOCK) on a.EMPLOYEE_NUMBER Collate Database_Default =b.EMPLOYEE_NUMBER
                            left join [dbo].[V_Department_Info] c WITH(NOLOCK) on b.DEPARTMENT_ID Collate Database_Default =c.DEPARTMENT_ID
                            left join [dbo].[V_Employee_Location] d WITH(NOLOCK) on a.EMPLOYEE_NUMBER Collate Database_Default =d.EMPLOYEE_NUMBER
                            left join [dbo].[V_Employee_Manager] e WITH(NOLOCK) on a.EMPLOYEE_NUMBER Collate Database_Default =e.EMPLOYEE_NUMBER
                            left join [dbo].[V_Employee_Position] f WITH(NOLOCK) on a.EMPLOYEE_NUMBER Collate Database_Default =f.EMPLOYEE_NUMBER 
                            left join [dbo].[V_Employee_Cnt_Location] g WITH(NOLOCK) on a.EMPLOYEE_NUMBER Collate Database_Default =g.EMPLOYEE_NUMBER
                            where a.EMPLOYEE_NUMBER='" + emo + "'";
            var dr = SqlHelper.ExecuteReader(Conn, CommandType.Text, strSql);
            var e = new EmployeeInfo();
            if (dr.Read())
            {
                e.EMPLOYEE_NUMBER = dr["EMPLOYEE_NUMBER"].ToString().Trim();
                e.EMPLOYEE_FIRSTNAME = dr["EMPLOYEE_FIRSTNAME"].ToString();
                e.EMPLOYEE_MIDDLENAME = dr["EMPLOYEE_MIDDLENAME"].ToString();
                e.EMPLOYEE_LASTNAME = dr["EMPLOYEE_LASTNAME"].ToString();
                e.E_MAIL = dr["E_MAIL"].ToString();
                e.OFFICE_LOCATION = dr["OFFICE_LOCATION"].ToString();
                e.COUNTRY = dr["COUNTRY"].ToString();
                e.PHONE_NUMBER = dr["PHONE_NUMBER"].ToString();
                e.MOBILE_NUMBER = dr["MOBILE_NUMBER"].ToString();
                e.DEPARTMENT_NO = dr["DEPARTMENT_NO"].ToString().Trim();
                e.DEPARTMENT_NAME = dr["DEPARTMENT_CN_NAME"].ToString();
                e.ROOT_DEPARTMENT_NUMBER = dr["ROOT_DEPARTMENT_NUMBER"].ToString().Trim();
                e.ROOT_DEPARTMENT_NAME = dr["ROOT_DEPARTMENT_NAME"].ToString();
                e.LOCATION = dr["LOCATION"].ToString();
                e.MANAGER_NUMBER = dr["MANAGER_NUMBER"].ToString().Trim();
                e.MANAGER_NAME = dr["MANAGER_NAME"].ToString();
                e.GRADE_NAME = dr["GRADE_NAME"].ToString();
                e.POSITION_EN_NAME = dr["POSITION_EN_NAME"].ToString();
                e.WS_WRK_LOC = dr["WS_WRK_LOC"].ToString();
            }
            return e;
        }


        /// <summary>
        /// 获取部门
        /// </summary>
        /// <returns></returns>
        public List<OU> GetDepartment()
        {
            var list = new List<OU>();
            string strSql = @"select distinct [DEPARTMENT_EN_NAME] from [dbo].[V_Department_Info] where [IS_EFFECTIVE]=1 and len([DEPARTMENT_NO])=4";
            var dr = SqlHelper.ExecuteReader(Conn, CommandType.Text, strSql);
            while (dr.Read())
            {
                var ou = new OU();
                ou.Value = dr["DEPARTMENT_EN_NAME"].ToString();
                list.Add(ou);
            }
            return list;
        }
        /// <summary>
        /// 根据部门级别获取部门
        /// </summary>
        /// <param name="departmentNoLen">（4,6,8,10）</param>
        /// <returns></returns>
        public List<OU> GetDepartment(int departmentNoLen)
        {
            var list = new List<OU>();
            string strSql = @"select distinct [DEPARTMENT_EN_NAME] from [dbo].[V_Department_Info] where [IS_EFFECTIVE]=1 and len([DEPARTMENT_NO])=" + departmentNoLen;
            var dr = SqlHelper.ExecuteReader(Conn, CommandType.Text, strSql);
            while (dr.Read())
            {
                var ou = new OU();
                ou.Value = dr["DEPARTMENT_EN_NAME"].ToString();
                list.Add(ou);
            }
            return list;
        }
        /// <summary>
        /// 创建级别OU以及组
        /// </summary>
        /// <param name="level">(1,2,3,4)</param>
        /// <param name="departmentGroupOU"></param>
        public void CreateLnOUAndGroup(int level, string departmentGroupOU)
        {
            string ln = "L" + level;
            DirectoryEntry ou = null;
            if (!ADHelper.ObjectExists(ln, "OU"))
            {
                ou = ADHelper.CreateOrganizeUnit(ln, departmentGroupOU);
            }
            
            List<OU> departments = GetDepartment(4 + 2 * (level - 1));
            foreach (OU department in departments)
            {
                try
                {
                    string ldapDN = "OU=" + ln + ",OU=" + departmentGroupOU;
                    if (!ADHelper.ObjectExists(department.Value, "Group"))
                    {
                        //如果连接到的AD是在服务器上那么格式写成LDAP:\\XX.XX.XX.XX\OU=XX部门,OU=XX公司,DC=域名,DC=COM（XX.XX.XX.XX为服务器IP）；
                        ADHelper.CreateNewGroup(ldapDN, department.Value, department.Value);
                    }
                    else
                    {
                        if (ou == null)
                            ou = ADHelper.FindObject("OU", ln);
                        //ADHelper.RemoveAllUserFromGroup(department.Value);
                        DirectoryEntry oGroup = ADHelper.GetDirectoryEntryOfGroup(department.Value);
                        oGroup.MoveTo(ou);
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLog(new LogModel(Level.Error, DateTime.Now, "CreateADGroup,DepartmentGroup," + department.Value + "," + ex.Message));
                }
            }
        }
        public void CreateGroup(string departmentGroupOU = "DepartmentGroup")
        {
            CreateLnOUAndGroup(1, departmentGroupOU); //BG3
            CreateLnOUAndGroup(2, departmentGroupOU); //BG3_BUINT
            CreateLnOUAndGroup(3, departmentGroupOU); //BG3_BUINT_BJ
            CreateLnOUAndGroup(4, departmentGroupOU); //BG3_BUINT_BJ_OTHERS
        }

        /// <summary>
        /// 获取办公地点
        /// </summary>
        /// <returns></returns>
        public List<OU> GetOfficeLocation()
        {
            var list = new List<OU>();
            string strSql = @"select distinct [OFFICE_LOCATION_NAME] from [dbo].[PSFT_V_Office_Location]";
            var dr = SqlHelper.ExecuteReader(Conn, CommandType.Text, strSql);
            while (dr.Read())
            {
                var ou = new OU();
                ou.Value = dr["OFFICE_LOCATION_NAME"].ToString();
                list.Add(ou);
            }
            return list;
        }

        /// <summary>
        /// 获取工作城市
        /// </summary>
        /// <returns></returns>
        public List<OU> GetCity()
        {
            var list = new List<OU>();
            string strSql = @"select distinct [COUNTRY_CODE]+'_'+[WORK_LOCATION] as [WORK_LOCATION] from [HRPRD].[dbo].[PSFT_V_Work_Location]";
            var dr = SqlHelper.ExecuteReader(Conn, CommandType.Text, strSql);
            while (dr.Read())
            {
                var ou = new OU();
                ou.Value = dr["WORK_LOCATION"].ToString();
                list.Add(ou);
            }
            return list;
        }

        /// <summary>
        /// 获取国家
        /// </summary>
        /// <returns></returns>
        public List<OU> GetCountry()
        {
            var list = new List<OU>();
            string strSql = @"select distinct [COUNTRY_ID] from [HRPRD].[dbo].[PSFT_V_COUNTRY]";
            var dr = SqlHelper.ExecuteReader(Conn, CommandType.Text, strSql);
            while (dr.Read())
            {
                var ou = new OU();
                ou.Value = dr["COUNTRY_ID"].ToString();
                list.Add(ou);
            }
            return list;
        }

        #endregion

        /// <summary>
        /// 导出AD用户信息到Excel
        /// </summary>
        /// <param name="adUser"></param>
        /// <returns></returns>
        private string ADUserInfoToExcel(ADUser adUser)
        {
            string fileFullName = ConfigurationManager.AppSettings["ScriptDir"] + "Mingming.xlsx";
            fileFullName = ExportHelper.CreateExcel(fileFullName, null, false);
            List<ADUser> objects = new List<ADUser>();
            objects.Add(adUser);
            ExportHelper.WriteExcel<ADUser>(fileFullName, objects, null, false);
            return fileFullName;
        }
    }

    public class MyParameter
    {
        public MyParameter() { }
        public MyParameter(EmployeeInfo employeeInfo, AutoResetEvent are)
        {
            this.EmployeeInfo = employeeInfo;
            this.Are = are;
        }
        public EmployeeInfo EmployeeInfo { get; set; }
        public AutoResetEvent Are { get; set; }
    }
}
