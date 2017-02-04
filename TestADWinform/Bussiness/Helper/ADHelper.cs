﻿using System;
using System.Collections.Generic;
using System.Text;
using System.DirectoryServices;
using System.Security.Principal;
using System.Runtime.InteropServices;
using System.Data;

namespace Bussiness
{
    public class ADHelper
    {
        #region 私有变量

        /// <summary>
        /// homeMTA
        /// </summary>
        public static string homeMTA = ""; //请填写自己的环境变量
        /// <summary>
        /// homeMDB
        /// </summary>
        public static string homeMDB = ""; //请填写自己的环境变量
        /// <summary>
        /// msExchHomeServerName
        /// </summary>
        public static string msExchHomeServerName = ""; //请填写自己的环境变量

        /// <summary>
        /// 域名
        /// </summary>
        public static string DomainName = "testcode.com";
        /// <summary>
        /// LDAP 地址
        /// </summary>
        public static string LDAPDomain = "dc=testcode,dc=com";
        /// <summary>
        /// LDAP绑定路径
        /// </summary>
        public static string ADPath;
        public static string sPrincpleNameTail = "@cinf.com";
        /// <summary>
        /// 登录帐号
        /// </summary>
        public static string ADUser;
        /// <summary>
        /// 登录密码
        /// </summary>
        public static string ADPassword;
        //public static string ADPassword = "";

        #endregion

        #region 枚举常量

        /// <summary>
        /// 用户登录验证结果
        /// </summary>
        public enum LoginResult
        {
            /// 
            /// 正常登录
            /// 
            LOGIN_USER_OK = 0,
            /// 
            /// 用户不存在
            /// 
            LOGIN_USER_DOESNT_EXIST,
            /// 
            /// 用户帐号被禁用
            /// 
            LOGIN_USER_ACCOUNT_INACTIVE,
            /// 
            /// 用户密码不正确
            /// 
            LOGIN_USER_PASSWORD_INCORRECT
        }
        /// <summary>
        /// 用户属性定义标志
        /// </summary>
        public enum ADS_USER_FLAG_ENUM
        {
            /// 
            /// 登录脚本标志。如果通过 ADSI LDAP 进行读或写操作时，
            /// 该标志失效。如果通过 ADSI WINNT，该标志为只读。
            /// 
            ADS_UF_SCRIPT = 0X0001,
            /// 
            /// 用户帐号禁用标志
            /// 
            ADS_UF_ACCOUNTDISABLE = 0X0002,
            /// 
            /// 主文件夹标志
            /// 
            ADS_UF_HOMEDIR_REQUIRED = 0X0008,
            /// 
            /// 过期标志
            /// 
            ADS_UF_LOCKOUT = 0X0010,
            /// 
            /// 用户密码不是必须的
            /// 
            ADS_UF_PASSWD_NOTREQD = 0X0020,
            /// 
            /// 密码不能更改标志
            /// 
            ADS_UF_PASSWD_CANT_CHANGE = 0X0040,
            /// 
            /// 使用可逆的加密保存密码
            /// 
            ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = 0X0080,
            /// 
            /// 本地帐号标志
            /// 
            ADS_UF_TEMP_DUPLICATE_ACCOUNT = 0X0100,
            /// 
            /// 普通用户的默认帐号类型
            /// 
            ADS_UF_NORMAL_ACCOUNT = 0X0200,
            /// 
            /// 跨域的信任帐号标志
            /// 
            ADS_UF_INTERDOMAIN_TRUST_ACCOUNT = 0X0800,
            /// 
            /// 工作站信任帐号标志
            /// 
            ADS_UF_WORKSTATION_TRUST_ACCOUNT = 0x1000,
            /// 
            /// 服务器信任帐号标志
            /// 
            ADS_UF_SERVER_TRUST_ACCOUNT = 0X2000,
            /// 
            /// 密码永不过期标志
            /// 
            ADS_UF_DONT_EXPIRE_PASSWD = 0X10000,
            /// 
            /// MNS 帐号标志
            /// 
            ADS_UF_MNS_LOGON_ACCOUNT = 0X20000,
            /// 
            /// 交互式登录必须使用智能卡
            /// 
            ADS_UF_SMARTCARD_REQUIRED = 0X40000,
            /// 
            /// 当设置该标志时，服务帐号（用户或计算机帐号）将通过 Kerberos 委托信任
            /// 
            ADS_UF_TRUSTED_FOR_DELEGATION = 0X80000,
            /// 
            /// 当设置该标志时，即使服务帐号是通过 Kerberos 委托信任的，敏感帐号不能被委托
            /// 
            ADS_UF_NOT_DELEGATED = 0X100000,
            /// 
            /// 此帐号需要 DES 加密类型
            /// 
            ADS_UF_USE_DES_KEY_ONLY = 0X200000,
            /// 
            /// 不要进行 Kerberos 预身份验证
            /// 
            ADS_UF_DONT_REQUIRE_PREAUTH = 0X4000000,
            /// 
            /// 用户密码过期标志
            /// 
            ADS_UF_PASSWORD_EXPIRED = 0X800000,
            /// 
            /// 用户帐号可委托标志
            /// 
            ADS_UF_TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = 0X1000000

        }

        #endregion

        #region 构造函数

        public ADHelper()
        {
            //
            // System.Environment.UserName
            //
        }

        /// <summary>
        /// 多联的AD构造函数
        /// </summary>
        /// <param name="sADPath"></param>
        /// <param name="sDomainName"></param>
        /// <param name="sADUser"></param>
        /// <param name="sADUserPWD"></param>
        public ADHelper(string sADPath, string sDomainName, string sADUser, string sADUserPWD)
        {
            ADPath = sADPath;
            DomainName = sDomainName;
            ADUser = sADUser;
            ADPassword = sADUserPWD;
        }

        #endregion

        #region GetDirectoryObject

        /// <summary>
        /// 获得DirectoryEntry对象实例,以管理员登陆AD
        /// </summary>
        /// <returns></returns>
        public static DirectoryEntry GetDirectoryObject()
        {
            DirectoryEntry entry = new DirectoryEntry(ADPath, ADUser, ADPassword, AuthenticationTypes.Secure);
            return entry;
        }

        /// <summary>
        /// 根据指定用户名和密码获得相应DirectoryEntry实体
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public static DirectoryEntry GetDirectoryObject(string userName, string password)
        {
            DirectoryEntry entry = new DirectoryEntry(ADPath,
                userName, password, AuthenticationTypes.None);
            return entry;
        }

        /// <summary>
        /// i.e. /CN=Users,DC=creditsights, DC=cyberelves, DC=Com
        /// </summary>
        /// <param name="domainReference"></param>
        /// <returns></returns>
        public static DirectoryEntry GetDirectoryObject(string domainReference)
        {
            DirectoryEntry entry = new DirectoryEntry(ADPath + domainReference, ADUser, ADPassword,
                AuthenticationTypes.Secure);
            return entry;
        }

        /// <summary>
        ///  获得以UserName,Password创建的DirectoryEntry
        /// </summary>
        /// <param name="domainReference"></param>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public static DirectoryEntry GetDirectoryObject(string domainReference,
            string userName, string password)
        {
            DirectoryEntry entry = new DirectoryEntry(ADPath + domainReference,
                userName, password, AuthenticationTypes.Secure);
            return entry;
        }

        #endregion

        #region GetDirectoryEntry

        /// <summary>
        /// 根据用户公共名称取得用户的 对象
        /// </summary>
        /// <param name="commonName">用户公共名称</param>
        /// <returns>如果找到该用户，则返回用户的 对象；否则返回 null</returns>
        public static DirectoryEntry GetDirectoryEntry(string commonName)
        {
            DirectoryEntry de = GetDirectoryObject();
            DirectorySearcher deSearch = new DirectorySearcher(de);
            deSearch.Filter = "(&(&(objectCategory=person)(objectClass=user))(cn=" +
commonName + "))";
            deSearch.SearchScope = SearchScope.Subtree;

            SearchResult result = deSearch.FindOne();
            if (result != null)
            {
                de = new DirectoryEntry(result.Path);
                de.Username = ADUser;
                de.Password = ADPassword;
            }
            else
            {
                de = null;
            }
            return de;
        }
        public static DirectoryEntry GetDirectoryEntry(string commonName, string adPath, string adUser, string adUserPwd)
        {
            DirectoryEntry de = new DirectoryEntry(adPath, adUser, adUserPwd, AuthenticationTypes.Secure);
            DirectorySearcher deSearch = new DirectorySearcher(de);
            deSearch.Filter = "(&(&(objectCategory=person)(objectClass=user))(cn=" +
commonName + "))";
            deSearch.SearchScope = SearchScope.Subtree;

            SearchResult result = deSearch.FindOne();
            if (result != null)
            {
                de = new DirectoryEntry(result.Path);
                de.Username = ADUser;
                de.Password = ADPassword;
            }
            else
            {
                de = null;
            }
            return de;
        }

        /// <summary>
        /// 根据用户公共名称和密码取得用户的 对象。
        /// </summary>
        /// <param name="commonName">用户公共名称</param>
        /// <param name="password">用户密码</param>
        /// <returns>如果找到该用户，则返回用户的 对象；否则返回 null</returns>
        public static DirectoryEntry GetDirectoryEntry(string commonName, string password)
        {
            DirectoryEntry de = GetDirectoryObject(commonName, password);
            DirectorySearcher deSearch = new DirectorySearcher(de);
            deSearch.Filter = "(&(&(objectCategory=person)(objectClass=user))(cn=" + commonName + "))";
            deSearch.SearchScope = SearchScope.Subtree;
            try
            {
                SearchResult result = deSearch.FindOne();
                de = new DirectoryEntry(result.Path);
                return de;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 根据用户帐号称取得用户的 对象
        /// </summary>
        /// <param name="sAMAccountName">用户帐号名</param>
        /// <returns>如果找到该用户，则返回用户的 对象；否则返回 null</returns>
        public static DirectoryEntry GetDirectoryEntryByAccount(string sAMAccountName)
        {
            DirectoryEntry de = GetDirectoryObject(ADUser, ADPassword);
            DirectorySearcher deSearch = new DirectorySearcher(de);
            deSearch.Filter = "(&(&(objectCategory=person)(objectClass=user))(sAMAccountName=" + sAMAccountName + "))";
            deSearch.SearchScope = SearchScope.Subtree;
            try
            {
                SearchResult result = deSearch.FindOne();
                de = new DirectoryEntry(result.Path, ADUser, ADPassword);
                return de;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 根据用户帐号和密码取得用户的 对象
        /// </summary>
        /// <param name="sAMAccountName">用户帐号名</param>
        /// <param name="password">用户密码</param>
        /// <returns>如果找到该用户，则返回用户的 对象；否则返回 null</returns>
        public static DirectoryEntry GetDirectoryEntryByAccount(string sAMAccountName, string password)
        {
            DirectoryEntry de = GetDirectoryEntryByAccount(sAMAccountName);
            if (de != null)
            {
                string commonName = de.Properties["cn"][0].ToString();
                if (GetDirectoryEntry(commonName, password) != null)
                    return GetDirectoryEntry(commonName, password);
                else
                    return null;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 根据组名取得用户组的 对象
        /// </summary>
        /// <param name="groupName">组名</param>
        /// <returns></returns>
        public static DirectoryEntry GetDirectoryEntryOfGroup(string groupName)
        {
            DirectoryEntry de = GetDirectoryObject();
            DirectorySearcher deSearch = new DirectorySearcher(de);
            deSearch.Filter = "(&(objectClass=group)(cn=" + groupName + "))";
            deSearch.SearchScope = SearchScope.Subtree;
            try
            {
                SearchResult result = deSearch.FindOne();
                de = new DirectoryEntry(result.Path);
                de.Username = ADUser;
                de.Password = ADPassword;
                return de;
            }
            catch
            {
                return null;
            }
        }
        public static DirectoryEntry GetDirectoryEntryOfGroup(string path, string adUserName, string password, string groupName)
        {
            DirectoryEntry de = new DirectoryEntry(path, adUserName, password, AuthenticationTypes.Secure);
            DirectorySearcher deSearch = new DirectorySearcher(de);
            deSearch.Filter = "(&(objectClass=group)(cn=" + groupName + "))";
            deSearch.SearchScope = SearchScope.Subtree;
            try
            {
                SearchResult result = deSearch.FindOne();
                de = new DirectoryEntry(result.Path);
                de.Username = ADUser;
                de.Password = ADPassword;
                return de;
            }
            catch
            {
                return null;
            }
        }

        #endregion

        #region GetProperty

        /// <summary>
        /// 获得指定 指定属性名对应的值 
        /// </summary>
        /// <param name="de">DirectoryEntry对象，如为用户则为用户的对象，部门则为部门的对象</param>
        /// <param name="propertyName">属性名称</param>
        /// <returns>属性值</returns>
        public static string GetProperty(DirectoryEntry de, string propertyName)
        {
            if (de.Properties.Contains(propertyName))
            {
                return de.Properties[propertyName][0].ToString();
            }
            else
            {
                return string.Empty;
            }
        }

        /// 
        /// 获得指定搜索结果 中指定属性名对应的值
        /// 
        /// 
        /// 属性名称
        /// 属性值
        public static string GetProperty(SearchResult searchResult, string propertyName)
        {
            if (searchResult.Properties.Contains(propertyName))
            {
                return searchResult.Properties[propertyName][0].ToString();
            }
            else
            {
                return string.Empty;
            }
        }

        #endregion

        #region SetProperty

        /// <summary>
        /// 设置指定 的属性值
        /// </summary>
        /// <param name="de"></param>
        /// <param name="propertyName">属性名称</param>
        /// <param name="propertyValue">属性值</param>
        public static void SetProperty(DirectoryEntry de, string propertyName, string propertyValue)
        {
            //if (propertyValue != string.Empty || propertyValue != "" || propertyValue != null)//
            //{
            //    if (de.Properties.Contains(propertyName))
            //    {
            //        de.Properties[propertyName][0] = propertyValue;
            //    }
            //    else
            //    {
            //        de.Properties[propertyName].Add(propertyValue);
            //    }
            //}
            if (de.Properties.Contains(propertyName))
            {
                if (string.IsNullOrEmpty(propertyValue))
                {
                    object o = de.Properties[propertyName].Value;
                    de.Properties[propertyName].Remove(o);
                }
                else
                {
                    de.Properties[propertyName][0] = propertyValue;
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(propertyValue))
                {
                    de.Properties[propertyName].Add(propertyValue);
                }
            }
        }

        #endregion

        #region 用户操作

        /// <summary>
        /// 创建带邮箱的用户
        /// </summary>
        /// <param name="CommonName">通用名（displayName,系统中显示的中文名字）</param>
        /// <param name="Account">帐户名（如ycan）</param>
        /// <param name="organizeName">组织单元名（有色院/科技处/信息中心）</param>
        /// <param name="password">密码</param>
        /// <param name="CreateMail">是否创建邮箱</param>
        /// <returns>新建用户的Path 错误则为“”</returns>
        public static string CreateADAccount(
            string CommonName,//CName
            string Account,//EName
            string organizeName,//
            string password,
            bool CreateMail
            )
        {
            //如果Acc存在则返回错误
            if (IsAccExists(Account))
                return "";
            DirectoryEntry entry = null;
            DirectoryEntry user = null;
            string samAccountName = Account;
            try
            {
                entry = new DirectoryEntry(GetOrganizeNamePath(organizeName), ADUser,
                ADPassword, AuthenticationTypes.Secure);

                user = entry.Children.Add("CN=" + CommonName, "user");
                user.Properties["userPrincipalName"].Value = samAccountName + sPrincpleNameTail;
                user.Properties["sAMAccountName"].Add(samAccountName);
                user.Properties["displayName"].Add(CommonName);

                user.CommitChanges();
                user.Invoke("SetPassword", new object[] { password });
                //This enables the new user.

                user.Properties["userAccountControl"].Value = //0x200; //ADS_UF_NORMAL_ACCOUNT
                     ADHelper.ADS_USER_FLAG_ENUM.ADS_UF_NORMAL_ACCOUNT | ADHelper.ADS_USER_FLAG_ENUM.ADS_UF_DONT_EXPIRE_PASSWD;
                user.CommitChanges();
                if (CreateMail)
                {
                    user.Invoke("CreateMailbox", new object[] { homeMDB });
                    user.CommitChanges();
                }
            }
            catch (Exception e)
            {
                try
                {
                    user.DeleteTree();
                }
                catch
                {
                    //return string.Empty;
                }
                throw e;
            }
            return user.Path;
        }


        /// <summary>
        /// 启用指定公共名称的用户
        /// </summary>
        /// <param name="commonName">用户公共名称</param>
        public static void EnableUser(string commonName)
        {
            EnableUser(GetDirectoryEntry(commonName));
        }


        /// <summary>
        /// 启用指定帐户
        /// </summary>
        /// <param name="de"></param>
        public static void EnableUser(DirectoryEntry de)
        {
            de.Properties["userAccountControl"][0] = ADHelper.ADS_USER_FLAG_ENUM.ADS_UF_NORMAL_ACCOUNT | ADHelper.ADS_USER_FLAG_ENUM.ADS_UF_DONT_EXPIRE_PASSWD;
            de.CommitChanges();
            de.Close();
        }

        /// <summary>
        /// 禁用指定公共名称的用户
        /// </summary>
        /// <param name="commonName">用户公共名称</param>
        public static void DisableUser(string commonName)
        {
            DisableUser(GetDirectoryEntry(commonName));
        }

        /// <summary>
        /// 禁用指定的帐户
        /// </summary>
        /// <param name="de"></param>
        public static void DisableUser(DirectoryEntry de)
        {
            //impersonate.BeginImpersonate();
            de.Properties["userAccountControl"][0] =
                ADHelper.ADS_USER_FLAG_ENUM.ADS_UF_NORMAL_ACCOUNT | ADHelper.ADS_USER_FLAG_ENUM.ADS_UF_DONT_EXPIRE_PASSWD | ADHelper.ADS_USER_FLAG_ENUM.ADS_UF_ACCOUNTDISABLE;
            de.CommitChanges();
            //impersonate.StopImpersonate();
            de.Close();
        }


        /// <summary> 
        /// 修改用户密码
        /// </summary>
        /// <param name="commonName">用户公共名称</param>
        /// <param name="oldPassword">旧密码</param>
        /// <param name="newPassword">新密码</param>
        public static void ChangeUserPassword(string commonName,
            string oldPassword, string newPassword)
        {
            // to-do: 需要解决密码策略问题
            DirectoryEntry oUser = GetDirectoryEntry(commonName);
            oUser.Invoke("ChangePassword", new Object[] { oldPassword, newPassword });
            oUser.Close();
        }

        /// <summary> 
        /// 修改Acc密码
        /// </summary>
        /// <param name="sAMacc">帐户名称</param>
        /// <param name="oldPassword">旧密码</param>
        /// <param name="newPassword">新密码</param>
        /// <returns>是否成功</returns>
        public static bool ChangeAccPassword(string sAMacc,
            string oldPassword, string newPassword)
        {
            try
            {
                if (IsAccExists(sAMacc))
                {
                    // to-do: 需要解决密码策略问题
                    DirectoryEntry oUser = GetDirectoryEntryByAccount(sAMacc);
                    //oUser.Invoke("ChangePassword", new Object[] { oldPassword, newPassword });
                    oUser.Invoke("SetPassword", new object[] { newPassword });
                    oUser.CommitChanges();
                    oUser.Close();
                    return true;
                }
                else
                    return false;
            }
            catch (Exception)
            {
                return false;
            }

        }

        /// <summary>
        /// 移动用户
        /// </summary>
        /// <param name="user_path">用户Path</param>
        /// <param name="target_path">目标path</param>
        /// <returns></returns>
        public static string MoveUser(string user_path, string target_path)
        {
            DirectoryEntry u = new DirectoryEntry(user_path, ADUser, ADPassword);
            DirectoryEntry t = new DirectoryEntry(target_path, ADUser, ADPassword);
            u.MoveTo(t);
            return u.Path;
        }

        /// <summary>
        /// 重名用户
        /// </summary>
        /// <param name="sAcc"></param>
        /// <param name="newDisplayName"></param>
        /// <returns></returns>
        public static bool RenameUser(string sAcc, string newDisplayName)
        {
            try
            {
                if (IsAccExists(sAcc))
                {
                    DirectoryEntry userEntry = GetDirectoryEntryByAccount(sAcc);
                    //userEntry.Properties["cn"][0] = newDisplayName;
                    userEntry.Properties["displayName"][0] = newDisplayName;

                    userEntry.Rename("CN=" + newDisplayName);
                    userEntry.CommitChanges();
                    userEntry.Dispose();
                    return true;
                }
                else
                    return false;
            }
            catch (Exception)
            {

                return false;
            }
        }

        /// <summary>
        /// 重命名Account
        /// </summary>
        /// <param name="oldAcc">原Acc</param>
        /// <param name="newAcc">新Acc</param>
        /// <returns>成功True，如果已经存在新、老Acc则返回false，错误也返回false</returns>
        public static bool RenameAcc(string oldAcc, string newAcc)
        {
            try
            {
                if (IsAccExists(oldAcc))
                {
                    //不允许同名ACC
                    if (IsAccExists(newAcc))
                        return false;
                    DirectoryEntry userEntry = GetDirectoryEntryByAccount(oldAcc);

                    userEntry.Properties["sAMAccountName"][0] = newAcc;
                    userEntry.CommitChanges();
                    userEntry.Dispose();
                    return true;
                }
                else
                    return false;
            }
            catch (Exception)
            {

                return false;
            }

        }

        public static void SetUserPassword(string userName, string password)
        {
            SetUserPassword(null, null, userName, password);
        }

        public static void SetUserPassword(string adminName, string adminPassword, string userName, string password)
        {
            adminName = ADUser;
            adminPassword = ADPassword;
            DirectoryEntry userEntry = FindObject(adminName, adminPassword, "user", userName);
            userEntry.Invoke("SetPassword", new object[] { password });
            userEntry.CommitChanges();
        }

        /// <summary>
        /// 删除AD账户，使用当前上下文的安全信息，一般用于Windows程序
        /// </summary>
        /// <param name="userName">用户名称</param>
        /// <returns></returns>
        public static bool DeleteADAccount(string Account)
        {
            //DeleteADAccount(null, null, userName);
            if (IsAccExists(Account))
            {
                string AccPath = GetAccPath(Account);
                DetTree(AccPath);
                return true;
            }
            else
                return false;
        }


        /// <summary>
        /// 删除AD账户，使用指定的用户名和密码来模拟，一般用于ASP.NET程序
        /// </summary>
        /// <param name="adminUser">AD管理员用户名，如默认，可以为null</param>
        /// <param name="adminPassword">AD管理员密码，如默认，可以为null</param>
        /// <param name="userName">用户名称</param>
        public static bool DeleteADAccount(string adminUser, string adminPassword, string userName)
        {
            DirectoryEntry user = null;
            DirectoryEntry Container = null;
            try
            {
                adminUser ="Administrator";//ADUser;
                adminPassword = "123.com";//ADPassword;
                user = FindObject(adminUser, adminPassword, "user", userName);
                Container = user.Parent;
                Container.Children.Remove(user);
                Container.CommitChanges();
                Container.Dispose();
                user.Dispose();
                return true;

            }
            catch (Exception)
            {
                Container.Dispose();
                user.Dispose();
                return false;
            }

        }

        public static DirectoryEntry FindObject(string category, string name)
        {
            return FindObject(null, null, category, name);
        }

        public static DirectoryEntry FindObject(string adminName,
            string adminPassword, string category, string name)
        {
            adminName = "Administrator"; //ADUser;
            adminPassword = "123.com"; //ADPassword;
            string ADPath = "LDAP://172.16.253.234";
            DirectoryEntry de = null;
            //if (adminName == null || adminPassword == null)
            //{
            de = new DirectoryEntry(ADPath, adminName, adminPassword, AuthenticationTypes.Secure);
            //}
            //else
            //{
            //    de = new DirectoryEntry();
            //}
            ///organizationalUnit
            DirectorySearcher ds = new DirectorySearcher(de);
            string queryFilter = string.Format("(&(objectCategory=" +
                category + ")(sAMAccountName={0}))", name);
            if (category.ToUpper() == "OU")
                queryFilter = string.Format("(&(objectCategory=OrganizationalUnit)(OU={0}))", name);
            ds.Filter = queryFilter;
            ds.Sort.PropertyName = "cn";

            DirectoryEntry userEntry = null;
            try
            {
                SearchResult sr = ds.FindOne();
                userEntry = sr.GetDirectoryEntry();
            }
            finally
            {
                if (de != null)
                {
                    de.Dispose();
                }
                if (ds != null)
                {
                    ds.Dispose();
                }
            }
            return userEntry;
        }

        public static DirectoryEntry FindObject2(string category, string name,string adPath)
        {
            DirectoryEntry de = null;
            //if (adminName == null || adminPassword == null)
            //{
            de = new DirectoryEntry(adPath, ADUser, ADPassword, AuthenticationTypes.Secure);
            //}
            //else
            //{
            //    de = new DirectoryEntry();
            //}
            ///organizationalUnit
            DirectorySearcher ds = new DirectorySearcher(de);
            string queryFilter = string.Format("(&(objectCategory=" +
                category + ")(sAMAccountName={0}))", name);
            if (category.ToUpper() == "OU")
                queryFilter = string.Format("(&(objectCategory=OrganizationalUnit)(OU={0}))", name);
            ds.Filter = queryFilter;
            ds.Sort.PropertyName = "cn";

            DirectoryEntry userEntry = null;
            try
            {
                SearchResult sr = ds.FindOne();
                userEntry = sr.GetDirectoryEntry();
            }
            finally
            {
                if (de != null)
                {
                    de.Dispose();
                }
                if (ds != null)
                {
                    ds.Dispose();
                }
            }
            return userEntry;
        }

        #endregion

        #region 组织单元操作

        /// <summary>
        /// 重命名OU
        /// </summary>>
        /// <param name="oldOUName">原名</param>
        /// <param name="newOUName">新名</param>
        /// <returns>是否成功</returns>
        public static bool RenameOU(string oldOUName, string newOUName)
        {

            try
            {
                if (CheckOU(oldOUName))
                {
                    DirectoryEntry userEntry = new DirectoryEntry(GetOrganizeNamePath(oldOUName), ADUser, ADPassword);
                    userEntry.Rename("OU=" + newOUName);
                    userEntry.CommitChanges();
                    userEntry.Dispose();
                    if (CheckOU(newOUName))
                        return true;
                    else
                        return false;
                }
                else
                    return false;
            }
            catch (Exception)
            {
                return false;
                //throw;
            }
        }

        /// <summary>
        /// Get Users From OU
        /// </summary>>     
        /// <returns></returns>
        public static DataTable GetUsersFromOU(string OUName)
        {
            try
            {
                if (CheckOU(OUName))
                {
                    DirectoryEntry OUEntry = new DirectoryEntry(GetOrganizeNamePath(OUName), ADUser, ADPassword);
                    //表名
                    DataTable tbUsers = new DataTable("users");
                    //用户列
                    tbUsers.Columns.Add("UserID", System.Type.GetType("System.String"));
                    tbUsers.Columns.Add("UserName", System.Type.GetType("System.String"));
                    tbUsers.Columns.Add("UserAccount", System.Type.GetType("System.String"));
                    DirectorySearcher src = new DirectorySearcher("(&(objectCategory=Person)(objectClass=user))");
                    DirectoryEntry de = new DirectoryEntry(ADPath + "/OU=" + OUName + "," + LDAPDomain, ADUser, ADPassword);
                    src.SearchRoot = de;
                    src.SearchScope = SearchScope.Subtree;
                    string strUserAccount = null;

                    foreach (SearchResult res in src.FindAll())
                    {
                        System.Collections.IDictionaryEnumerator ien = res.Properties.GetEnumerator();
                        DataRow topRow = tbUsers.NewRow();
                        topRow["UserID"] = GetProperty(res, "Name");
                        topRow["UserName"] = GetProperty(res, "displayname");
                        //得到用户状态
                        int userAccountControl = Convert.ToInt32(GetProperty(res, "userAccountControl"));
                        //判断用户是否启用
                        if (IsAccountActive(userAccountControl))
                        {
                            topRow["UserAccount"] = "激活";
                        }
                        else
                        {
                            topRow["UserAccount"] = "禁用";
                        }
                        tbUsers.Rows.Add(topRow);
                    }
                    return tbUsers;
                }
                else
                    return null;
            }
            catch (Exception)
            {
                return null;
                //throw;
            }
        }

        /// <summary>
        /// Get Users From File
        /// </summary>>     
        /// <returns></returns>
        public static DataTable GetUsersFromFile(string FileName)
        {
            try
            {
                //表名
                DataTable tbUsers = new DataTable("users");
                //用户列
                tbUsers.Columns.Add("UserID", System.Type.GetType("System.String"));
                tbUsers.Columns.Add("UserName", System.Type.GetType("System.String"));
                tbUsers.Columns.Add("UserAccount", System.Type.GetType("System.String"));
                DirectorySearcher src = new DirectorySearcher("(&(objectCategory=Person)(objectClass=user))");
                DirectoryEntry de = new DirectoryEntry(ADPath + "/CN=" + FileName + "," + LDAPDomain);
                src.SearchRoot = de;
                src.SearchScope = SearchScope.Subtree;
                string strUserAccount = null;

                foreach (SearchResult res in src.FindAll())
                {
                    System.Collections.IDictionaryEnumerator ien = res.Properties.GetEnumerator();
                    DataRow topRow = tbUsers.NewRow();
                    topRow["UserID"] = GetProperty(res, "Name");
                    topRow["UserName"] = GetProperty(res, "displayname");
                    //得到用户状态
                    int userAccountControl = Convert.ToInt32(GetProperty(res, "userAccountControl"));
                    //判断用户是否启用
                    if (IsAccountActive(userAccountControl))
                    {
                        topRow["UserAccount"] = "激活";
                    }
                    else
                    {
                        topRow["UserAccount"] = "禁用";
                    }
                    tbUsers.Rows.Add(topRow);
                }
                return tbUsers;
            }
            catch (Exception)
            {
                return null;
                //throw;
            }
        }

        /// <summary>
        /// 检查组织单位（OU）是否存在
        /// </summary>
        /// <param name="sOU">组织单位名称，示例：PacteraUsers 或 PacteraUsers/Others</param>
        /// <returns>成功返回True,否则返回False</returns>
        public static bool CheckOU(string sOU)
        {
            DirectoryEntry objOU = null;
            try
            {
                string path = GetOrganizeNamePath(sOU);
                objOU = new DirectoryEntry(path, ADUser, ADPassword);
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

        /// <summary>
        /// 创建OU，需要指定连接到AD的授权信息
        /// </summary>
        /// <param name="adminName"></param>
        /// <param name="adminPassword"></param>
        /// <param name="name"></param>
        /// <param name="parentOrganizeUnit"></param>
        public static DirectoryEntry CreateOrganizeUnit(string adminName,
            string adminPassword, string name, string parentOrganizeUnit)
        {
            adminName = ADUser;
            adminPassword = ADPassword;
            DirectoryEntry parentEntry = null;
            if (adminName == null || adminPassword == null)
            {
                parentEntry = new DirectoryEntry(GetOrganizeNamePath(parentOrganizeUnit));
            }
            else
            {
                if (parentOrganizeUnit != "")
                {
                    parentEntry = new DirectoryEntry(GetOrganizeNamePath(parentOrganizeUnit),
                       adminName, adminPassword,
                       AuthenticationTypes.Secure);
                }
                else
                {
                    string path = GetOrganizeNamePath(parentOrganizeUnit);
                    parentEntry = new DirectoryEntry(path, adminName, adminPassword, AuthenticationTypes.Secure);
                }
            }
            DirectoryEntry organizeEntry = parentEntry.Children.Add("OU=" +
                name, "organizationalUnit");
            organizeEntry.CommitChanges();
            return organizeEntry;
        }

        /// <summary>
        /// 创建OU，不需要指定连接到AD的授权信息，用于Windows程序
        /// </summary>
        /// <param name="name">OU name</param>
        /// <param name="parentOrganizeUnit">父OU路径如a/b/c</param>
        public static DirectoryEntry CreateOrganizeUnit(string name, string parentOrganizeUnit)
        {
            return CreateOrganizeUnit(null, null, name, parentOrganizeUnit);
        }

        /// <summary>
        /// 将用户加入到用户组中
        /// </summary>
        /// <param name="userName">用户名</param>
        /// <param name="organizeName">组织名</param>
        /// <param name="groupName">组名</param>
        /// <exception cref="InvalidObjectException">用户名或用户组不存在</exception>
        public static void AddUserToGroup(string userName, string groupName)
        {
            AddUserToGroup(null, null, userName, groupName);
        }

        /// <summary>
        /// 将用户加入到用户组中
        /// </summary>
        /// <param name="adminName"></param>
        /// <param name="adminPassword"></param>
        /// <param name="userName">用户名</param>
        /// <param name="groupName">组名</param>
        /// <exception cref="InvalidObjectException">用户名或用户组不存在</exception>
        public static void AddUserToGroup(string adminName, string adminPassword, string userName, string groupName)
        {
            adminName = ADUser;
            adminPassword = ADPassword;
            DirectoryEntry rootUser = null;
            if (adminName == null || adminPassword == null)
            {
                rootUser = new DirectoryEntry(GetUserPath(), adminName, adminPassword, AuthenticationTypes.Secure);
            }
            else
            {
                rootUser = new DirectoryEntry(GetUserPath());
            }

            DirectoryEntry group = null;
            DirectoryEntry user = null;

            try
            {
                group = rootUser.Children.Find("CN=" + groupName);
            }
            catch (Exception)
            {
                throw new Exception("在域中不存在组“" + groupName + "”");
            }

            try
            {
                user = FindObject(adminName, adminPassword, "user", userName);
            }
            catch (Exception)
            {
                throw new Exception("在域中不存在用户“" + userName + "”");
            }

            //加入用户到用户组中
            group.Properties["member"].Add(user.Properties["distinguishedName"].Value);
            group.CommitChanges();
        }


        /// <summary>
        /// 彻底删除OU/也可以通过参数控制删除空OU,即安全删除
        /// </summary>
        /// <param name="OUName">OU名,如a/b/c,必须写绝对OU名</param>
        /// <param name="boolComplete">true为彻底删除,false为安全删除</param>
        /// <returns>成功与否</returns>
        public static bool DeleteOU(string OUName, bool boolComplete)
        {
            try
            {
                if (!boolComplete)
                    return false;
                if (CheckOU(OUName))
                {
                    string OUpath = GetOrganizeNamePath(OUName);
                    if (OUpath == "")
                        return false;
                    else
                    {
                        DetTree(OUpath);
                        return true;
                    }
                }
                else
                    return true;
            }
            catch (Exception)
            {
                throw;
                //return false;
            }
        }

        /// <summary>
        /// 删除OU（仅仅支持纯部门删除，OU中不能有其他用户、OU等信息的子节点）
        /// </summary>
        /// <param name="OUName">OU名,如a/b/c,必须写绝对OU名</param>
        /// <returns>"成功"  "不存在"  "非空"</returns>
        public static string DeleteOU(string OUName)
        {
            try
            {
                if (CheckOU(OUName))
                {
                    DirectoryEntry objOU = new DirectoryEntry(GetOrganizeNamePath(OUName));
                    DirectoryEntry parOU = objOU.Parent;
                    parOU.Children.Remove(objOU);
                    parOU.CommitChanges();
                    return "成功";
                }
                else
                    return "不存在";
            }
            catch (Exception)
            {
                return "非空";
            }
        }

        #endregion

        #region 登录相关

        /// <summary> 
        /// 判断用户与密码是否足够以满足身份验证进而登录
        /// </summary>
        /// <param name="commonName">用户公共名称</param>
        /// <param name="password">密码</param>
        /// <returns>如能可正常登录，则返回 true；否则返回 false</returns>
        public static LoginResult Login(string commonName, string password)
        {
            DirectoryEntry de = GetDirectoryEntry(commonName);

            if (de != null)
            {
                // 必须在判断用户密码正确前，对帐号激活属性进行判断；否则将出现异常。
                int userAccountControl =
                    Convert.ToInt32(de.Properties["userAccountControl"][0]);
                de.Close();
                if (!IsAccountActive(userAccountControl))

                    return LoginResult.LOGIN_USER_ACCOUNT_INACTIVE;
                if (GetDirectoryEntry(commonName, password) != null)
                    return LoginResult.LOGIN_USER_OK;
                else
                    return LoginResult.LOGIN_USER_PASSWORD_INCORRECT;
            }
            else
            {
                return LoginResult.LOGIN_USER_DOESNT_EXIST;
            }
        }


        /// <summary>
        /// 判断用户帐号与密码是否足够以满足身份验证进而登录
        /// </summary>
        /// <param name="sAMAccountName">用户帐号</param>
        /// <param name="password">密码</param>
        /// <returns>如能可正常登录，则返回 true；否则返回 false</returns>
        public static LoginResult LoginByAccount(string sAMAccountName, string password)
        {
            DirectoryEntry de = GetDirectoryEntryByAccount(sAMAccountName);
            if (de != null)
            {
                // 必须在判断用户密码正确前，对帐号激活属性进行判断；否则将出现异常。
                int userAccountControl =
                    Convert.ToInt32(de.Properties["userAccountControl"][0]);
                de.Close();
                if (!IsAccountActive(userAccountControl))
                    return LoginResult.LOGIN_USER_ACCOUNT_INACTIVE;
                if (GetDirectoryEntryByAccount(sAMAccountName, password) != null)
                    return LoginResult.LOGIN_USER_OK;
                else
                    return LoginResult.LOGIN_USER_PASSWORD_INCORRECT;
            }
            else
            {
                return LoginResult.LOGIN_USER_DOESNT_EXIST;
            }
        }

        /// <summary>
        /// 判断用户帐号是否激活
        /// </summary>
        /// <param name="userAccountControl">用户帐号属性控制器</param>
        /// <returns>如果用户帐号已经激活，返回 true；否则返回 false</returns>
        public static bool IsAccountActive(int userAccountControl)
        {
            int userAccountControl_Disabled =
                Convert.ToInt32(ADS_USER_FLAG_ENUM.ADS_UF_ACCOUNTDISABLE);
            int flagExists = userAccountControl & userAccountControl_Disabled;
            if (flagExists > 0)
                return false;
            else
                return true;
        }

        /// <summary>
        ///  判断用户是否存在
        /// </summary>
        /// <param name="commonName">公用名，例如：San Zhang</param>
        /// <returns></returns>
        public static bool IsUserExists(string commonName)
        {
            DirectoryEntry de = GetDirectoryObject();
            DirectorySearcher deSearch = new DirectorySearcher(de);
            deSearch.Filter = "(&(&(objectCategory=person)(objectClass=user))(cn=" + commonName + "))";       // LDAP 查询串
            SearchResultCollection results = deSearch.FindAll();
            if (results.Count == 0)
                return false;
            else
                return true;
        }

        /// <summary>
        ///  判断帐户是否存在（Windows2000以前版本）
        /// </summary>
        /// <param name="commonName">Account用户名</param>
        /// <returns>是否存在</returns>
        public static bool IsAccExists(string sAMAccountName)
        {
            DirectoryEntry de = GetDirectoryObject();
            DirectorySearcher deSearch = new DirectorySearcher(de);
            deSearch.Filter = "(&(&(objectCategory=person)(objectClass=user))(sAMAccountName=" +
                sAMAccountName + "))";       // LDAP 查询串
            SearchResultCollection results = deSearch.FindAll();
            if (results.Count == 0)
                return false;
            else
                return true;
        }
        /// <summary>
        /// 判断帐户是否存在（还没有测试过）
        /// </summary>
        /// <param name="userPrincipalName">Account主体名，例如：San.Zhang</param>
        /// <returns>是否存在</returns>
        public static bool IsAccExists2(string userPrincipalName)
        {
            DirectoryEntry de = GetDirectoryObject();
            DirectorySearcher deSearch = new DirectorySearcher(de);
            deSearch.Filter = "(&(&(objectCategory=person)(objectClass=user))(userPrincipalName=" + userPrincipalName + "))";       // LDAP 查询串
            SearchResultCollection results = deSearch.FindAll();
            if (results.Count == 0)
            {
                deSearch.Filter = "(&(&(objectCategory=person)(objectClass=user))(userPrincipalName=" + userPrincipalName + ")(userAccountControl=514))";
                results = deSearch.FindAll();
                if (results.Count == 0)
                    return false;
                else
                    return true;
            }
            else
                return true;
        }

        #endregion

        #region 与AD的DN解析有关工具函数

        /// <summary>
        /// 删除对象
        /// </summary>
        /// <param name="path">对象路径</param>
        public static void DetTree(string path)
        {
            try
            {
                DirectoryEntry ent = new DirectoryEntry(path, ADUser, ADPassword);
                ent.DeleteTree();

            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 获取所有用户所在的安全组
        /// </summary>
        /// <returns></returns>
        public static string GetUserPath()
        {
            return GetUserPath(null);
        }

        /// <summary>
        /// 获取所有没有在AD组织中的用户DN名称
        /// </summary>
        /// <param name="userName"></param>
        /// <returns></returns>
        public static string GetUserPath(string userName)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(ADPath);
            if (userName != null && userName.Length > 0)
            {
                sb.Append("CN=").Append(userName).Append(",");
            }
            sb.Append("CN=Users,").Append(GetDomainDN());
            return sb.ToString();
        }

        /// <summary>
        /// 根据用户所在的组织结构来构造用户在AD中的DN路径
        /// </summary>
        /// <param name="userName">用户名称</param>
        /// <param name="organzieName">组织结构</param>
        /// <returns></returns>
        public static string GetUserPath(string userName, string organzieName)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(ADPath);
            sb.Append("CN=").Append(userName).Append(",").Append(SplitOrganizeNameToDN(organzieName));
            return sb.ToString();
        }

        /// <summary>
        /// 获得帐户Acc的Path
        /// </summary>
        /// <param name="sAcc">Acc</param>
        /// <returns>错误就是empty</returns>
        public static string GetAccPath(string sAcc)
        {
            DirectoryEntry de = GetDirectoryObject();
            DirectorySearcher deSearch = new DirectorySearcher(de);
            deSearch.Filter = "(&(&(objectCategory=person)(objectClass=user))(sAMAccountName=" + sAcc + "))";
            deSearch.SearchScope = SearchScope.Subtree;
            try
            {
                SearchResult result = deSearch.FindOne();
                string Apath = result.Path;
                result = null;
                deSearch.Dispose();
                de.Dispose();
                return Apath;
            }
            catch (Exception)
            {
                deSearch.Dispose();
                de.Dispose();
                return string.Empty;
            }
        }

        /// <summary>
        /// 获取域的后缀DN名,如域为cinf.com,则返回"DC=cinf,DC=Com"
        /// </summary>
        /// <returns></returns>
        public static string GetDomainDN()
        {
            return LDAPDomain;
        }

        /// <summary>
        /// 获得OU的Path
        /// </summary>
        /// <param name="organizeUnit">OU名</param>
        /// <returns></returns>
        public static string GetOrganizeNamePath(string organizeUnit)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(ADPath);//LDAP_IDENTITY
            sb.Append("/");
            string path = sb.Append(SplitOrganizeNameToDN(organizeUnit)).ToString();
            return path;
        }
        public static string GetOrganizeNamePath(string organizeUnit, string adPath)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(adPath);//示例：LDAP://192.168.88.64
            sb.Append("/");
            string path = sb.Append(SplitOrganizeNameToDN(organizeUnit)).ToString();
            return path;
        }

        /// <summary>
        /// 分离组织名称为标准AD的DN名称,各个组织级别以"/"或"\"分开。如"总部/物业公司/小区"，并且当前域为
        /// ExchangeTest.Com，则返回的AD的DN表示名为"OU=小区,OU=物业公司,OU=总 
        /// 部,DC=ExchangeTest,DC=Com"。 
        /// </summary>
        /// <param name="organizeName">组织名称</param>
        /// <returns>返回一个级别</returns>
        public static string SplitOrganizeNameToDN(string organizeName)
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
            sb.Append(GetDomainDN());
            return sb.ToString();
        }

        #endregion

        #region 扩展

        public static void AddUserToGroup1(DirectoryEntry oGroup, DirectoryEntry oUser)
        {
            if (oGroup != null && !oGroup.Properties["member"].Contains(oUser.Properties["distinguishedName"].Value))
            {
                oGroup.Properties["member"].Add(oUser.Properties["distinguishedName"].Value);
                oGroup.CommitChanges();
            }
            if (oGroup != null)
                oGroup.Close();
            if (oUser != null)
                oUser.Close();
        }

        /// <summary>
        ///将指定的用户添加到指定的组中。默认为Users下的组和用户。
        ///用户公共名称
        ///组名
        /// </summary>
        /// <param name="userCommonName"></param>
        /// <param name="groupName"></param>
        public static void AddUserToGroup2(string userCommonName, string groupName)
        {
            DirectoryEntry oGroup = GetDirectoryEntryOfGroup(groupName);
            DirectoryEntry oUser = GetDirectoryEntry(userCommonName);

            if (oGroup != null && !oGroup.Properties["member"].Contains(oUser.Properties["distinguishedName"].Value))
            {
                oGroup.Properties["member"].Add(oUser.Properties["distinguishedName"].Value);
                oGroup.CommitChanges();
            }
            if (oGroup != null)
                oGroup.Close();
            if (oUser != null)
                oUser.Close();
        }

        /// <summary>
        /// 从组中移除用户
        /// </summary>
        /// <param name="userCommonName"></param>
        /// <param name="groupName"></param>
        public static void RemoveUserFromGroup(string userCommonName, string groupName)
        {
            DirectoryEntry oGroup = GetDirectoryEntryOfGroup(groupName);
            DirectoryEntry oUser = GetDirectoryEntry(userCommonName);

            if (!oGroup.Properties["member"].Contains(oUser.Properties["distinguishedName"].Value))
            {
                oGroup.Properties["member"].Remove(oUser.Properties["distinguishedName"].Value);
                oGroup.CommitChanges();
            }

            oGroup.Close();
            oUser.Close();
        }

        /// <summary>
        /// 移除组中所有用户
        /// </summary>
        /// <param name="groupName"></param>
        public static void RemoveAllUserFromGroup(string groupName)
        {
            DirectoryEntry oGroup = GetDirectoryEntryOfGroup(groupName);
            oGroup.Properties["member"].Clear();
            oGroup.CommitChanges();
            oGroup.Close();
        }

        /// <summary>
        ///创建新的组
        ///DN 位置。例如：OU=共享平台 或 CN=Users
        ///公共名称
        ///帐号
        /// </summary>
        /// <param name="ldapDN"></param>
        /// <param name="commonName"></param>
        /// <param name="sAMAccountName"></param>
        /// <returns></returns>
        public static DirectoryEntry CreateNewGroup(string ldapDN, string commonName, string sAMAccountName)
        {
            DirectoryEntry entry = GetDirectoryObject();
            DirectoryEntry subEntry = entry.Children.Find(ldapDN);
            DirectoryEntry deGroup = subEntry.Children.Add("CN=" + commonName, "group");
            deGroup.Properties["sAMAccountName"].Value = sAMAccountName;
            deGroup.Properties["groupType"].Value = 8;
            deGroup.CommitChanges();
            deGroup.Close();
            return deGroup;
        }

        /// <summary>
        /// 创建新的用户
        /// </summary>
        /// <param name="ldapDN">PacteraUsers 或 PacteraUsers/Others（旧的：OU=PacteraUsers 或 CN=Users）</param>
        /// <param name="commonName"></param>
        /// <param name="sAMAccountName"></param>
        /// <param name="password"></param>
        /// <param name="adPath">示例：LDAP://192.168.88.64</param>
        /// <param name="adUser"></param>
        /// <param name="adUserPwd"></param>
        /// <returns></returns>
        public static DirectoryEntry CreateNewUser(string ldapDN, string commonName, string sAMAccountName, string password, string adPath, string adUser, string adUserPwd)
        {
            DirectoryEntry deUser = null;
            string path = GetOrganizeNamePath(ldapDN, adPath);
            DirectoryEntry entry = null;
            try
            {
                entry = new DirectoryEntry(path, ADUser, ADPassword, AuthenticationTypes.Secure);
            }
            catch (Exception ex)
            {
                throw new Exception("连接域控失败：" + ex.Message);
            }

            if (entry != null)
            {
                deUser = entry.Children.Add("CN=" + commonName.Replace(",", "\\,"), "user");
                deUser.Properties["sAMAccountName"].Value = sAMAccountName;
            }
            try
            {
                deUser.CommitChanges();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            DirectoryEntry de = GetDirectoryEntry(commonName, adPath, adUser, adUserPwd);
            //DirectoryEntry de = GetDirectoryEntry(commonName);
            if (de != null)
            {
                de.Invoke("SetPassword", new object[] { password });
                de.CommitChanges();
                de.Properties["userAccountControl"][0] = ADHelper.ADS_USER_FLAG_ENUM.ADS_UF_NORMAL_ACCOUNT | ADHelper.ADS_USER_FLAG_ENUM.ADS_UF_DONT_EXPIRE_PASSWD;
                de.CommitChanges();
            }
            else
            {
                throw new Exception("设置密码和启用账号失败");
            }
            deUser.Close();
            return deUser;
        }
        /// <summary>
        ///创建新的用户
        /// </summary>
        /// <param name="ldapDN">PacteraUsers 或 PacteraUsers/Others（旧的：OU=PacteraUsers 或 CN=Users）</param>
        /// <param name="commonName">公共名称</param>
        /// <param name="sAMAccountName">sAMAccountName</param>
        /// <param name="password">sAMAccountName的密码</param>
        /// <returns></returns>
        public static DirectoryEntry CreateNewUser(string ldapDN, string commonName, string sAMAccountName, string password)
        {
            DirectoryEntry deUser = null;
            string path = GetOrganizeNamePath(ldapDN);
            DirectoryEntry entry = null;
            try
            {
                entry = new DirectoryEntry(path, ADUser, ADPassword, AuthenticationTypes.Secure);
            }
            catch (Exception ex)
            {
                throw new Exception("连接域控失败：" + ex.Message);
            }

            if (entry != null)
            {
                deUser = entry.Children.Add("CN=" + commonName.Replace(",", "\\,"), "user");
                deUser.Properties["sAMAccountName"].Value = sAMAccountName;
            }
            try
            {
                deUser.CommitChanges();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            DirectoryEntry de = null;
            string msg = string.Empty;
            try
            {
                de = GetDirectoryEntry(commonName);
            }
            catch (Exception ex)
            {
                msg = ex.Message;
            }
            if (de != null)
            {
                de.Invoke("SetPassword", new object[] { password });
                de.CommitChanges();
                de.Properties["userAccountControl"][0] = ADHelper.ADS_USER_FLAG_ENUM.ADS_UF_NORMAL_ACCOUNT | ADHelper.ADS_USER_FLAG_ENUM.ADS_UF_DONT_EXPIRE_PASSWD;
                de.CommitChanges();
            }
            else
            {
                throw new Exception("设置密码和启用账号失败：" + msg);
            }
            deUser.Close();

            //DirectoryEntry entry = GetDirectoryObject();
            //DirectoryEntry subEntry = entry.Children.Find(ldapDN);
            //deUser = subEntry.Children.Add("CN=" + commonName, "user");
            //deUser.Properties["sAMAccountName"].Value = sAMAccountName;
            //deUser.CommitChanges();
            //ADHelper.SetPassword(commonName, password);
            //ADHelper.EnableUser(commonName);
            //deUser.Close();
            return deUser;
        }

        /// <summary>
        ///创建新的用户。默认创建在Users单元下。
        ///公共名称
        ///帐号
        ///密码
        /// </summary>
        /// <param name="commonName"></param>
        /// <param name="sAMAccountName"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public static DirectoryEntry CreateNewUser(string commonName, string sAMAccountName, string password)
        {
            return CreateNewUser("CN=Users", commonName, sAMAccountName, password);
        }

        /// <summary>
        ///设置用户密码，管理员可以通过它来修改指定用户的密码。
        ///用户公共名称
        ///用户新密码
        /// </summary>
        /// <param name="commonName"></param>
        /// <param name="newPassword"></param>
        public static void SetPassword(string commonName, string newPassword)
        {
            DirectoryEntry de = GetDirectoryEntry(commonName);
            // 模拟超级管理员，以达到有权限修改用户密码;
            de.Invoke("SetPassword", new object[] { newPassword });
            de.Close();
        }

        /// <summary>
        ///设置帐号密码，管理员可以通过它来修改指定帐号的密码。
        ///用户帐号
        ///用户新密码
        /// </summary>
        /// <param name="sAMAccountName"></param>
        /// <param name="newPassword"></param>
        public static void SetPasswordByAccount(string sAMAccountName, string newPassword)
        {
            DirectoryEntry de = GetDirectoryEntryByAccount(sAMAccountName);
            // 模拟超级管理员，以达到有权限修改用户密码
            de.Invoke("SetPassword", new object[] { newPassword });
            de.Close();
        }

        /// <summary>
        /// 判断是否存在
        /// </summary>
        /// <param name="objectName"></param>
        /// <param name="catalog"></param>
        /// <returns></returns>
        public static bool ObjectExists(string objectName, string catalog)
        {
            DirectoryEntry de = GetDirectoryObject();
            DirectorySearcher deSearch = new DirectorySearcher();
            deSearch.SearchRoot = de;
            switch (catalog)
            {
                case "User": deSearch.Filter = "(&(objectClass=user) (cn=" + objectName + "))"; break;
                case "Group": deSearch.Filter = "(&(objectClass=group) (cn=" + objectName + "))"; break;
                case "OU": deSearch.Filter = "(&(objectClass=OrganizationalUnit) (OU=" + objectName + "))"; break;
                default: break;
            }
            SearchResultCollection results = deSearch.FindAll();
            if (results.Count == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        #endregion
    }
}