using ADModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices;
using System.Linq;
using System.Text;

namespace Bussiness
{
    public class ADAccountManage
    {
        public ExecResult CreateAccount(ADAccount adAccount)
        {
            //所有的AD用户都创建在“PacteraUsers”OU中
            string ldapDN = GetPacteraUsersOU();
            //创建账号
            DirectoryEntry de = ADHelper.CreateNewUser(ldapDN, adAccount.CommonName, adAccount.SamAccountName, adAccount.Password);
            return ExecResult.Success;
        }

        public ExecResult UpdateAccount(ADAccount adAccount)
        {
            return ExecResult.Success;
        }

        public ExecResult DeleteAccount(string samAccountName)
        {
            return ExecResult.Success;
        }

        public ExecResult DisableAccount(string samAccountName)
        {
            return ExecResult.Success;
        }

        public void QueryAccount(string samAccountName, string commonName)
        {
            //
        }

        #region private

        /// <summary>
        /// 获取 PacteraUsers 组织单元 LDAP
        /// </summary>
        /// <returns></returns>
        private string GetPacteraUsersOU()
        {
            var pacteraUsersOU = ConfigurationManager.AppSettings["PacteraUsersOU"];
            if (!ADHelper.ObjectExists(pacteraUsersOU, "OU"))
            {
                ADHelper.CreateOrganizeUnit(pacteraUsersOU, "");
            }
            return "OU=" + pacteraUsersOU;
        }

        protected void SetUserProperty(DirectoryEntry de, ADUser adUser)
        {
            //更新属性
            if (!string.IsNullOrEmpty(adUser.ManagerNumber) && ADHelper.IsAccExists(adUser.ManagerNumber))
            {
                var manager = ADHelper.FindObject("user", adUser.ManagerNumber);
                ADHelper.SetProperty(de, "manager", manager.Properties["distinguishedName"][0].ToString());
            }
            ADHelper.SetProperty(de, "department", adUser.Department);
            ADHelper.SetProperty(de, "company", adUser.Company);
            ADHelper.SetProperty(de, "title", adUser.PositionName);
            ADHelper.SetProperty(de, "physicalDeliveryOfficeName", adUser.OfficeLocation);
            ADHelper.SetProperty(de, "givenName", adUser.GivenName);
            ADHelper.SetProperty(de, "sn", adUser.SN);
            ADHelper.SetProperty(de, "userprincipalname", adUser.UserPrincipalName);
            ADHelper.SetProperty(de, "displayname", adUser.DisplayName);
            //ADHelper.SetProperty(de, "telephoneNumber", e.PHONE_NUMBER);
            de.CommitChanges();
            de.Close();

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

        #endregion
    }
}
