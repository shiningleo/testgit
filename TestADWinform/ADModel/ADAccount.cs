using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ADModel
{
    public class ADAccount
    {
        public string SamAccountName { get; set; }
        public string CommonName { get; set; }
        public string Password { get; set; }
        public string UserPrincipalName { get; set; }
        /// <summary>
        /// 名
        /// </summary>
        public string FirstName { get; set; }
        /// <summary>
        /// 名
        /// </summary>
        public string GivenName { get; set; }
        public string MiddleName { get; set; }
        /// <summary>
        /// 姓
        /// </summary>
        public string LastName { get; set; }
        /// <summary>
        /// 姓
        /// </summary>
        public string SN { get; set; }
        public string DisplayName { get; set; }
        public string Company { get; set; }
        public string DepartmentNo { get; set; }
        public string Department { get; set; }
        public string RootDepartmentNo { get; set; }
        public string RootDepartmentName { get; set; }
        public string Country { get; set; }
        public string State { get; set; }
        public string City { get; set; }
        public string Location { get; set; }
        public string OfficeLocation { get; set; }
        public string StreetAddress { get; set; }
        public string Domain { get; set; }
        public string OU { get; set; }
        public string Grade { get; set; }
        public string Type { get; set; }
        public string PositionName { get; set; }
        public string ManagerNumber { get; set; }
        public string ManagerName { get; set; }
        public string PhoneNumber { get; set; }
        public string MobileNumber { get; set; }
        /// <summary>
        /// 每个账户唯一详细信息
        /// </summary>
        public string DistinguishedName { get; set; }
        /// <summary>
        /// 是否二次入职
        /// </summary>
        public bool IsRelocate { get; set; }
        public string RCreateTime { get; set; }
        public string RUpdateTime { get; set; }
        public int RFlag { get; set; }

        public ADAccount()
        {
            this.RCreateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            this.RUpdateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            this.RFlag = 0;
        }

        public void Verify()
        { 
            
        }
    }
}
