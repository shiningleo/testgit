using ADModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Text;

namespace Bussiness
{
    public class ExchangeManager
    {
        public Uri ExchangeUri { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }

        #region 为AD用户新建Exchange邮箱

        public List<ADModel.ADUser> EnableExchangeMainlbox(List<ADModel.ADUser> adUserList)
        {
            List<ADModel.ADUser> userEnabled = new List<ADModel.ADUser>();
            if (string.IsNullOrEmpty(UserName) || string.IsNullOrEmpty(Password) || ExchangeUri == null)
                return userEnabled;

            int errorCount = 0;
            try
            {
                SecureString ssRunasPassword = new SecureString();
                foreach (char x in Password)
                {
                    ssRunasPassword.AppendChar(x);
                }
                PSCredential credentials = new PSCredential(UserName, ssRunasPassword);
                WSManConnectionInfo connectionInfo = new WSManConnectionInfo(ExchangeUri, "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credentials);
                connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;

                Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo);
                runspace.Open();

                foreach (var adUser in adUserList)
                {
                    Pipeline pipeline = runspace.CreatePipeline();

                    Command commandSetADServer = new Command("Set-ADServerSettings -ViewEntireForest:$True", true);
                    pipeline.Commands.Add(commandSetADServer);

                    string cmdRemoteEmail = adUser.SamAccountName + "@pactors.mail.onmicrosoft.com";
                    string cmd = @"Get-User -Identity " + adUser.SamAccountName + " | Enable-RemoteMailbox -PrimarySmtpAddress " + adUser.UserPrincipalName + " -RemoteRoutingAddress " + cmdRemoteEmail;
                    Command command = new Command(cmd, true);
                    pipeline.Commands.Add(command);

                    ICollection<PSObject> results = pipeline.Invoke();
                    errorCount = pipeline.Error.Count;
                    //if (errorCount > 0)
                    //{
                    //    Bussiness.PowerShellAD adShell = new Bussiness.PowerShellAD();
                    //    bool isExistEmail = true;
                    //    string email = adShell.GetEmailProperty(adUser.SamAccountName);
                    //    if (string.IsNullOrEmpty(email))
                    //        isExistEmail = false;

                    //    if (isExistEmail)
                    //    {
                    //        Bussiness.LogHelper.WriteLog(new Bussiness.LogModel(Bussiness.Level.Error, DateTime.Now, "Enable-Mailbox: pipeline error count:" + pipeline.Error.Count));
                    //        errorCount = 0;
                    //    }
                    //    else
                    //    {
                    //        Bussiness.LogHelper.WriteLog(new Bussiness.LogModel(Bussiness.Level.Error, DateTime.Now, "Enable-Mailbox: " + pipeline.Error.ToString()));
                    //    }
                    //}

                    if (errorCount <= 0)
                    {
                        userEnabled.Add(adUser);
                    }

                    pipeline.Commands.Clear();
                    pipeline.Dispose();
                }

                #region 输出显示
                //if (results != null)
                //{
                //    foreach (PSObject result in results)
                //    {
                //        Console.WriteLine("------------------Get-User -Identity T000000X---------------------");
                //        Console.WriteLine("【SamAccountName】: " + result.Properties["SamAccountName"].Value);
                //        Console.WriteLine("【Name】: " + result.Properties["Name"].Value);
                //        Console.WriteLine("【DisplayName】: " + result.Properties["DisplayName"].Value);
                //        Console.WriteLine("【UserPrincipalName】: " + result.Properties["UserPrincipalName"].Value);
                //        Console.WriteLine("【WindowsEmailAddress】: " + result.Properties["WindowsEmailAddress"].Value);
                //        Console.WriteLine("【Identity】: " + result.Properties["Identity"].Value);
                //        Console.WriteLine("【OrganizationalUnit】: " + result.Properties["OrganizationalUnit"].Value);
                //        Console.WriteLine("【DistinguishedName】: " + result.Properties["DistinguishedName"].Value);
                //        //Console.WriteLine("【PrimarySmtpAddress】: " + result.Properties["PrimarySmtpAddress"].Value);
                //        //Console.WriteLine("【alias】: " + result.Properties["alias"].Value);
                //        //Console.WriteLine("【legacyexchangeDN】: " + result.Properties["legacyexchangeDN "].Value);
                //        Console.WriteLine("==================Get-User -Identity T000000X=====================");
                //    }
                //}
                //Console.WriteLine("通道错误数：" + pipeline.Error.Count);
                #endregion

                runspace.Dispose();
            }
            catch (Exception ex)
            {
                errorCount = 1;
                Bussiness.LogHelper.WriteLog(new Bussiness.LogModel(Bussiness.Level.Error, DateTime.Now, "EnableExchangeMainlbox(): " + ex.Message));
            }

            return userEnabled;
        }

        public List<ADModel.ADUser> EnableLocationMailBox(List<ADModel.ADUser> adUserList)
        {
            List<ADModel.ADUser> userEnabled = new List<ADModel.ADUser>();
            if (string.IsNullOrEmpty(UserName) || string.IsNullOrEmpty(Password) || ExchangeUri == null)
                return userEnabled;

            int errorCount = 0;
            try
            {
                SecureString ssRunasPassword = new SecureString();
                foreach (char x in Password)
                {
                    ssRunasPassword.AppendChar(x);
                }
                PSCredential credentials = new PSCredential(UserName, ssRunasPassword);
                WSManConnectionInfo connectionInfo = new WSManConnectionInfo(ExchangeUri, "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credentials);
                connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;

                Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo);
                runspace.Open();

                foreach (var adUser in adUserList)
                {
                    Pipeline pipeline = runspace.CreatePipeline();

                    Command commandSetADServer = new Command("Set-ADServerSettings -ViewEntireForest:$True", true);
                    pipeline.Commands.Add(commandSetADServer);

                    string cmd = @"Get-User -Identity " + adUser.SamAccountName + " | Enable-Mailbox";
                    Command command = new Command(cmd, true);
                    pipeline.Commands.Add(command);

                    ICollection<PSObject> results = pipeline.Invoke();
                    errorCount = pipeline.Error.Count;
                    if (errorCount > 0)
                    {
                        Bussiness.PowerShellAD adShell = new Bussiness.PowerShellAD();
                        bool isExistEmail = true;
                        string email = adShell.GetEmailProperty(adUser.SamAccountName);
                        if (string.IsNullOrEmpty(email))
                            isExistEmail = false;

                        if (isExistEmail)
                        {
                            Bussiness.LogHelper.WriteLog(new Bussiness.LogModel(Bussiness.Level.Error, DateTime.Now, "Enable-Locate-Mailbox: pipeline error count:" + pipeline.Error.Count));
                            errorCount = 0;
                        }
                        else
                        {
                            Bussiness.LogHelper.WriteLog(new Bussiness.LogModel(Bussiness.Level.Error, DateTime.Now, "Enable-Locate-Mailbox: " + pipeline.Error.ToString()));
                        }
                    }

                    if (errorCount <= 0)
                    {
                        userEnabled.Add(adUser);
                    }

                    pipeline.Commands.Clear();
                    pipeline.Dispose();
                }

                runspace.Dispose();
            }
            catch (Exception ex)
            {
                errorCount = 1;
                Bussiness.LogHelper.WriteLog(new Bussiness.LogModel(Bussiness.Level.Error, DateTime.Now, "EnableLocationMailBox(): " + ex.Message));
            }

            return userEnabled;
        }

        public bool EnableExchangeMainlboxByCertificate(List<ADModel.ADUser> adUserList)
        {
            if (ExchangeUri == null)
                return false;

            int errorCount = 0;
            try
            {
                SecureString ssRunasPassword = new SecureString();
                foreach (char x in Password)
                {
                    ssRunasPassword.AppendChar(x);
                }
                //PSCredential credentials = new PSCredential(UserName, ssRunasPassword);
                string thumbprint = "‎‎‎50 2d fd 2b e4 c0 e5 a5 ab c1 6f c0 fe 0c de 97 3e b0 a8 6f";
                //string thumbprint = "‎‎‎*INVISIBLECHARACTER*502dfd2be4c0e5a5abc16fc0fe0cde973eb0a86f"; 
                WSManConnectionInfo connectionInfo = new WSManConnectionInfo(ExchangeUri, "http://schemas.microsoft.com/powershell/Microsoft.Exchange", thumbprint);
                //connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;

                Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo);
                runspace.Open();

                Pipeline pipeline = runspace.CreatePipeline();
                foreach (var adUser in adUserList)
                {
                    #region test
                    //string cmd = @"Get-User -Identity " + adUser.SamAccountName + " | Enable-Mailbox"; //@"Get-User -Identity T0000006 | Enable-Mailbox";
                    //Command command = new Command(cmd, true);
                    ////command.Parameters.Add("Identity", adUser.SamAccountName);
                    //pipeline.Commands.Add(command);
                    #endregion

                    #region zs

                    string cmdSetADServer = "Set-ADServerSettings -ViewEntireForest:$True";
                    Command commandSetADServer = new Command(cmdSetADServer, true);
                    pipeline.Commands.Add(commandSetADServer);

                    string cmdRemoteEmail = adUser.SamAccountName + "@pactors.mail.onmicrosoft.com";
                    //Command commandRemoteEmail = new Command(cmdRemoteEmail, true);
                    //pipeline.Commands.Add(commandRemoteEmail);

                    string cmd = @"Get-User -Identity " + adUser.SamAccountName + " | Enable-RemoteMailbox -PrimarySmtpAddress " + adUser.UserPrincipalName + " -RemoteRoutingAddress " + cmdRemoteEmail;
                    //string cmd = @"Get-User -Identity " + adUser.SamAccountName + " | Enable-Mailbox"; //@"Get-User -Identity T0000006 | Enable-Mailbox";
                    Command command = new Command(cmd, true);
                    //command.Parameters.Add("Identity", adUser.SamAccountName);
                    pipeline.Commands.Add(command);

                    #endregion
                }

                ICollection<PSObject> results = pipeline.Invoke();
                errorCount = pipeline.Error.Count;
                #region 输出显示
                //if (errorCount > 0)
                //{
                //    foreach (PSObject result in results)
                //    {
                //        Console.WriteLine("------------------Get-User -Identity T000000X---------------------");
                //        Console.WriteLine("【SamAccountName】: " + result.Properties["SamAccountName"].Value);
                //        Console.WriteLine("【Name】: " + result.Properties["Name"].Value);
                //        Console.WriteLine("【DisplayName】: " + result.Properties["DisplayName"].Value);
                //        Console.WriteLine("【UserPrincipalName】: " + result.Properties["UserPrincipalName"].Value);
                //        Console.WriteLine("【WindowsEmailAddress】: " + result.Properties["WindowsEmailAddress"].Value);
                //        Console.WriteLine("【Identity】: " + result.Properties["Identity"].Value);
                //        Console.WriteLine("【OrganizationalUnit】: " + result.Properties["OrganizationalUnit"].Value);
                //        Console.WriteLine("【DistinguishedName】: " + result.Properties["DistinguishedName"].Value);
                //        //Console.WriteLine("【PrimarySmtpAddress】: " + result.Properties["PrimarySmtpAddress"].Value);
                //        //Console.WriteLine("【alias】: " + result.Properties["alias"].Value);
                //        //Console.WriteLine("【legacyexchangeDN】: " + result.Properties["legacyexchangeDN "].Value);
                //        Console.WriteLine("==================Get-User -Identity T000000X=====================");
                //    }
                //}
                //Console.WriteLine("通道错误数：" + pipeline.Error.Count);
                #endregion
                runspace.Dispose();
            }
            catch (Exception ex)
            {
                errorCount = 1;
                Bussiness.LogHelper.WriteLog(new Bussiness.LogModel(Bussiness.Level.Error, DateTime.Now, "EnableExchangeMainlbox(): " + ex.Message));
            }

            if (errorCount <= 0)
                return true;
            else
                return false;
        }

        #endregion

        private void CreateRemotePolicy(Pipeline pipelinePolicy, List<ADModel.ADUser> adUserList)
        {
            foreach (var adUser in adUserList)
            {
                string cmdText = string.Empty;
                Command commandFalse = null;
                Command commandTrue = null;
                if ((adUser.Type != null && adUser.Type.Trim().ToUpper() == "E1") || (adUser.Grade != null && adUser.Grade.Trim().ToUpper() == "E1"))
                {
                    cmdText = "Get-RemoteMailbox -Identity " + adUser.SamAccountName + " |  Set-RemoteMailbox -EmailAddressPolicyEnabled $false -CustomAttribute11 E1";
                    commandFalse = new Command(cmdText, true);
                    pipelinePolicy.Commands.Add(commandFalse);

                    cmdText = "Get-RemoteMailbox -Identity " + adUser.SamAccountName + " |  Set-RemoteMailbox -EmailAddressPolicyEnabled $True -CustomAttribute11 E1";
                    commandTrue = new Command(cmdText, true);
                    pipelinePolicy.Commands.Add(commandTrue);
                }
                else if ((adUser.Type != null && adUser.Type.Trim().ToUpper() == "E3") || (adUser.Grade != null && adUser.Grade.Trim().ToUpper() == "E3"))
                {
                    cmdText = "Get-RemoteMailbox -Identity " + adUser.SamAccountName + " |  Set-RemoteMailbox -EmailAddressPolicyEnabled $false";
                    commandFalse = new Command(cmdText, true);
                    pipelinePolicy.Commands.Add(commandFalse);

                    cmdText = "Get-RemoteMailbox -Identity " + adUser.SamAccountName + " |  Set-RemoteMailbox -EmailAddressPolicyEnabled $True";
                    commandTrue = new Command(cmdText, true);
                    pipelinePolicy.Commands.Add(commandTrue);
                }
                else
                {
                    Bussiness.LogHelper.WriteLog(new Bussiness.LogModel(Bussiness.Level.Error, DateTime.Now, adUser.SamAccountName + "指定了错误的许可授权,没有成功开启邮箱地址策略"));
                }
                if (commandFalse != null)
                    pipelinePolicy.Commands.Add(commandFalse);
                if (commandTrue != null)
                    pipelinePolicy.Commands.Add(commandTrue); ;

                Command command = null;
                if (adUser.Country != null && adUser.Country.Trim().ToUpper() == "CN")
                {
                    cmdText = "Add-DistributionGroupMember 'everyone_china_v' -Member " + adUser.SamAccountName;
                    command = new Command(cmdText, true);
                }
                else
                {
                    cmdText = "Add-DistributionGroupMember 'everyone_oversea' -Member " + adUser.SamAccountName;
                    command = new Command(cmdText, true);
                }
                if (command != null)
                    pipelinePolicy.Commands.Add(command);
            }


            ICollection<PSObject> results = null;
            if (pipelinePolicy.Commands.Count > 0)
                results = pipelinePolicy.Invoke();
            int errorCount = pipelinePolicy.Error.Count;

            if (errorCount > 0)
                Bussiness.LogHelper.WriteLog(new Bussiness.LogModel(Bussiness.Level.Error, DateTime.Now, "创建策略时有错误产生：" + errorCount));
        }

        #region 调用方法

        public List<ADUser> GetADUserList(string connectionString)
        {
            string cmdText = @"select isnull(SamAccountName,'') as SamAccountName
                                ,isnull(Password,'') as Password
                                ,isnull(UserPrincipalName,'') as UserPrincipalName
                                ,isnull(FirstName,'') as FirstName
                                ,isnull(LastName,'') as LastName
                                ,isnull(DisplayName,'') as DisplayName
                                ,isnull(Department,'') as Department
                                ,isnull(StreetAddress,'') as StreetAddress
                                ,isnull(City,'') as City
                                ,isnull(State,'') as State
                                ,isnull(Country,'') as Country
                                ,isnull(Domain,'') as Domain
                                ,isnull(OU,'') as OU
                                ,isnull(Grade,'') as Grade
                                ,isnull(Type,'') as Type
                                ,isnull(RCreateTime,'') as RCreateTime
                                ,isnull(RUpdateTime,'') as RUpdateTime
                                ,isnull(RFlag,'') as RFlag 
                                from T_ADUserInfo 
                                where RFlag = 0 and RCreateTime < '" + DateTime.Now.AddMinutes(-29).ToString() + "'";
            SqlDataReader reader = SqlHelper.ExecuteReader(connectionString, CommandType.Text, cmdText);
            List<ADUser> adUsers = new List<ADUser>();
            while (reader.Read())
            {
                ADUser adUser = new ADUser();
                adUser.SamAccountName = reader["SamAccountName"].ToString();
                adUser.Password = reader["Password"].ToString();
                adUser.UserPrincipalName = reader["UserPrincipalName"].ToString();
                adUser.FirstName = reader["FirstName"].ToString();
                adUser.LastName = reader["LastName"].ToString();
                adUser.DisplayName = reader["DisplayName"].ToString();
                adUser.Department = reader["Department"].ToString();
                adUser.StreetAddress = reader["StreetAddress"].ToString();
                adUser.City = reader["City"].ToString();
                adUser.State = reader["State"].ToString();
                adUser.Country = reader["Country"].ToString();
                adUser.Domain = reader["Domain"].ToString();
                adUser.OU = reader["OU"].ToString();
                adUser.Type = reader["Type"].ToString();
                adUser.RCreateTime = reader["RCreateTime"].ToString();
                adUser.RUpdateTime = reader["RUpdateTime"].ToString();
                adUser.RFlag = reader["RFlag"].ToString();
                adUsers.Add(adUser);
            }
            return adUsers;
        }

        public void UpdateInfo(string connectionString, List<ADUser> adUsers)
        {
            foreach (ADUser adUser in adUsers)
            {
                try
                {
                    string commandText = "Update T_ADUserInfo Set RFlag = 1,RUpdateTime = '" + DateTime.Now.ToString() + "' where SamAccountName = N'" + adUser.SamAccountName + "' and RFlag = 0";
                    SqlHelper.ExecuteNonQuery(connectionString, CommandType.Text, commandText);
                }
                catch
                {
                    Bussiness.LogHelper.WriteLog(new Bussiness.LogModel(Bussiness.Level.Error, DateTime.Now, "EnableEmailBox: Update RFlag of " + adUser.SamAccountName + " failure."));
                }
            }
        }

        #endregion

        //public List<LDAP> LDAPStore { get; set; }

        //#region 为AD用户新建Exchange邮箱

        //public void CreateExchangeMainlbox(string psFile, List<ADUser> adUserList)
        //{
        //    string runasUsername = string.Empty;
        //    string runasPassword = string.Empty;
        //    Uri uri = null;

        //    string exId = ConfigurationManager.AppSettings["ExchangeLDAPID"];
        //    var ldap = LDAPStore.FirstOrDefault(l => l.ID == Convert.ToInt32(exId));
        //    if (ldap != null)
        //    {
        //        runasUsername = ldap.ADUser;
        //        runasPassword = ldap.ADPassword;
        //        uri = new Uri(ldap.ADPath.Replace("LDAP", "http") + "/PowerShell");
        //    }

        //    if (string.IsNullOrEmpty(runasUsername) || string.IsNullOrEmpty(runasPassword) || uri == null)
        //        return;

        //    SecureString ssRunasPassword = new SecureString();
        //    foreach (char x in runasPassword)
        //    {
        //        ssRunasPassword.AppendChar(x);
        //    }
        //    PSCredential credentials = new PSCredential(runasUsername, ssRunasPassword);
        //    WSManConnectionInfo connectionInfo = new WSManConnectionInfo(uri, "http://schemas.microsoft.com/powershell/Microsoft.Exchange", credentials);
        //    connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;

        //    Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo);
        //    runspace.Open();

        //    Pipeline pipeline = runspace.CreatePipeline();
        //    foreach (var adUser in adUserList)
        //    {
        //        string cmd = @"Get-User -Identity " + adUser.SamAccountName; //@"Get-User -Identity T0000006 | Enable-Mailbox";
        //        Command command = new Command(cmd, true);
        //        //command.Parameters.Add("resultsize", "unlimited");
        //        pipeline.Commands.Add(command);
        //    }

        //    ICollection<PSObject> results = pipeline.Invoke();

        //    if (results != null)
        //    {
        //        foreach (PSObject result in results)
        //        {
        //            Console.WriteLine("------------------Get-User -Identity T0000006---------------------");
        //            Console.WriteLine("【SamAccountName】: " + result.Properties["SamAccountName"].Value);
        //            Console.WriteLine("【Name】: " + result.Properties["Name"].Value);
        //            Console.WriteLine("【DisplayName】: " + result.Properties["DisplayName"].Value);
        //            Console.WriteLine("【UserPrincipalName】: " + result.Properties["UserPrincipalName"].Value);
        //            Console.WriteLine("【WindowsEmailAddress】: " + result.Properties["WindowsEmailAddress"].Value);
        //            Console.WriteLine("【Identity】: " + result.Properties["Identity"].Value);
        //            Console.WriteLine("【OrganizationalUnit】: " + result.Properties["OrganizationalUnit"].Value);
        //            Console.WriteLine("【DistinguishedName】: " + result.Properties["DistinguishedName"].Value);
        //            //Console.WriteLine("【PrimarySmtpAddress】: " + result.Properties["PrimarySmtpAddress"].Value);
        //            //Console.WriteLine("【alias】: " + result.Properties["alias"].Value);
        //            //Console.WriteLine("【legacyexchangeDN】: " + result.Properties["legacyexchangeDN "].Value);
        //            Console.WriteLine("==================Get-User -Identity T0000006=====================");
        //        }
        //    }

        //    Console.WriteLine("通道错误数：" + pipeline.Error.Count);
        //    runspace.Dispose();
        //}

        //#endregion

        ///// <summary>
        ///// 临时使用
        ///// </summary>
        ///// <param name="adUser"></param>
        //public void ReadADUserInfo(ADUser adUser)
        //{
        //    string conn = "server=172.16.253.172;database=ADDB;uid=sa;pwd=1234-abcd";
        //    string sql = "select * from T_ADUserInfo where SamAccountName = N'" + adUser.SamAccountName + "'";
        //    System.Data.SqlClient.SqlDataReader reader = SqlHelper.ExecuteReader(conn, CommandType.Text, sql);
        //    if (reader.Read())
        //    {
        //        adUser.CommonName = reader["CommonName"].ToString();
        //        adUser.Password = reader["Password"].ToString();
        //        adUser.UserPrincipalName = reader["UserPrincipalName"].ToString();
        //        adUser.DisplayName = reader["DisplayName"].ToString();
        //        adUser.Company = reader["Company"].ToString();
        //        adUser.Department = reader["Department"].ToString();
        //        adUser.Country = reader["Country"].ToString();
        //        adUser.State = reader["State"].ToString();
        //        adUser.OfficeLocation = reader["OfficeLocation"].ToString();
        //        adUser.StreetAddress = reader["StreetAddress"].ToString();
        //        adUser.Domain = reader["Domain"].ToString();
        //        adUser.OU = reader["OU"].ToString();
        //        adUser.Type = reader["Type"].ToString();
        //        adUser.ManagerNumber = reader["ManagerNumber"].ToString();
        //    }
        //}
    }
}
