﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices.Protocols;
using System.Net;
using System.DirectoryServices;

namespace TestADWinform
{
   public  class ConnectLdap//: DirectoryConnection, IDisposable
    {

        static LdapConnection ldapConnection;
        static string ldapServer;
        static NetworkCredential credential;
        static string targetOU;
        static string pwd;
        public   void LdapBind()
        {
            ldapServer = "172.16.253.";
            targetOU = "cn=Administrators,dc=,dc=";//cn=Manager,cn=Builtin,
            pwd = "";

            credential = new NetworkCredential(String.Empty, String.Empty);
            //credential = new NetworkCredential(targetOU, pwd);

             
            string dn = "";

            ldapConnection = new LdapConnection(new LdapDirectoryIdentifier(ldapServer));
            ldapConnection.SessionOptions.ProtocolVersion = 3;//Ldap协议版本
           ldapConnection.AuthType = AuthType.Anonymous;//不传递密码进行连接

            //ldapConnection = new LdapConnection(ldapServer);
            //ldapConnection.AuthType = AuthType.Basic;
            //ldapConnection.Credential = credential;

            try
            {
                Console.WriteLine("链接.");
                ldapConnection.Bind();
                Console.WriteLine("链接成功");

            }
            catch (Exception ee)
            {
               Console.WriteLine(ee.Message);
            }


            ldapConnection.Dispose();
        
        }


//如果我们使用ldapConnection.AuthType = AuthType.Anonymous; 的认证方式，就一定要让Dn与Pwd为空，实现匿名认证方式，如：

//credential = new NetworkCredential(String.Empty, String.Empty);


    
public void Dispose()
{
 	throw new NotImplementedException();
}
}
}
