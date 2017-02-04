using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;

namespace Bussiness
{
    public class Utilities
    {
        /// <summary>
        /// 获取一个随机密码
        /// </summary>
        /// <param name="passwordLength">密码长度</param>
        /// <param name="offsetNum">如果是一次获取多个密码，那么这里的值就是循环变量的值</param>
        /// <returns></returns>
        public static string GetPassword(int passwordLength, int offsetNum = 0)
        {
            string password = string.Empty;
            long tick = DateTime.Now.Ticks + passwordLength;
            Random random = new Random((int)(tick & 0xffffffffL) | (int)(tick >> offsetNum));
            for (int i = 0; i < passwordLength; i++)
            {
                char ch;
                int num = random.Next();

                if (i == 0)
                {
                    Random ran = new Random();
                    if (ran.Next() % 2 == 0)
                    {
                        ch = (char)(0x41 + ((ushort)(num % 0x1a)));
                    }
                    else
                    {
                        ch = (char)(0x61 + ((ushort)(num % 0x1a)));
                    }
                }
                else if (i % 3 == 0)
                {
                    ch = (char)(0x30 + ((ushort)(num % 10)));
                }
                else if (i % 3 == 2)
                {
                    ushort offset = (ushort)(num % 0x05);
                    if (offset == 3)
                        offset = 29;
                    if (offset == 4)
                        offset = 60;
                    ch = (char)(0x23 + offset);
                }
                else
                {
                    if ((num % 3) == 0)
                    {
                        ch = (char)(0x30 + ((ushort)(num % 10)));
                    }
                    else if ((num % 3) == 1)
                    {
                        Random ran = new Random(i);
                        if (ran.Next() % 2 == 0)
                        {
                            ch = (char)(0x41 + ((ushort)(num % 0x1a)));
                        }
                        else
                        {
                            ch = (char)(0x61 + ((ushort)(num % 0x1a)));
                        }
                    }
                    else
                    {
                        ushort offset = (ushort)(num % 0x05);
                        if (offset == 3)
                            offset = 29;
                        if (offset == 4)
                            offset = 60;
                        ch = (char)(0x23 + offset);
                    }
                }
                password += ch.ToString();
            }

            return password;
        }
        public static int GetPasswordStrong(string password)
        {
            if (string.IsNullOrEmpty(password)) return 0;

            int level = 1;
            if (password.Length > 0 && password.Length < 7)
            {
                return level;
            }

            System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex("[a-z]+");
            if (reg.IsMatch(password)) 
                ++level;

            reg = new System.Text.RegularExpressions.Regex("[A-Z]+");
            if (reg.IsMatch(password)) 
                ++level;

            reg = new System.Text.RegularExpressions.Regex("[0-9]+");
            if (reg.IsMatch(password)) 
                ++level;

            string chr = "";
            string specialChar = "^!@#$%&*_.?";
            for (int i = 0; i < password.Length; i++)
            {
                chr = password.Substring(i, 1);
                if (specialChar.IndexOf(chr) >= 0)
                {
                    ++level;
                    break;
                }
            }
            return level;
        }

        /// <summary>
        /// 执行PowerShell脚本
        /// </summary>
        /// <param name="psFile">PowerShell脚本路径</param>
        /// <param name="adUserObject">参数对象</param>
        public static void ExecuteShell(string psFile, object adUserList)
        {
            SetExecutionPolicy("RemoteSigned");

            RunspaceConfiguration runspaceConfiguration = RunspaceConfiguration.Create();
            Runspace newspace = RunspaceFactory.CreateRunspace(runspaceConfiguration);
            newspace.Open();
            Pipeline newline = newspace.CreatePipeline();
            RunspaceInvoke scriptInvoke = new RunspaceInvoke(newspace);
            Command getCmd = new Command(psFile);
            newline.Commands.Add(getCmd);

            newspace.SessionStateProxy.SetVariable("ADUserList", adUserList);
            var getTakeres = newline.Invoke();

            //SetExecutionPolicy("Restricted");
        }

        /// <summary>
        /// 执行 PowerShell 脚本核心方法
        /// </summary>
        /// <param name="getShellcmdletList">Shell脚本集合</param>
        /// <param name="getShellParameterList">脚本中涉及对应参数</param>
        /// <returns></returns>
        public static string ExecuteShellScript(List<string> getShellcmdletList, List<ShellParameter> getShellParameterList)
        {
            string result = string.Empty;

            try
            {
                Runspace newspace = RunspaceFactory.CreateRunspace();
                Pipeline newline = newspace.CreatePipeline();

                newspace.Open();

                if (getShellcmdletList.Count > 0)
                {
                    foreach (var getShellcmdlet in getShellcmdletList)
                    {
                        newline.Commands.AddScript(getShellcmdlet);
                    }
                }

                if (getShellParameterList != null && getShellParameterList.Count > 0)
                {
                    int count = 0;
                    foreach (var getShellParameter in getShellParameterList)
                    {
                        CommandParameter cmdParameter = new CommandParameter(getShellParameter.ShellKey,
                            getShellParameter.ShellValue);
                        newline.Commands[count].Parameters.Add(cmdParameter);
                    }
                }

                var getResult = newline.Invoke();
                if (getResult != null)
                {
                    StringBuilder strBuilder = new StringBuilder();
                    foreach (var getStr in getResult)
                    {
                        strBuilder.AppendLine(getStr.ToString());
                    }
                    result = strBuilder.ToString();
                }

                newspace.Close();
            }
            catch (Exception)
            {
                throw;
            }

            return result;
        }

        /// <summary>
        /// 执行cmdlet，设置安全策略
        /// </summary>
        /// <param name="policyValue"></param>
        public static void SetExecutionPolicy(string policyValue)
        {
            List<string> getList = new List<string>();
            getList.Add("Set-ExecutionPolicy " + policyValue); //先执行启动安全策略，使系统可以执行powershell脚本文件
            ExecuteShellScript(getList, null);
        }

        /// <summary>
        /// Int32转换为Excel列名
        /// </summary>
        /// <param name="index">最小值为1</param>
        /// <returns></returns>
        public static string ExcelColumnIndexToName(int index)
        {
            if (index < 1) { throw new Exception("invalid parameter"); }

            List<string> chars = new List<string>();
            do
            {
                if (chars.Count > 0)
                    index--;
                chars.Insert(0, ((char)(index % 26 + 64)).ToString());
                index = (index - index % 26) / 26;
            } while (index > 0);

            return String.Join(string.Empty, chars.ToArray());
        }

        /// <summary>
        /// 判断文件是否存在。如果存在，则为文件添加序号，如：test(1).xlsx
        /// </summary>
        /// <param name="fileFullName">文件全名，如：D:\test.xlsx</param>
        /// <returns></returns>
        public static void DealWithExistedFileName(ref string fileFullName)
        {
            if (string.IsNullOrEmpty(fileFullName))
            {
                string dir = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                fileFullName = dir + "\\Mingming.xlsx";
            }

            if (System.IO.File.Exists(fileFullName))
            {
                int pos = fileFullName.LastIndexOf(".");
                string dotSuffix = fileFullName.Substring(pos);
                fileFullName = fileFullName.Substring(0, pos) + "(1)" + fileFullName.Substring(pos);

                pos = fileFullName.LastIndexOf(@"\");
                string path = fileFullName.Substring(0, pos + 1);
                string fileName = fileFullName.Substring(pos + 1);
                string pattern = @"\([0-9]+\)" + dotSuffix + "$";
                int i = 1;
                while (System.IO.File.Exists(fileFullName))
                {
                    i++;
                    System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(fileName, pattern);
                    if (match.Success)
                    {
                        string fileNumber = "(" + i + ")";
                        fileFullName = path + fileName.Replace(match.Value, fileNumber + dotSuffix);
                    }
                }
            }
        }
    }

    public class ShellParameter
    {
        public string ShellKey { get; set; }
        public string ShellValue { get; set; }
    }
}
