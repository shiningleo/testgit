using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;

namespace Bussiness
{
    public static class LogHelper
    {
        /// <summary>
        /// 记录日志
        ///LogHelper.WriteLog("Message");
        /// </summary>
        /// <param name="log">日志信息</param>
        public static void WriteLog(LogModel log)
        {
            var msg = log.Date.ToString("hh:mm:ss") + " [" + log.Level + "]:" + log.Description;
            var filePath = ConfigurationManager.AppSettings["LogPath"] + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            if (!File.Exists(filePath))
            {
                StreamWriter sw = new StreamWriter(filePath, true, Encoding.UTF8);
                sw.WriteLine(msg);
                sw.Close();
            }
            else 
            {
                StreamWriter sw = new StreamWriter(filePath, true, Encoding.UTF8);
                sw.WriteLine(msg);
                sw.Close();
            }
        }
        public static void WriteLogRepeatUser(LogModel log)
        {
            var msg = log.Date.ToString("hh:mm:ss") + " [" + log.Level + "]:" + log.Description;
            var filePath = ConfigurationManager.AppSettings["LogPath"] + "RepeatADUser.txt";
            if (!File.Exists(filePath))
            {
                StreamWriter sw = new StreamWriter(filePath, true, Encoding.UTF8);
                sw.WriteLine(msg);
                sw.Close();
            }
            else
            {
                StreamWriter sw = new StreamWriter(filePath, true, Encoding.UTF8);
                sw.WriteLine(msg);
                sw.Close();
            }
        }
        public static void WriteRemotePolicyLog(LogModel log)
        {
            var msg = log.Date.ToString("hh:mm:ss") + " [" + log.Level + "]:" + log.Description;
            var filePath = ConfigurationManager.AppSettings["RemotePolicyLogPath"] + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            if (!File.Exists(filePath))
            {
                StreamWriter sw = new StreamWriter(filePath, true, Encoding.UTF8);
                sw.WriteLine(msg);
                sw.Close();
            }
            else
            {
                StreamWriter sw = new StreamWriter(filePath, true, Encoding.UTF8);
                sw.WriteLine(msg);
                sw.Close();
            }
        }
    }

    /// <summary>
    /// 日志级别
    /// </summary>
    public enum Level
    {
        Info = 1,
        Error = 2
    }

    /// <summary>
    /// 日志实体
    /// </summary>
    public class LogModel
    {
        public LogModel(Level level, DateTime date, string description)
        {
            Level = level;
            Date = date;
            Description = description;
        }

        public Level Level { get; set; }
        public DateTime Date { get; set; }
        public string Description { get; set; }
    }
}
