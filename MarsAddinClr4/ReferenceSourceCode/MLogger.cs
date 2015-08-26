#define _LOG4NET

using System;
using System.Collections.Generic;
#if tiger_dotNet4
using System.Linq;
#endif
using System.Text;
using System.IO;

namespace Route2NSEx.src.Marquis.systemUtil
{
    public class MLogger
    {

        private const string LOGGER_NAME = "MarsAddins";

        private string mstrCurrentClassName;

#if _LOG4NET
        private log4net.ILog mobjLog = log4net.LogManager.GetLogger(LOGGER_NAME);
        private static bool ISLoad = false;

        public static MLogger GetLogger(string className)
        {
            if (!ISLoad)
            {
                log4net.Config.XmlConfigurator.Configure(new System.IO.FileInfo(typeof(MLogger).Assembly.Location + ".config"));
                ISLoad = true;
            }
            MLogger objResult = new MLogger() { mstrCurrentClassName = className };
            return objResult;
        }

        public static MLogger GetLogger(Type oneType)
        {
            if (!ISLoad)
            {
                log4net.Config.XmlConfigurator.Configure(new System.IO.FileInfo(typeof(MLogger).Assembly.Location + ".config"));
                ISLoad = true;
            }
            MLogger objResult = new MLogger() { mstrCurrentClassName = oneType.ToString() };
            return objResult;
        }

        public void logBegin(string strMethodName)
        {
            //mobjLog.Info(string.Format("[INFO] {0:MM/dd/yyyy H:mm:ss zzz} {1}.{2} begins...", DateTime.Now, mstrCurrentClassName, strMethodName));
            mobjLog.Info(string.Format("[INFO] {0}.{1} begins...", mstrCurrentClassName, strMethodName));
        }

        public void logEnd(string strMethodName)
        {
            //mobjLog.Info(string.Format("[INFO] {0:MM/dd/yyyy H:mm:ss zzz} {1}.{2} end.", DateTime.Now, mstrCurrentClassName, strMethodName));
            mobjLog.Info(string.Format("[INFO] {0}.{1} end.", mstrCurrentClassName, strMethodName));
        }

        public void Info(string strMethodName, string strInfo)
        {
            //mobjLog.Info(string.Format("[INFO] {0:MM/dd/yyyy H:mm:ss zzz} {1}.{2} {3}", DateTime.Now, mstrCurrentClassName, strMethodName, strInfo));
            mobjLog.Info(string.Format("[INFO] {0}.{1} {2}", mstrCurrentClassName, strMethodName, strInfo));
        }

        public void Error(string strMethodName, string strErrorMsg)
        {
            //mobjLog.Error(string.Format("[ERROR] {0:MM/dd/yyyy H:mm:ss zzz} {1}.{2} {3}", DateTime.Now, mstrCurrentClassName, strMethodName, strErrorMsg));
            mobjLog.Error(string.Format("[ERROR] {0}.{1} {2}", mstrCurrentClassName, strMethodName, strErrorMsg));
        }
        public void Error(string strMethodName, string strErrorMsg, Exception e)
        {
            //mobjLog.Error(string.Format("[ERROR] {0:MM/dd/yyyy H:mm:ss zzz} {1}.{2} {3}", DateTime.Now, mstrCurrentClassName, strMethodName, strErrorMsg), e);
            mobjLog.Error(string.Format("[ERROR] {0}.{1} {2}", mstrCurrentClassName, strMethodName, strErrorMsg), e);
        }
#else
        /*** text mode ***/
        protected static StreamWriter gLogFile = null;

        protected void initLogFile()
        {
            if (gLogFile==null)
                gLogFile = new StreamWriter(string.Format(@"c:\temp\tiger_log_{0:yyyyMMdd}.log", DateTime.Now)) ;
        }

        public static MLogger GetLogger(string className)
        {
            MLogger objResult = new MLogger() { mstrCurrentClassName = className };
            return objResult;
        }

        public static MLogger GetLogger(Type oneType)
        {
            MLogger objResult = new MLogger() { mstrCurrentClassName = oneType.ToString() };
            return objResult;
        }

        public void logBegin(string strMethodName)
        {
            initLogFile() ;
            gLogFile.WriteLine(string.Format("[INFO] [{0:MM/dd/yyyy H:mm:ss zzz}] {1}.{2} begin...", DateTime.Now, this.mstrCurrentClassName, strMethodName));
            gLogFile.Flush();
        }
        public void logEnd(string strMethodName)
        {
            initLogFile();
            gLogFile.WriteLine(string.Format("[INFO] [{0:MM/dd/yyyy H:mm:ss zzz}] {1}.{2} End...", DateTime.Now, this.mstrCurrentClassName, strMethodName));
            gLogFile.Flush();
        }
        public void Info(string strMethodName, string strInfo)
        {
            initLogFile();
            gLogFile.WriteLine(string.Format("[INFO] [{0:MM/dd/yyyy H:mm:ss zzz}] {1}.{2} {3}", DateTime.Now, this.mstrCurrentClassName, strMethodName, strInfo));
            gLogFile.Flush();
        }
        public void Error(string strMethodName, string strErrorMsg)
        {
            initLogFile();
            gLogFile.WriteLine(string.Format("[ERROR] [{0:MM/dd/yyyy H:mm:ss zzz}] {1}.{2} \r\n\t{3}", DateTime.Now, this.mstrCurrentClassName, strMethodName, strErrorMsg));
            gLogFile.Flush();
        }
        public void Error(string strMethodName, string strErrorMsg, Exception e)
        {
            initLogFile();
            gLogFile.WriteLine(string.Format("[ERROR] [{0:MM/dd/yyyy H:mm:ss zzz}] {1}.{2} \r\n\t{3} \r\n\t Exceptions:\r\n\t\t{4}", DateTime.Now, this.mstrCurrentClassName, strMethodName, strErrorMsg, e.StackTrace.ToString()));
            gLogFile.Flush();
        }
#endif
    }

}
