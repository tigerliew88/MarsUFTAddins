using Route2NSEx.src.Marquis.systemUtil;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace MarsTestFrame.systemUtil
{
    public class TigerMarsUtil
    {
        private static MLogger Logger = MLogger.GetLogger(typeof(TigerMarsUtil));
        public static string GetPathWithoutFileName(string strFileWithPath)
        {
            Logger.logBegin("GetPathWithoutFileName");
            try
            {
                if (strFileWithPath == null) return null;

                int iLastPos = strFileWithPath.LastIndexOf("\\");
                if (iLastPos == -1)
                {
                    return null;
                }

                return strFileWithPath.Substring(0, iLastPos);

            }
            finally
            {
                Logger.logEnd("GetPathWithoutFileName");

            }
        }

        public static string GetParameter(string strParaName, string strValue)
        {
            return string.Format(" ,[{0}={1}] ", strParaName, strValue);
        }

        public static string GetParameter(string[] arrParaName, string[] strValues)
        {
            string strFormat = "";
            int iMaxLen = arrParaName == null ? -1 : arrParaName.Length;
            iMaxLen = Math.Max(iMaxLen, strValues == null ? -1 : strValues.Length);
            for (int i=0;i<iMaxLen;i++)
            {
                strFormat = string.Format("{0},[{1}={2}]", strFormat, arrParaName[i], strValues[i]);
            }
            return strFormat;
        }

        public static bool RegularTest(string strPartern, string strValue)
        {
            RegexOptions options = RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace ;
            return Regex.IsMatch(strValue, strPartern, options);            
        }

        public static int ConvertObject2Int(object objSrc, int iDefault=-1)
        {
            Logger.Info("ConvertObject2Int", string.Format("{0} {1}", GetParameter("objSrc", objSrc==null?"":objSrc.ToString()), iDefault));
            int iResult = iDefault;
            int.TryParse(objSrc==null?"NULL": objSrc.ToString().Trim(),out iResult);
            return iResult;
        }
    }
}
