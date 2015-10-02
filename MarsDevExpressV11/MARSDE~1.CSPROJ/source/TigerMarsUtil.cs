using Route2NSEx.src.Marquis.systemUtil;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace MarsTestFrame.systemUtil
{
    public class TigerMarsUtil
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;
        public const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        public const int MOUSEEVENTF_RIGHTUP = 0x10;


        //This simulates a left mouse click
        public static void LeftMouseClick(int xpos, int ypos)
        {
            SetCursorPos(xpos, ypos);
            Thread.Sleep(50);
            mouse_event(MOUSEEVENTF_LEFTDOWN, xpos, ypos, 1, 0);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTUP, xpos, ypos, 1, 0);
        }

        public static void RightMouseClick(int xpos, int ypos)
        {
            SetCursorPos(xpos, ypos);
            Thread.Sleep(50);
            mouse_event(MOUSEEVENTF_RIGHTDOWN, xpos, ypos, 1, 0);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_RIGHTUP, xpos, ypos, 1, 0);
        }


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
