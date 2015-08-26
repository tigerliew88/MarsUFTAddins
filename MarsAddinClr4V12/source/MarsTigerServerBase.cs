using System;
using System.Collections.Generic;
#if tiger_dotNet4
using System.Linq;
#endif
using System.Text;
using Mercury.QTP.CustomServer;
using Route2NSEx.src.Marquis.systemUtil;
using System.Windows.Forms;
using System.Threading;

namespace MarsUFTAddins.IMars.tiger.infragistics.v12
{

    public interface IMarsTigerReplayBase
    {
        
    }

    public class MarsTigerServerBase : CustomServerBase
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;
        public const int MOUSEEVENTF_RIGHTDOWN = 0x0008;
        public const int MOUSEEVENTF_RIGHTUP = 0x0010;

        //This simulates a left mouse click
        public static void LeftMouseClick(int xpos, int ypos)
        {
            SetCursorPos(xpos, ypos);
            mouse_event(MOUSEEVENTF_LEFTDOWN, xpos, ypos, 0, 0);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTUP, xpos, ypos, 0, 0);
        }

        public void RightMouseClick(int xPos, int yPos)
        {
            SetCursorPos(xPos, yPos);
            mouse_event(MOUSEEVENTF_RIGHTDOWN, xPos, yPos, 1, 0);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_RIGHTUP, xPos, yPos, 1, 0);
        }

        private static MLogger Logger = MLogger.GetLogger(typeof(MarsTigerServerBase));

        protected ReflectorForCSharp mobjReflector = new ReflectorForCSharp();

        public override void InitEventListener()
        {
            Logger.logBegin("InitEventListener");
            AddEvent();
            Logger.logEnd("InitEventListener");
        }

        protected void QtpShowException(string strException)
        {
#if _ShowQTPException
            base.ReplayThrowError(strException);
#endif
        }

        protected virtual void AddEvent()
        {
            Logger.logBegin("AddEvent");
#if _tigerDebugMouseDown
            Delegate eMouseDown = new System.Windows.Forms.MouseEventHandler(this.tigerMouseDown);
            base.AddHandler(MarsTigerServerConst.CNST_EVNT_MOUSEDOWN, eMouseDown);
#endif
            Logger.logEnd("AddEvent");
        }

        public override void ReleaseEventListener()
        {
            
        }

        protected virtual void tigerMouseDown(object objSender, MouseEventArgs e)
        {
            Logger.logBegin("tigerMouseDown");
#if _tigerDebug
            Logger.Info("tigerMouseDown", string.Format("Current objects type:{0}", base.SourceControl.GetType().ToString()));
            Logger.Info("tigerMouseDown", string.Format("Current objects base type:{0}", base.SourceControl.GetType().BaseType.ToString()));

#endif
#if _tigerEvent
            List<string> lstEvntName = new List<string>() ;
            this.mobjReflector.GetAllEventsName(base.SourceControl,ref lstEvntName);
            foreach (string strEvnt in lstEvntName)
            {
                Logger.Info("****Test****", strEvnt);
            }
            lstEvntName.Clear();
            lstEvntName = null;
#endif
            Logger.logEnd("tigerMouseDown");
        }
    }
}
