using System;
using System.Collections.Generic;
using System.Text;
using Mercury.QTP.CustomServer;
using Route2NSEx.src.Marquis.systemUtil;
using Infragistics.Win.UltraWinTabControl;
using Infragistics.Win;
using System.Windows.Forms;

namespace MarsUFTAddins.IMars.tiger.infragistics.v12
{

    [ReplayInterface]
    public interface IMarsTab
    {
        void CloseTab();
    }

    public class MarsTabControl:CustomServerBase,
        IMarsTab
    {
        private static MLogger Logger = MLogger.GetLogger(typeof(MarsTabControl));
        protected ReflectorForCSharp mobjReflector = new ReflectorForCSharp();

        public MarsTabControl()
        {
        }

        public override void InitEventListener()
        {
            Logger.logBegin("InitEventListener");
            Delegate objDlgt = new System.Windows.Forms.MouseEventHandler(this.OnMouseDown);
            base.AddHandler(MarsTigerServerConst.CNST_EVNT_MOUSEDOWN, objDlgt);
            Logger.logEnd("InitEventListener");
        }

        public override void ReleaseEventListener()
        {
            
        }

        public void OnMouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            Logger.logBegin("OnMouseDown");
            try
            {
                if (e.Button == System.Windows.Forms.MouseButtons.Left)
                {
                    Logger.Info("OnMouseDown****", base.SourceControl.GetType().BaseType.MakeArrayType().ToString());
                    Infragistics.Win.UltraWinTabControl.UltraTabControl objTab = (Infragistics.Win.UltraWinTabControl.UltraTabControl)base.SourceControl;
                    //objTab.GetType().BaseType.ToString()
                    //Infragistics.Win.UltraWinTabControl.UltraTabControl objTab = (Infragistics.Win.UltraWinTabControl.UltraTabControl)base.SourceControl;

                    base.RecordFunction("ActiveTab", RecordingMode.RECORD_SEND_LINE, objTab.ActiveTab.Text);
                }
            }
            finally
            {

            }
        }

        public void CloseTab()
        {
            Logger.logBegin("CloseTab");
            PrepareForReplay();
            var obj = base.SourceControl;
            if (obj is Infragistics.Win.UltraWinTabControl.UltraTabControl)
            {
                Infragistics.Win.UltraWinTabControl.UltraTabControl objTab = (Infragistics.Win.UltraWinTabControl.UltraTabControl)obj;
                
                UltraTabControlUIElement objUI = objTab.UIElement;
                TabHeaderAreaUIElement objTabHeader = (TabHeaderAreaUIElement)mobjReflector.GetMember<TabHeaderAreaUIElement>(objUI, "TabAreaUIElement");
                //objTabHeader.ChildElements
                //UltraTab objTab.ActiveTab
                MouseMove(objTabHeader.Rect.Left, objTabHeader.Rect.Top);
                MessageBox.Show(string.Format("Original position:[x:{0}, y:{1}]", objTabHeader.Rect.Left, objTabHeader.Rect.Top));

               // objTab.TabMa
            }
            else
            {
            }
        }
    }
}
