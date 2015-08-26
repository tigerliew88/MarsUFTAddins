using System;
using System.Collections.Generic;
#if tiger_dotNet4
using System.Linq;
#endif
using System.Text;
using Mercury.QTP.CustomServer;
using Route2NSEx.src.Marquis.systemUtil;
using System.Drawing;
using Infragistics.Win.UltraWinToolbars;
using System.Reflection;
using Infragistics.Win;

namespace MarsUFTAddins.IMars.tiger.infragistics.v12
{

    internal enum E_TIGERMENU_CLICK
    {
        E_OLD=0x01,
        E_STARTWITH_MENU=0x02
    }

    /*** 
     * VERSION : 1.0
     * WRITER  : GENGFENG LIU
     * DATE    : 2/10/2015
     * 
     ---------------------------------- IMPORTANT!!!! --------------------------------------------------------------------------------------
     * MarsPopupMenuControl only provids record, no replay funtion is available.
     * Use toolbarsDocArea for replaying ClickMenuItem!!!!!!!
     * Usage:
     * 1, using this class to record, script like below will be generated:
     * SwfWindow("Misys Summit FT - Bond").SwfWindow("SwfWindow").SwfWindow("SwfWindow").SwfToolbar("SwfToolbar").ClickMenuItem "74:Envelope" 
     * 2,change parent to TOOLBARSDOCARE object, like below (EXTREMLY IMPORTANT)
     * SwfWindow("Misys Summit FT - Bond").SwfToolbar("_toolDockare").ClickMenuItem "Menu:74:Envelope"
     * 
     * 注意：
     * 该类不支持回访动作
     * 要实现回访动作，需要将ClickMenuItem前面的部分改成toolbarsDockArae为对象的，这样，QTP将强制使用toolbar调用该行函数
     * 实例如下：
     * 原：
     * SwfWindow("Misys Summit FT - Bond").SwfWindow("SwfWindow").SwfWindow("SwfWindow").SwfToolbar("SwfToolbar").ClickMenuItem "74:Envelope" 
     * 修改后：
     * SwfWindow("Misys Summit FT - Bond").SwfToolbar("_toolDockare").ClickMenuItem "Menu:74:Envelope"
     ---------------------------------------------------------------------------------------------------------------------------------------
     * * ***/

    [ReplayInterface]
    public interface MenuReplay
    {
        #region Wizard generated sample code (commented)
        //		void  CustomMouseDown(int X, int Y);
        //void Press(string strItemToPress);
        void MouseDown(int x, int y);
        void ClickMenuItem(string strMenuInfo);
        #endregion
    }

    public class MarsPopupMenuControl : CustomServerBase,
        MenuReplay
    {
        private static MLogger Logger = MLogger.GetLogger(typeof(MLogger));
        protected ReflectorForCSharp mobjReflector = new ReflectorForCSharp();

        public MarsPopupMenuControl()
        {
        }

        public override void InitEventListener()
        {
            Logger.logBegin("InitEventListener");
            Delegate objDlgt = new System.Windows.Forms.MouseEventHandler(this.OnMouseDown);
            base.AddHandler(MarsTigerServerConst.CNST_EVNT_MOUSEDOWN, objDlgt);
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
                    Point pt = new Point(e.X, e.Y);
                    PopupMenuControl objListMenuControl = (PopupMenuControl)base.SourceControl;
                    
                    Logger.Info("----OnMouseDown----", objListMenuControl.GetType().ToString());

                    ControlUIElementBase objUIELement = mobjReflector.GetPrivateProperty<ControlUIElementBase>(objListMenuControl, "ControlUIElement");
                    string strPreKey = "" ;
                    string strMenuPath = "";
                    List<string> lstKeyStack = new List<string>();
#if _tigerDebug
                    List<string> lstType = new List<string>();
#endif
                    if (objUIELement.UltraControl is Infragistics.Win.UltraWinToolbars.UltraToolbarsManager)
                    {
                        Infragistics.Win.UltraWinToolbars.UltraToolbarsManager objToolMgr = (Infragistics.Win.UltraWinToolbars.UltraToolbarsManager)objUIELement.UltraControl ;
                        strPreKey = objToolMgr.ActiveTool.Key ;
                        strMenuPath = strPreKey;
                        Logger.Info("------------------",string.Format("OwningMenu:[{0}],type:[{1}]",objToolMgr.ActiveTool.OwningMenu.Key,objToolMgr.ActiveTool.OwningMenu.GetType().ToString())) ;
                        Infragistics.Win.UltraWinToolbars.PopupMenuTool objOnwerMenu = (Infragistics.Win.UltraWinToolbars.PopupMenuTool)objToolMgr.ActiveTool.OwningMenu;

                        
                        object objCurrentOwner = objToolMgr.ActiveTool.OwningMenu ;
                        while (objCurrentOwner is Infragistics.Win.UltraWinToolbars.PopupMenuTool)
                        {
                            strMenuPath = string.Format("[{1}]:[{0}]", strMenuPath, ((Infragistics.Win.UltraWinToolbars.PopupMenuTool)objCurrentOwner).Key);
                            lstKeyStack.Insert(0,((Infragistics.Win.UltraWinToolbars.PopupMenuTool)objCurrentOwner).Key) ;
#if _tigerDebug
                            lstType.Insert(0, string.Format("key:[{0}],type:[{1}]",((Infragistics.Win.UltraWinToolbars.PopupMenuTool)objCurrentOwner).Key,objCurrentOwner.GetType().ToString()));
#endif
                            if ((((PopupMenuTool)objCurrentOwner).Owner is Infragistics.Win.UltraWinToolbars.UltraToolbar))
                            {
                                string strOwnerToolbarKey = ((Infragistics.Win.UltraWinToolbars.UltraToolbar)(((PopupMenuTool)objCurrentOwner).Owner)).Key;
#if _tigerDebug
                                lstType.Insert(0, string.Format("key:[{0},type:[{1}]", ((Infragistics.Win.UltraWinToolbars.UltraToolbar)(((PopupMenuTool)objCurrentOwner).Owner)).Key, ((Infragistics.Win.UltraWinToolbars.UltraToolbar)(((PopupMenuTool)objCurrentOwner).Owner)).GetType().ToString()));
#endif
                                lstKeyStack.Insert(0, strOwnerToolbarKey);
                                strMenuPath = string.Format("[{0}]:[{1}]", strOwnerToolbarKey, strMenuPath);
                                break;
                            }
                            else
                            {
                                if (((PopupMenuTool)objCurrentOwner).OwnerIsMenu || (((PopupMenuTool)objCurrentOwner).OwningMenu is PopupMenuTool))
                                {
                                    //strMenuPath = string.Format("[{0}]:[{1}]", ((PopupMenuTool)objCurrentOwner).Key, strMenuPath);
                                    objCurrentOwner = ((PopupMenuTool)objCurrentOwner).OwningMenu;
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }

                        //Logger.Info("==================", string.Format("Owner:[{0}],type:[{1}]", objOnwerMenu.Owner, objOnwerMenu.Owner.GetType().ToString()));
                        if (lstKeyStack.Count > 0)
                        {
#if tiger_dotNet4
                            string strStrt = lstKeyStack.First<string>();
#else
                            string strStrt = lstKeyStack[0];
#endif
                            strStrt = string.Format("{0}{1}", MarsTigerServerConst.CNST_TOOLBAR_CLICK_PARA_PREFIX_MENU, strStrt);
                            lstKeyStack.RemoveAt(0);
                            lstKeyStack.Insert(0,strStrt);
                        }
                    }
                    //objUIELement.UltraControl
                    UIElement obj = objUIELement.ElementFromPoint(new Point(e.X, e.Y));
                    if (obj == null)
                    {
                        Logger.Info("OnMouseDown", string.Format("can't find object from x:{0} y:{1}", e.X, e.Y));
                        return;
                    }
                    Logger.Info("OnMouseDown", string.Format("Find object with type Name:{0}", obj.GetType().ToString()));
                    //base.RecordFunction("ClickMenuItem", RecordingMode.RECORD_SEND_LINE, new string[] { strPreKey});
                    base.RecordFunction("ClickMenuItem", RecordingMode.RECORD_SEND_LINE, lstKeyStack.ToArray());
#if _tigerDebug
                    base.RecordFunction("TypeInfo", RecordingMode.RECORD_SEND_LINE, lstType.ToArray());
#endif
                    return;
                    //if (obj is Infragistics.Win.TextUIElement)
                    //{
                    //    Infragistics.Win.TextUIElement objTxtElement = (Infragistics.Win.TextUIElement)obj;
                    //    if (objTxtElement.Parent != null)
                    //        Logger.Info("OnMouseDown",string.Format("has parent with type:{0}", objTxtElement.Parent.GetType().ToString()));
                    //    List<string> objListStack = new List<string>();
                    //    //base.RecordFunction("ClickMenuItem", RecordingMode.RECORD_SEND_LINE, new string[] { objTxtElement.Text });
                    //    GetMenuCommandStrs((UIElement)objTxtElement, ref objListStack);
                    //    objListStack.Add(objTxtElement.Text.Replace("&",""));
                    //    base.RecordFunction("ClickMenuItem", RecordingMode.RECORD_SEND_LINE, objListStack.ToArray());
                    //    return;
                    //}
                    //Logger.Info("OnMouseDown", string.Format("Not implement for object type:{0}", obj.GetType().ToString()));
                }
                else
                {
                    Logger.Info("OnMouseDown", "Not clicked by Left button, just ignored");
                }
            }
            finally
            {
                Logger.logEnd("OnMouseDown");
            }


        }

        private void GetMenuCommandStrs(UIElement objSrc, ref List<string> lstResult)
        {
            Logger.logBegin("GetMenuCommandStrs");
            if (lstResult == null)
            {
                lstResult = new List<string>();
            }

            if ((objSrc.Parent == null))
            {
                if (objSrc is Infragistics.Win.UltraWinToolbars.PopupMenuItemUIElement)
                {
                    object objCntx = ((Infragistics.Win.UltraWinToolbars.PopupMenuItemUIElement)objSrc).GetContext();
                    string strTxt = this.mobjReflector.GetMember<string>(objCntx, "Text");
                    lstResult.Insert(0, strTxt==null?"":strTxt.Replace("&",""));
                }
            }
            else
            {
                
                if ((objSrc.Parent is Infragistics.Win.UltraWinToolbars.PopupMenuItemUIElement))
                {
                    Logger.Info("----GetMenuCommandStrs----", ((Infragistics.Win.UltraWinToolbars.PopupMenuItemUIElement)objSrc).Parent.GetType().ToString());
                    GetMenuCommandStrs(objSrc.Parent, ref lstResult);
                    string strTxt = this.mobjReflector.GetMember<string>(objSrc, "Text");
                    lstResult.Add(strTxt == null ? "" : strTxt.Replace("&", ""));
                }
                else
                {
                    if ((objSrc.Parent is Infragistics.Win.UltraWinToolbars.PopupMenuItemAreaUIElement))
                    {
                        object objCntx = ((Infragistics.Win.UltraWinToolbars.PopupMenuItemAreaUIElement)objSrc.Parent).GetContext();
                        string strTxt = this.mobjReflector.GetMember<string>(objCntx, "Text");
                        lstResult.Insert(0, strTxt == null ? "" : strTxt.Replace("&", ""));
                    }
                    else
                    {
                        object objCntx = ((Infragistics.Win.UltraWinToolbars.PopupMenuItemUIElement)objSrc).GetContext();
                        string strTxt = this.mobjReflector.GetMember<string>(objCntx, "Text");
                        lstResult.Insert(0, strTxt == null ? "" : strTxt.Replace("&", ""));
                    }
                }
            }
            Logger.logEnd("GetMenuCommandStrs");
        }

        public void MouseDown(int x, int y)
        {
        }

        public void ClickMenuItem(string strMenuInfo)
        {
            Logger.logBegin("ClickMenuItem " + MarsUFTAddins.IMars.tiger.ReflectorForCSharp.MarsTigerUtility.CombinePara("strMenuInfo", strMenuInfo));
            /***
             * 有两种menu Item模式。
             * 1，Old模式，采用数字和部分key
             * 2，新模式，采用Menu开头
             * ***/
            E_TIGERMENU_CLICK eMenuMode = mGetClickMenuMode(strMenuInfo);
            switch (eMenuMode)
            {
                case E_TIGERMENU_CLICK.E_STARTWITH_MENU:
                    mProcessStartWithMenuMode(strMenuInfo.Replace(MarsTigerServerConst.CNST_TOOLBAR_CLICK_PARA_PREFIX_MENU, ""));
                    break;
                default:
                    mProcessNumberMode(strMenuInfo);
                    break;
            }


            Logger.logEnd("ClickMenuItem ");
        }

        private void mProcessNumberMode(string strMenuInfo)
        {
            Logger.logBegin("mProcessNumberMode "+MarsUFTAddins.IMars.tiger.ReflectorForCSharp.MarsTigerUtility.CombinePara("MenuInfo", strMenuInfo));
            try
            {
                //UltraToolbarsManager objTool
            }
            finally
            {
                Logger.logEnd("mProcessNumberMode");
            }
        }

        private void mProcessStartWithMenuMode(string strMenuInfo)
        {
            throw new NotImplementedException();
        }

        private E_TIGERMENU_CLICK mGetClickMenuMode(string strMenuInfo)
        {
            E_TIGERMENU_CLICK eResult = E_TIGERMENU_CLICK.E_OLD;
            if (MarsUFTAddins.IMars.tiger.ReflectorForCSharp.MarsTigerUtility.RegularExpressChecking(MarsTigerServerConst.CNST_TOOLBAR_CLICK_PARA_PREFIX_MENU_REG, strMenuInfo))
            {
                return eResult = E_TIGERMENU_CLICK.E_STARTWITH_MENU;
            }
            return eResult;
        }
    }
}
