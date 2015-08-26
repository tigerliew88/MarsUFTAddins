using System;
using Mercury.QTP.CustomServer;
using MarsUFTAddins.IMars.tiger;
using System.Windows.Forms;
using Route2NSEx.src.Marquis.systemUtil;
using System.Drawing;
using Infragistics.Win.UltraWinToolbars;
using Infragistics.Win;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Collections;
using System.Reflection;
using System.Threading;
using Infragistics.Shared;
using MarsTestFrame.systemUtil;



namespace MarsUFTAddins.IMars.tiger.infragistics.v12
{
    public enum E_ClickToolbarMode
    {
        e_Defautl_from_TA = 0x0,
        e_ClickButton,
        e_ClickMenu
    }


    [ReplayInterface]
    public interface ToolbarReplay
    {
        #region Wizard generated sample code (commented)
        //		void  CustomMouseDown(int X, int Y);
        void Press(string strItemToPress);
        void MouseDown(int x, int y);
        void ClickToolbarTool(params object[] arrStrKeys);
        void ClickMenuItem(params object[] strMenuInfo);
        bool Wait4Validate();
        void WriteCursorStatus();
        #endregion
    }
    /// <summary>
    /// Summary description for Mars.tiger.infragistics.v12.toolbar.
    /// </summary>
    public class Toolbar :
        CustomServerBase,
        ToolbarReplay
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;

        //This simulates a left mouse click
        public static void LeftMouseClick(int xpos, int ypos)
        {
            SetCursorPos(xpos, ypos);
            Thread.Sleep(50);
            mouse_event(MOUSEEVENTF_LEFTDOWN, xpos, ypos, 1, 0);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTUP, xpos, ypos, 1, 0);
        }


        #region Tiger Log
        private static MLogger Logger = MLogger.GetLogger(typeof(Toolbar));
        #endregion

        #region Channel property
        protected ReflectorForCSharp mobjReflector = new ReflectorForCSharp();
        #endregion

        // Do not call Base class methods or properties in the constructor.
        // The services are not initialized.
        public Toolbar()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        #region IRecord override Methods
        #region Wizard generated sample code (commented)
        /*		/// <summary>
		/// To change Window messages filter, implement this method.
		/// The default implementation is to get only the 
        /// Control's window messages.
		/// </summary>
		public override WND_MsgFilter GetWndMessageFilter()
		{
			return WND_MsgFilter.WND_MSGS;
		}
*/
        /// <summary>
        /// To catch window messages, implement this method.
        /// This method is called only if the CustomServer is running
        /// in the QuickTest context.
        /// </summary>
        public override RecordStatus OnMessage(ref Message tMsg)
        {
            Logger.logBegin("OnMessage");
            // TODO:  Add OnMessage implementation.
            return RecordStatus.RECORD_HANDLED;
        }

        #endregion
        /// <summary>
        /// To extend the Record process, add Events handlers
        /// to listen to the custom control's events.
        /// </summary>
        public override void InitEventListener()
        {
            #region Wizard generated sample code (commented)
            /*			// You can add as many handlers as you need.
			// For example, to add an OnMouseDown handler, 
            // first create the Delegate:
			Delegate  e = new System.Windows.Forms.MouseEventHandler(this.OnMouseDown);

            // Then, add the event handler as the first handler of the event.
			// The first argument is the name of the event for which to listen.
			// This must be an event that the control supports. You can
			// use the .NET Spy to obtain the list of events supported by the control.
			// The second argument is the event handler delegate. 

			AddHandler("MouseDown", e);
*/
            #endregion
#if _DEBUG
            MessageBox.Show("load");
#endif
            Logger.logBegin("InitEventListener");
            Delegate eClick = new System.EventHandler(this.OnItemClicked);
            base.AddHandler("Click", eClick);
            //Delegate eMouse = new System.Windows.Forms.MouseEventHandler(this.OnMouseDown) ;
            // base.AddHandler("MouseDown", eMouse);
        }



        /// <summary>
        /// Called by QuickTest to release the handlers
        /// added in the InitEventListener method.
        /// Only handlers added using QuickTest methods are released 
        /// by the QuickTest infrastructure. If you use standard C# syntax,
        /// you must release the handlers in your code at the end 
        /// of the Record process.
        /// </summary>
        public override void ReleaseEventListener()
        {
        }

        #endregion


        #region Record events handlers
        #region Wizard generated sample code (commented)
        /*		public void OnMouseDown(object sender, System.Windows.Forms.MouseEventArgs  e)
		{
			// This example shows how to create a script line in QuickTest
			// when a MouseDown event is encountered during recording.
			if(e.Button == System.Windows.Forms.MouseButtons.Left)
			{			
				RecordFunction( "CustomMouseDown", RecordingMode.RECORD_SEND_LINE, e.X, e.Y);
			}
		}
*/
        #endregion

        public void OnMouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            Logger.logBegin("OnMouseDown");
            base.RecordFunction("click", RecordingMode.RECORD_SEND_LINE, string.Format("clicked x:{0}", e.X));
            Logger.Info("OnMouseDown", string.Format("get positon :x:{0} y:{1}", e.X, e.Y));
            UltraToolbarsDockArea objToolbarArea = (UltraToolbarsDockArea)SourceControl;
            IList lstDockRows = mobjReflector.GetPrivateProperty<IList>(objToolbarArea, "DockRows");

            //Control objControl = objToolbarArea.GetChildAtPoint(new Point(e.X, e.Y));
            //Logger.Info("OnMouseDown", string.Format("Get instance type:{0}", (objControl == null) ? "NONE" : objControl.GetType().ToString()));


            //objControl = objToolbarArea.Parent.GetChildAtPoint(new Point(e.X, e.Y));
            Logger.logEnd("OnMouseDown");
        }

        public void OnItemClicked(object sender, System.EventArgs e)
        {

            Logger.logBegin("OnItemClicked");
            /** get postion of the current mouse **/
            Logger.Info("OnItemClicked", "begin to get MousePostion");
            Point pt = System.Windows.Forms.Control.MousePosition;
            Logger.Info("OnItemClicked", string.Format("get positon :x:{0} y:{1}", pt.X, pt.Y));

            UltraToolbarsDockArea objToolbarArea = (UltraToolbarsDockArea)SourceControl;

            UltraToolbarsManager objToolMgr = mobjReflector.GetMember<UltraToolbarsManager>(objToolbarArea, "ToolbarsManager");
            if (objToolMgr == default(UltraToolbarsManager))
            {
                Logger.Info("OnItemClicked", "Can't get :" + "ToolbarsManager");
                return;
            }

            ToolBase objActive = objToolMgr.ActiveTool;
            if (objActive == null)
            {
                Logger.Info("OnItemClicked", "Active Tool is null");
                return;
            }
            try
            {
                if (objActive is Infragistics.Win.UltraWinToolbars.PopupMenuTool)
                {
                    Logger.Info("OnItemClicked", "Clicked PopupMenuTool");
                    Infragistics.Win.UltraWinToolbars.PopupMenuTool objPopMenu = (Infragistics.Win.UltraWinToolbars.PopupMenuTool)objActive;
                    base.RecordFunction(MarsTigerServerConst.CNST_TOOLBAR_CLICK, RecordingMode.RECORD_SEND_LINE, string.Format("Menu:{0}", objPopMenu.Key));
                    Rectangle rctgl = objActive.Bounds;
                    Point ptx = new Point((rctgl.X) + rctgl.Width / 2, (rctgl.Y) + rctgl.Height / 2);
                    Point ptObj = objActive.UIElement.Control.PointToClient(ptx);
                    Logger.Info("OnItemClicked", string.Format("Menu:{4} Oraginal:{0},{1}, Client:{2},{3}", ptx.X, ptx.Y, ptObj.X, ptObj.Y, objActive.Key));
                    return;
                }

                if (objActive is Infragistics.Win.UltraWinToolbars.ButtonTool)
                {
                    Logger.Info("OnItemClicked", "Clicked ButtonTool");
                    Infragistics.Win.UltraWinToolbars.ButtonTool objButton = (Infragistics.Win.UltraWinToolbars.ButtonTool)objActive;
                    base.RecordFunction(MarsTigerServerConst.CNST_TOOLBAR_CLICK, RecordingMode.RECORD_SEND_LINE, string.Format("Button:{0}", objButton.Key));

                    Rectangle rctgl = objActive.Bounds;
                    Point ptx = new Point((rctgl.X) + rctgl.Width / 2, (rctgl.Y) + rctgl.Height / 2);
                    Point ptObj = objActive.UIElement.Control.PointToClient(ptx);
                    Logger.Info("OnItemClicked", string.Format("Oraginal:{0},{1}, Client:{2},{3}", ptx.X, ptx.Y, ptObj.X, ptObj.Y));
                    //base.MouseClick(ptObj.X,ptObj.Y, MOUSE_BUTTON.LEFT_MOUSE_BUTTON);
                    /*
                    CallEventForToolButton(objToolMgr,objActive, "ToolClick");

                    MethodInfo objMethod= mobjReflector.GetMethod(objToolMgr, "FireEvent");
                    //MethodInfo objMethod2 = mobjReflector.GetMethod(objActive, "OnToolClick") ;
                    //objMethod2.Invoke(objActive, new object[]{new ToolClickEventArgs(objActive,null)}) ;
                    //objToolMgr.Toolbars
                    if (objMethod == null)
                    {
                        Logger.Info("OnItemClicked", "no such Event");
                    }
                    else
                    {
                        objMethod.Invoke(objToolMgr, new object[] { ToolbarEventIds.ToolClick, new ToolClickEventArgs(objActive,null) });
                    }
                    //this.ToolbarsManager.FireEvent(ToolbarEventIds.ToolClick, new ToolClickEventArgs(this, null));
                    */
                    return;
                }

                Logger.Info("OnItemClicked", "unsupported type:" + objActive.GetType().ToString());
                return;
            }
            finally
            {
                Logger.logEnd("OnItemClicked");
            }
            /*
            UIElement objUIx = objToolMgr.UIElementFromPoint(pt);
            //this.ToolbarsManager.ToolFromPoint(new Point(Control.MousePosition.X - 1, Control.MousePosition.Y - 1));
            UltraToolbarsDockAreaUIElement objUI = mobjReflector.GetPrivateProperty<UltraToolbarsDockAreaUIElement>(objToolbarArea, "mainElement");
            object objPoint = objUI.ElementFromPoint(pt);
            Logger.Info("OnItemClicked", string.Format("Get ElementFromPoint type:{0}", (objPoint == null) ? "NONE" : objPoint.GetType().ToString()));
            Point ptClient = objToolbarArea.PointToClient(pt);
            objPoint = objUI.ElementFromPoint(pt);
            Logger.Info("OnItemClicked", string.Format("Get ElementFromPoint type:{0}", (objPoint == null) ? "NONE" : objPoint.GetType().ToString()));


            //Logger.Info("OnItemClicked", string.Format("Get instance type:{0}", (objInstance == null) ? "NONE" : objInstance.GetType().ToString()));
            //Control objInstance1 = objToolbarArea.MainUIElement.ElementFromPoint(ptClient, true);

            base.RecordFunction("Press", RecordingMode.RECORD_SEND_LINE, strButtonName);
             */

        }

        #endregion


        #region Replay interface implementation
        #region Wizard generated sample code (commented)
        /*		public void CustomMouseDown(int X, int Y)
		{
			MouseClick(X, Y, MOUSE_BUTTON.LEFT_MOUSE_BUTTON);
		}
*/
        #endregion
        public void Press(string strItemToPress)
        {
            Logger.logBegin("Pressx");
            PrepareForReplay();

            UltraToolbarsDockArea objToolbarArea = (UltraToolbarsDockArea)SourceControl;
            UltraToolbarsManager objToolMgr = mobjReflector.GetMember<UltraToolbarsManager>(objToolbarArea, "ToolbarsManager");
            base.MouseClick(22, 36, MOUSE_BUTTON.LEFT_MOUSE_BUTTON);

            ReplayReportStep("Press", EventStatus.EVENTSTATUS_GENERAL, strItemToPress);
            //base.MouseDown(22, 36, MOUSE_BUTTON.LEFT_MOUSE_BUTTON);
            //MessageBox.Show("click client");
            //base.MouseDown(22, 58, MOUSE_BUTTON.LEFT_MOUSE_BUTTON);
            /*
            for (int i = 0; i < (objToolMgr == null ? -1 : objToolMgr.Toolbars.Count); i++)
            {
                ToolsCollection lstTools = objToolMgr.Toolbars[i].Tools;
                for (int j = 0; j < (lstTools == null ? -1 : lstTools.Count); j++)
                {
                    //if (string.Compare lstTools[j].Key
                }
            }
             * */

        }

        public void CallEventForToolButton(UltraToolbarsManager objToolMgr, ToolBase objTool, string strEventName)
        {
            Logger.logBegin("CallEventForToolButton");
            if (objToolMgr == null) return;
            var objE = objToolMgr.GetType().GetEvent(strEventName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
            Type objDlgt = objE.EventHandlerType;

            BindingFlags b = BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static;
            MethodInfo objMethodMgr = objToolMgr.GetType().GetMethod(strEventName, b);
            MethodInfo objMethodBar = objTool.GetType().GetMethod(strEventName, b);

            Delegate objDMgr = Delegate.CreateDelegate(objDlgt, objMethodMgr);
            Delegate objDBar = Delegate.CreateDelegate(objDlgt, objMethodBar);


            var objED = (MulticastDelegate)objToolMgr.GetType().GetField(strEventName, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public).GetValue(objToolMgr);
            if (objED != null)
            {
                foreach (var handler in objED.GetInvocationList())
                {
                    handler.Method.Invoke(handler.Target, new object[] { objToolMgr, new ToolClickEventArgs(objTool, null) });
                }
            }
            Logger.logEnd("CallEventForToolButton");
        }

        public void ClickToolbarTool(params object[] arrStrKeys)
        {
            Logger.logBegin(MarsTigerServerConst.CNST_TOOLBAR_CLICK);
            try
            {
                bool isWrongParameters = false;
                if (arrStrKeys == null)
                {
                    isWrongParameters = true;
                }

                if ((!isWrongParameters) && (arrStrKeys.Length <= 0))
                {
                    isWrongParameters = true;
                }
                if (isWrongParameters)
                {
                    Logger.Error(MarsTigerServerConst.CNST_TOOLBAR_CLICK, "No paramters added");
                    return;
                }
                E_ClickToolbarMode eMode = E_ClickToolbarMode.e_Defautl_from_TA;//default mode, from Test advantage code
                if (MarsUFTAddins.IMars.tiger.ReflectorForCSharp.MarsTigerUtility.RegularExpressChecking(MarsTigerServerConst.CNST_TOOLBAR_CLICK_PARA_PREFIX_BUTTON_REG, arrStrKeys[0].ToString()))
                {
                    eMode = E_ClickToolbarMode.e_ClickButton;
                }
                if (MarsUFTAddins.IMars.tiger.ReflectorForCSharp.MarsTigerUtility.RegularExpressChecking(MarsTigerServerConst.CNST_TOOLBAR_CLICK_PARA_PREFIX_MENU_REG, arrStrKeys[0].ToString()))
                {
                    eMode = E_ClickToolbarMode.e_ClickMenu;
                }
                string strKey = arrStrKeys[0].ToString();
                switch (eMode)
                {
                    case E_ClickToolbarMode.e_Defautl_from_TA:
                        strKey = arrStrKeys[arrStrKeys.Length - 1].ToString();
                        break;
                    case E_ClickToolbarMode.e_ClickButton:
                        strKey = strKey.Replace(MarsTigerServerConst.CNST_TOOLBAR_CLICK_PARA_PREFIX_BUTTON, "");
                        break;
                    case E_ClickToolbarMode.e_ClickMenu:
                        strKey = strKey.Replace(MarsTigerServerConst.CNST_TOOLBAR_CLICK_PARA_PREFIX_MENU, "");
                        break;
                }

                UltraToolbarsDockArea objToolbarArea = (UltraToolbarsDockArea)SourceControl;
                UltraToolbarsManager objToolMgr = mobjReflector.GetMember<UltraToolbarsManager>(objToolbarArea, MarsTigerServerConst.CNST_INFR_TOOLBAR_MANAGER);
                string strError = "";
                if (objToolMgr == null)
                {
                    strError = string.Format("No {0} find, wrong object!", MarsTigerServerConst.CNST_INFR_TOOLBAR_MANAGER);
                    Logger.Error("ClickToolbarTool", strError);
                    base.PrepareForReplay();
                    base.ReplayThrowError(strError);
                    return;
                }

                for (int i = 0; i < objToolMgr.Toolbars.Count; i++)
                {
                    UltraToolbar objToolBar = objToolMgr.Toolbars[i];
                    if (objToolBar == null) continue;

                    for (int j = 0; j < objToolBar.Tools.Count; j++)
                    {
                        ToolBase objToolItem = objToolBar.Tools[j];
                        if (objToolItem == null) continue;
                        if (string.Compare(objToolItem.Key, strKey, true) == 0)
                        {
                            Rectangle rect = objToolItem.Bounds;
                            Point ptOrignal = new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2);
                            Point ptClnt = objToolItem.UIElement.Control.PointToClient(ptOrignal);
                            if (eMode == E_ClickToolbarMode.e_ClickMenu)
                            {
                                //popup Top menu
                                //objToolBar as PopupMen
                                if (objToolItem is Infragistics.Win.UltraWinToolbars.PopupMenuTool)
                                {
                                    ((Infragistics.Win.UltraWinToolbars.PopupMenuTool)objToolItem).DropDown();
                                    Thread.Sleep(50);
                                    //active MenuItem
                                    Infragistics.Win.UltraWinToolbars.PopupMenuTool objMenu = (Infragistics.Win.UltraWinToolbars.PopupMenuTool)objToolItem;
                                    //objToolItem.UIElement.Control.
                                    Logger.Info("ClickToolbarTool", string.Format("ClickToolbar Menu mode, Position:{0}, {1}", ptClnt.X, ptClnt.Y));
                                    base.MouseClick(ptClnt.X, ptClnt.Y, MOUSE_BUTTON.LEFT_MOUSE_BUTTON);
                                }
                                else
                                {
                                    //Error 
                                    strError = string.Format("Not a sub branch of a menu {0}", strKey);
                                    Logger.Error("ClickToolbarTool", strError);
                                    base.ReplayThrowError(strError);
                                    return;
                                }
                            }
                            else
                            {
                                LeftMouseClick(ptOrignal.X, ptOrignal.Y);
                                //base.MouseClick(ptClnt.X, ptClnt.Y, MOUSE_BUTTON.LEFT_MOUSE_BUTTON);
                            }
                            base.ReplayReportStep(MarsTigerServerConst.CNST_TOOLBAR_CLICK, EventStatus.EVENTSTATUS_PASS, arrStrKeys);
                            return;
                        }
                    }
                }
                strError = string.Format("Can't find Item [{0}] from the tool bar", arrStrKeys[0]);
                base.ReplayThrowError(strError);
            }
            catch (Exception e)
            {
                string strError = string.Format("Exceptions when call {0} with Excptions:{1}", MarsTigerServerConst.CNST_TOOLBAR_CLICK, e.Message);
                Logger.Error("ClickToolbarTool", strError, e);
                base.ReplayReportStep("ClickToolbarTool", EventStatus.EVENTSTATUS_FAIL, arrStrKeys);
            }
            finally
            {
                Logger.logEnd(MarsTigerServerConst.CNST_TOOLBAR_CLICK);
            }
        }

        public void ClickMenuItem(params object[] strMenuInfo)
        {
            Logger.logBegin(string.Format("ClickMenuItem {0}", "ClickMenuItem"));
            /**
             * ---------------------------------- IMPORTANT!!!! --------------------------------------------------------------------------------------
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
             * */
            E_TIGERMENU_CLICK eMenuMode = E_TIGERMENU_CLICK.E_OLD;
            string strError = "";
            if (strMenuInfo.Length == 0)
            {
                strError = "No available Menu Item passed.";
                base.ReplayThrowError(strError);
                return;
            }
            if (strMenuInfo.Length >= 1)
            {
                eMenuMode = E_TIGERMENU_CLICK.E_STARTWITH_MENU;
            }
            else
            {
                eMenuMode = mGetClickMenuMode(strMenuInfo[0].ToString());
            }
            switch (eMenuMode)
            {
                case E_TIGERMENU_CLICK.E_OLD:
                    mProcessMenuOldMode(strMenuInfo);
                    break;
                case E_TIGERMENU_CLICK.E_STARTWITH_MENU:
                    mProcessMenuWithStartWithMode(strMenuInfo);
                    break;
                /*
            default:
                PrepareForReplay();
                base.ReplayThrowError("Unsupported clickMenuMode: "+strMenuInfo);
                 */
            }

            Logger.logEnd("ClickMenuItem");
        }

        private void mProcessMenuWithStartWithMode(params object[] strMenuInfo)
        {
            Logger.logBegin(string.Format("mProcessMenuWithStartWithMode {0}", strMenuInfo));
            PrepareForReplay();

            try
            {
                string strIdx = strMenuInfo[0].ToString();
                strIdx = strIdx.Replace(MarsTigerServerConst.CNST_TOOLBAR_CLICK_PARA_PREFIX_MENU, "");

                UltraToolbarsDockArea objToolbarArea = (UltraToolbarsDockArea)SourceControl;
                UltraToolbarsManager objToolMgr = mobjReflector.GetMember<UltraToolbarsManager>(objToolbarArea, MarsTigerServerConst.CNST_INFR_TOOLBAR_MANAGER);

                int iCurrentLevel = 0;
                bool bContinue = iCurrentLevel < strMenuInfo.Length, bFindToolBar = false;
                string strCurrentKey = strIdx;
                ToolsCollection lstToolsCurrent = null;
                object objPopupTool = null;
                string strError = "";
                Point ptOff = Point.Empty;

                int iModeNumberId = -1;

                //re-write the code
                //"Desktop Menu","(MenuIdex|MenuName){1,}","Toolbars","Customize..."
                for (int i = 0; i < (strMenuInfo == null ? -1 : strMenuInfo.Length); i++)
                {
                    strCurrentKey = i == 0 ? strIdx : strMenuInfo[i].ToString();
                    bFindToolBar = false;
                    Thread.Sleep(1000);
                    if (i == 0)
                    {
                        #region ToolBar DocArea
                        for (int ii = 0; ii < objToolMgr.Toolbars.Count; ii++)
                        {
                            if (string.Compare(objToolMgr.Toolbars[ii].Key, strCurrentKey, true) == 0)
                            {
                                bFindToolBar = true;
                                UltraToolbar objToolBar = objToolMgr.Toolbars[ii];
                                lstToolsCurrent = objToolBar.Tools;
                                Thread.Sleep(100);
                                break;
                            }
                        }
                        #endregion

                    }
                    else
                    {

                        if (i == 1) // parent is tool bar
                        {
                            if (int.TryParse(strCurrentKey, out iModeNumberId))
                            {
                                //Number Mode                                
                                if (lstToolsCurrent[iModeNumberId] is Infragistics.Win.UltraWinToolbars.PopupMenuTool)
                                {
                                    bFindToolBar = true;
                                    objPopupTool = (Infragistics.Win.UltraWinToolbars.PopupMenuTool)lstToolsCurrent[iModeNumberId];
                                    Rectangle rect = ((Infragistics.Win.UltraWinToolbars.PopupMenuTool)objPopupTool).Bounds;
                                    Point ptClient = ((Infragistics.Win.UltraWinToolbars.PopupMenuTool)objPopupTool).UIElement.Control.PointToClient(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2));
                                    //base.MouseClick(ptClient.X, ptClient.Y, MOUSE_BUTTON.LEFT_MOUSE_BUTTON);
                                    LeftMouseClick(rect.X + rect.Width / 2, rect.Y + rect.Height / 2);
                                    Thread.Sleep(100);
                                    //break;
                                }
                                else
                                {
                                    strError = string.Format("Only PopupMenuTool on second level, but [{1}]'s type:[{0}] find.", lstToolsCurrent[iModeNumberId].GetType().ToString(), strCurrentKey);
                                    Logger.Error("mProcessMenuWithStartWithMode", strError);
                                    base.ReplayThrowError(strError);
                                    base.ReplayReportStep("ClickMenuItem", EventStatus.EVENTSTATUS_FAIL, strMenuInfo);
                                    return;
                                }                                
                            }
                            else
                            {
                                #region to find first key of menu
                                for (int ii = 0; ii < (lstToolsCurrent == null ? -1 : lstToolsCurrent.Count); ii++)
                                {
                                    if ((string.Compare(lstToolsCurrent[ii].Key, strCurrentKey, true) == 0) || (string.Compare(lstToolsCurrent[ii].CaptionResolved.Replace("&", ""), strCurrentKey, true) == 0)
                                        || (TigerMarsUtil.RegularTest(strCurrentKey, lstToolsCurrent[ii].Key))
                                        || (TigerMarsUtil.RegularTest(strCurrentKey, lstToolsCurrent[ii].CaptionResolved.Replace("&",""))))
                                    {
                                        if (lstToolsCurrent[ii] is Infragistics.Win.UltraWinToolbars.PopupMenuTool)
                                        {
                                            bFindToolBar = true;
                                            objPopupTool = (Infragistics.Win.UltraWinToolbars.PopupMenuTool)lstToolsCurrent[ii];
                                            Rectangle rect = ((Infragistics.Win.UltraWinToolbars.PopupMenuTool)objPopupTool).Bounds;
                                            Point ptClient = ((Infragistics.Win.UltraWinToolbars.PopupMenuTool)objPopupTool).UIElement.Control.PointToClient(new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2));
                                            //base.MouseClick(ptClient.X, ptClient.Y, MOUSE_BUTTON.LEFT_MOUSE_BUTTON);
                                            LeftMouseClick(rect.X + rect.Width / 2, rect.Y + rect.Height / 2);
                                            Thread.Sleep(100);
                                            break;
                                        }
                                        else
                                        {
                                            strError = string.Format("Only PopupMenuTool on second level, but [{1}]'s type:[{0}] find.", lstToolsCurrent[ii].GetType().ToString(), strCurrentKey);
                                            Logger.Error("mProcessMenuWithStartWithMode", strError);
                                            base.ReplayThrowError(strError);
                                            base.ReplayReportStep("ClickMenuItem", EventStatus.EVENTSTATUS_FAIL, strMenuInfo);
                                            return;
                                        }
                                    }
                                }
                                #endregion
                            }
                        }
                        else
                        {
                            // logic problem here
                            #region menu items
                            if ((objPopupTool is Infragistics.Win.UltraWinToolbars.PopupMenuTool) && (i != (strMenuInfo.Length - 1)))
                            {
#if _tigerDebug
                                Logger.Info("-------PopupMenuTool-------", "");
#endif
                                Infragistics.Win.UltraWinToolbars.PopupMenuTool objPopMenu = (Infragistics.Win.UltraWinToolbars.PopupMenuTool)objPopupTool;
                                if (int.TryParse(strCurrentKey, out iModeNumberId))
                                {
                                    Thread.Sleep(100);
                                    if ((iModeNumberId < 0) || (iModeNumberId >= (objPopMenu.Tools == null ? -int.MaxValue : objPopMenu.Tools.Count)))
                                    {
                                        strError = string.Format("No such Number of MenuItem Idx found:[{0}]", strCurrentKey);
                                        base.ReplayReportStep("ClickMenuItem/SelectMenuItem",EventStatus.EVENTSTATUS_FAIL, new object[]{strError});
                                        return;
                                    }
                                    ToolBase objToolItem = objPopMenu.Tools[iModeNumberId];
                                    if (objToolItem is Infragistics.Win.UltraWinToolbars.PopupMenuTool)
                                    {
                                        objPopupTool = objToolItem;

                                        /** click it **/
                                        Rectangle rect = ((Infragistics.Win.UltraWinToolbars.PopupMenuTool)objToolItem).Bounds;
                                        LeftMouseClick(rect.X + rect.Width / 2, rect.Y + rect.Height / 2);
                                        bFindToolBar = true;
                                        Thread.Sleep(100);
                                    }
                                    else
                                    {
                                        strError = string.Format("Only PopupMenuTool of this level is supported, but current type is [{0}]",objToolItem.GetType().ToString());
                                        base.ReplayReportStep("ClickMenuItem/SelectMenuItem", EventStatus.EVENTSTATUS_FAIL, new object[]{strError});
                                        return;
                                    }
                                }
                                else
                                {
                                    for (int j = 0; j < (objPopMenu.Tools == null ? -1 : objPopMenu.Tools.Count); j++)
                                    {
                                        #region to find tools
                                        ToolBase objToolItem = objPopMenu.Tools[j];
                                        bFindToolBar = false;
                                        if (!((string.Compare(objToolItem.Key, strCurrentKey, true) == 0) || (string.Compare(objToolItem.CaptionResolved.Replace("&", ""), strCurrentKey, true) == 0)))
                                        {
                                            continue;
                                        }
                                        // find key
                                        if (objToolItem is Infragistics.Win.UltraWinToolbars.PopupMenuTool)
                                        {
                                            if ((string.Compare(objToolItem.Key, strCurrentKey, true) == 0) || (string.Compare(objToolItem.CaptionResolved.Replace("&", ""), strCurrentKey, true) == 0))
                                            {
                                                objPopupTool = objToolItem;

                                                /** click it **/
                                                Rectangle rect = ((Infragistics.Win.UltraWinToolbars.PopupMenuTool)objToolItem).Bounds;                                          
                                                LeftMouseClick(rect.X + rect.Width / 2, rect.Y + rect.Height / 2);
                                                bFindToolBar = true;
                                                Thread.Sleep(100);
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            // then the leaf node is found
#if _tigerDebug
                                            // 可以是叶节点，也可能是错误的设置。
                                            // 如果
#endif
                                            if (i == strMenuInfo.Length - 1)
                                            {
                                                /** find parent **/

                                                bool isMenuParent = this.mobjReflector.GetMember<bool>(objPopupTool, "OwnerIsMenu");
                                                if (!isMenuParent)
                                                {
                                                    strError = string.Format("Non-popupmenu's parent should be Menu,but OwnerIsMenu is false,type:[{0}], value:[{1}]", objPopupTool.GetType().ToString(), isMenuParent);
                                                    base.ReplayReportStep("ClickMenuItem", EventStatus.EVENTSTATUS_FAIL, new object[] { strError });
                                                    return;
                                                }
                                                /*** get parent infor ***/
                                                StateButtonTool objButton = (StateButtonTool)objPopupTool;
                                            }
                                            else
                                            {
                                                // Error
                                                strError = string.Format("The last item could be non-popupmenu, but for lever:[{0}/{1}], it is type:[{2}]", i + 1, strMenuInfo.Length, objPopupTool.GetType().ToString());
                                                base.ReplayReportStep("ClickMenuItem", EventStatus.EVENTSTATUS_FAIL, new object[] { strError });
                                                return;
                                            }
                                        }
                                        #endregion
                                    }
                                }
                                if (!bFindToolBar)
                                {
                                    // then nothing matched for the current key.
                                    strError = string.Format("No such key found:[{0}]", strCurrentKey);
                                    base.ReplayReportStep("ClickMenuItem", EventStatus.EVENTSTATUS_FAIL, new object[] { strError });
                                    return;
                                }
                            }
                            #endregion menu items
                            Thread.Sleep(1000) ;
                            if (i == strMenuInfo.Length - 1)
                            {

                                #region last parameter is a number
                                bool isNumber = false;
                                int iKeyDownCnt = -1;
                                try
                                {
                                    iKeyDownCnt = int.Parse(strMenuInfo[i].ToString());
                                    isNumber = true;
                                }
                                catch (Exception)
                                {
                                    isNumber = false;
                                }
                                

                                #endregion //last parameter is a number

                                #region last node
                                SubObjectBase objMenuAgent = (SubObjectBase)mobjReflector.GetMember<SubObjectBase>(objPopupTool, "MenuAgent");
                                if (objMenuAgent != null)
                                {
                                    object objMenuItems = (object)mobjReflector.GetMember<object>(objMenuAgent, "MenuItems");
                                    if (objMenuItems != null)
                                    {
                                        int iCnt = mobjReflector.GetMember<int>(objMenuItems, "Count");
                                        SubObjectsCollectionBase lst = (SubObjectsCollectionBase)objMenuItems;
                                        if (isNumber)
                                        {
                                            if (iKeyDownCnt > lst.Count)
                                            {
                                                strError = string.Format("No such many items [{0}] ,Only [{1}] Items found", iKeyDownCnt, lst.Count);
                                                Logger.Error("mProcessMenuWithStartWithMode", strError);
                                                base.ReplayThrowError(strError);
                                                base.ReplayReportStep("ClickMenuItem", EventStatus.EVENTSTATUS_FAIL, strMenuInfo);
                                                return;
                                            }
                                            
                                            Rectangle rect = mobjReflector.GetMember<Rectangle>(lst.GetItem(iKeyDownCnt-1), "Bounds");
                                            //MessageBox.Show(string.Format("type:[{4}],x:{0} y:{1} width:{2} height:{3}", rect.X, rect.Y, rect.Width, rect.Height, lst.GetItem(iKeyDownCnt - 1).GetType().ToString()));
                                            LeftMouseClick(rect.X + rect.Width / 2, rect.Y + rect.Height / 2);
                                            bFindToolBar = true;
                                        }
                                        else{

                                            for (int ii = 0; ii < lst.Count; ii++)
                                            {
                                                object objMenuItem = lst.GetItem(ii);
                                                if (objMenuItem == null) continue;

                                                string strMenuItemTxt = mobjReflector.GetMember<string>(objMenuItem, "Text");
#if _tigerDebug
                                                Logger.Info("~~~~~~~", string.Format("to compare:[{0}], with:[{1}]", strMenuItemTxt, strCurrentKey));
#endif
                                                if (string.Compare(strMenuItemTxt.Replace("&", ""), strCurrentKey) == 0)
                                                {
                                                    int iWaitCnt = 0;
                                                    Rectangle rect = Rectangle.Empty;
                                                    rect.Width = 0;
                                                    while((iWaitCnt<100) &&(rect.Width==0))
                                                    {
                                                        Thread.Sleep(100);
                                                        rect = mobjReflector.GetMember<Rectangle>(objMenuItem, "Bounds");
                                                    }

                                                    LeftMouseClick(rect.X + rect.Width / 2, rect.Y + rect.Height / 2);
                                                    Thread.Sleep(150);
#if _tigerDebug
                                                    Logger.Info("~~~~~~11", string.Format("before change:[{0}]", objPopupTool.GetType().ToString()));
#endif
                                                    objPopupTool = objMenuItem;
#if _tigerDebug
                                                    Logger.Info("~~~~~~11", string.Format("After change:[{0}]", objPopupTool.GetType().ToString()));
#endif
                                                    bFindToolBar = true;
                                                    break;
                                                }

                                            }
                                        }
                                    }

                                }
                                else
                                {
                                    strError = string.Format("Not a MenuAgent parent, [{0}], type:[{1}]", strCurrentKey, objPopupTool.GetType().ToString());
                                    Logger.Error("mProcessMenuWithStartWithMode", strError);
                                    base.ReplayThrowError(strError);
                                    base.ReplayReportStep("ClickMenuItem", EventStatus.EVENTSTATUS_FAIL, strMenuInfo);
                                    return;
                                }
                                #endregion last node
                            }
                            /*
                        else
                        {
                            // PopupMenuItem
                            if (!(objPopupTool is Infragistics.Win.UltraWinToolbars.PopupMenuTool))
                            {
                                Infragistics.Win.UltraWinToolbars.PopupMenuTool objMenuTool = (Infragistics.Win.UltraWinToolbars.PopupMenuTool)objPopupTool;
                                SubObjectBase objMenuAgent = (SubObjectBase)mobjReflector.GetMember<SubObjectBase>(objPopupTool, "MenuAgent");

                            }
                            else
                            {
                                strError = string.Format("Not a PopupMenuTool, [{0}], type:[{1}]", strCurrentKey, objPopupTool.GetType().ToString());
                                Logger.Error("mProcessMenuWithStartWithMode", strError);
                                base.ReplayReportStep("ClickMenuItem", EventStatus.EVENTSTATUS_FAIL, strMenuInfo);
                                return;
                            }
                        }
                             * */
                        }
                    }
                    if (!bFindToolBar)
                    {
                        strError = string.Format("Can't find ToolBar information with key :[{0}]", strCurrentKey);
                        Logger.Error("mProcessMenuWithStartWithMode", strError);
                        base.ReplayThrowError(strError);
                        base.ReplayReportStep("ClickMenuItem", EventStatus.EVENTSTATUS_FAIL, strMenuInfo);
                        return;
                    }
                }
                #region _oldCode
                /*
                while (bContinue)
                {
                    bFindToolBar = false;
                    if (iCurrentLevel == 0)
                    {
                        strCurrentKey = strIdx;
                    }
                    else
                    {
                        strCurrentKey = strMenuInfo[iCurrentLevel].ToString();
                    }
                    iCurrentLevel++;
                    bContinue = iCurrentLevel < strMenuInfo.Length;

                    if ((iCurrentLevel-1) == 0)
                    {
                        // get menu 
                        for (int i = 0; i < objToolMgr.Toolbars.Count; i++)
                        {
                            if (string.Compare(objToolMgr.Toolbars[i].Key, strCurrentKey, true) == 0)
                            {
                                bFindToolBar = true;
                                UltraToolbar objToolBar = objToolMgr.Toolbars[i];
                                lstToolsCurrent = objToolBar.Tools;
                                
                                break;
                            }
                        }
                        bFindToolBar = false;
                    }
                    else
                    {
                        if (lstToolsCurrent == null)
                        {
                            // something wrong here
                            strError = string.Format("Can't find keys info, Current Level:[{0}], Current Keys:[{1}]", iCurrentLevel - 1, strMenuInfo[iCurrentLevel - 1].ToString());
                            Logger.Error("mProcessMenuWithStartWithMode", strError);
                            base.ReplayThrowError(strError);
                            break;
                        }
                        else
                        {
                            bFindToolBar = false;
                            Point ptClient = Point.Empty;
                            for (int i = 0; i < lstToolsCurrent.Count; i++)
                            {
#if _tigerDebug
                                Logger.Info("++++++++++", string.Format("Key find:[{0}], Input Kye:[{1}]",lstToolsCurrent[i].Key, strCurrentKey));
                                Logger.Info("~~~~~~~~~~", string.Format("Key find:[{0}], Type:[{1}], is PopupMenutool?[{2}]", lstToolsCurrent[i].Key, lstToolsCurrent[i].GetType().ToString(), (lstToolsCurrent[i] is Infragistics.Win.UltraWinToolbars.PopupMenuTool)));
#endif
                                if (string.Compare(lstToolsCurrent[i].Key, strCurrentKey, true) == 0)
                                {
                                    if (lstToolsCurrent[i] is Infragistics.Win.UltraWinToolbars.PopupMenuTool)
                                    {
                                        bFindToolBar = true;

                                        Infragistics.Win.UltraWinToolbars.PopupMenuTool objPopupTool = (Infragistics.Win.UltraWinToolbars.PopupMenuTool)lstToolsCurrent[i];
                                        SubObjectBase objMenuAgent = (SubObjectBase)mobjReflector.GetMember<SubObjectBase>(objPopupTool,"MenuAgent");
                                        //PopupMenuAgent
                                        
                                        //if (objPopupTool is Infragistics.Win.UltraWinToolbars.PopupMenuControlTrusted)
                                        //{
                                        //    Logger.Info("---", "PopupMenuControlTrusted");
                                        //}

                                        //objPopupTool

                                        //do click

                                        
                                        if (ptOff == Point.Empty)
                                        {
                                            ptClient = objPopupTool.UIElement.Control.PointToClient(new Point(objPopupTool.Bounds.X + objPopupTool.Bounds.Width / 2, objPopupTool.Bounds.Y + objPopupTool.Bounds.Height / 2));
                                            ptOff = new Point(objPopupTool.Bounds.X, objPopupTool.Bounds.Height);
                                        }
                                        else
                                        {
                                            ptClient = objPopupTool.UIElement.Control.PointToClient(new Point(objPopupTool.Bounds.X + objPopupTool.Bounds.Width / 2, objPopupTool.Bounds.Y + ptOff.Y + objPopupTool.Bounds.Height / 2));
                                        }
                                        //objPopupTool.DropDown();
#if _tigerDebug
                                        Logger.Info("===========", string.Format("trying to click -x:[{0}], y:[{1}], Original-x:[{2}],y:[{3}]", ptClient.X, ptClient.Y, objPopupTool.Bounds.X, objPopupTool.Bounds.Y));
#endif
                                        LeftMouseClick(objPopupTool.Bounds.X + objPopupTool.Bounds.Width / 2,objPopupTool.Bounds.Y + objPopupTool.Bounds.Height / 2

                                        //base.MouseClick(ptClient.X, ptClient.Y, MOUSE_BUTTON.LEFT_MOUSE_BUTTON);
                                        Thread.Sleep(350);

                                        if (objMenuAgent != null)
                                        {
                                            Logger.Info("^^^^^^^1", objMenuAgent.GetType().ToString());
                                            object objMenuItems = (object)mobjReflector.GetMember<object>(objMenuAgent, "MenuItems");
                                            if (objMenuItems != null)
                                            {
                                                Logger.Info("^^^^^^^2", objMenuItems.GetType().ToString());
                                                int iCnt = mobjReflector.GetMember<int>(objMenuItems, "Count");
                                                Logger.Info("^^^^^^^2", string.Format("get count:{0}", iCnt));
                                                SubObjectsCollectionBase lst = (SubObjectsCollectionBase)objMenuItems;
                                                for (int ii = 0; ii < lst.Count; ii++)
                                                {
                                                    object objMenuItem = lst.GetItem(ii);
                                                    if (objMenuItem == null) continue;
                                                    Logger.Info("^^^^^^^3", string.Format("type:[{0}]", objMenuItem.GetType().ToString()));
                                                    string strMenuItemTxt = mobjReflector.GetMember<string>(objMenuItem, "Text");
                                                    Logger.Info("^^^^^^^3-strMenuItemTxt", string.Format("Text:[{0}]", strMenuItemTxt));

                                                    if (string.Compare(strMenuItemTxt.Replace("&", ""), "Customize...") == 0)
                                                    {
                                                        //mobjReflector.CallMethod(objMenuItem, "Activate", null);
                                                        //Logger.Info("^^^^^^^3-Activate", string.Format("Text:[{0}]", "OnClick"));

                                                        Rectangle rect = mobjReflector.GetMember<Rectangle>(objMenuItem, "Bounds");
                                                        Logger.Info("^^^^^^^3-GetBounds", string.Format("Bounds:[{0}]", rect));
                                                        LeftMouseClick(rect.X + 1, rect.Y + 1);

                                                        //base.MouseClick(rect.X, rect.Y, MOUSE_BUTTON.LEFT_MOUSE_BUTTON);
                                                        
                                                    }

                                                    UIElement objItemUI = mobjReflector.GetMember<UIElement>(objMenuItem, "UIElement");

                                                    if (objItemUI == null) continue;
                                                    Logger.Info("^^^^^^^3-UIELement", string.Format("type:[{0}]", objItemUI.GetType().ToString()));
                                                    //Infragistics.Win.UltraWinToolbars.ToolMenuItem 

                                                }


                                                //IEnumerable lstMenu = (IEnumerable)objMenuItems;
                                                //IEnumerator objIETr = lstMenu.GetEnumerator();
                                                //while (objIETr.MoveNext())
                                                //{
                                                //    object objMenuItem = lstMenu.GetEnumerator().Current;
                                                //    if (objMenuItem != null)
                                                //    {
                                                //        UIElement objItemUI = mobjReflector.GetMember<UIElement>(objMenuItem, "");
                                                //    }
                                                //}


                                            }
                                        }

                                        lstToolsCurrent = objPopupTool.Tools;
                                        break;
                                    }
                                    else
                                    {
                                        if (lstToolsCurrent[i] is Infragistics.Win.UltraWinToolbars.ButtonTool)
                                        {
                                            // the last one
                                            Infragistics.Win.UltraWinToolbars.ButtonTool objButton = (Infragistics.Win.UltraWinToolbars.ButtonTool)lstToolsCurrent[i];
                                            ptClient = objButton.UIElement.Control.PointToClient(new Point(objButton.Bounds.X + objButton.Bounds.Width / 2, objButton.Bounds.Y + objButton.Bounds.Height / 2));
#if _tigerDebug
                                            Logger.Info("=====ButtonTool======", string.Format("trying to click -x:[{0}], y:[{1}], Original-x:[{2}],y:[{3}]", ptClient.X, ptClient.Y, objButton.Bounds.X, objButton.Bounds.Y));
#endif
                                            bFindToolBar = true;
                                            //base.MouseClick(ptClient.X, ptClient.Y, MOUSE_BUTTON.LEFT_MOUSE_BUTTON);
                                            bFindToolBar = true;
                                        }
                                        else
                                        {
                                            bFindToolBar = false;
                                        }
                                    }
                                }
                                
                            }
                            if (!bFindToolBar)
                            {
                                strError = string.Format("Can't find keys info, Current Level:[{0}], Current Keys:[{1}]", iCurrentLevel - 1, strMenuInfo[iCurrentLevel - 1].ToString());
                                Logger.Error("mProcessMenuWithStartWithMode", strError);
                                base.ReplayThrowError(strError);
                                bContinue = false;
                                break;
                            }
                        }
                    }

                }
                 */
                #endregion
                base.ReplayReportStep("ClickMenuItem", EventStatus.EVENTSTATUS_FAIL, strMenuInfo);


                /** get list **/
                /*
                bool bFind = true;
                List<object> lstStack = new List<object>();
                for (int i = 0; i < (objToolMgr == null ? -1 : (objToolMgr.Toolbars == null ? -1 : objToolMgr.Toolbars.Count)); i++)
                {
                    UltraToolbar objToolBar = objToolMgr.Toolbars[i];

                    
                    //bool bFind = mGetMenuInfoStack(objToolBar, strMenuInfo,ref lstStack);
                    if (!bFind)
                    {
                        lstStack.Clear();
                    }
                    else
                    {

                    }
                }
                 * */
            }
            catch (Exception e)
            {
                string strError = string.Format("Exceptions when replay ClickMenuItem:{0}, Exceptions:{1}", strMenuInfo, e.Message);
                Logger.Error("mProcessMenuWithStartWithMode", strError, e);
                base.ReplayThrowError(strError);
            }
            finally
            {
                Logger.logEnd("mProcessMenuWithStartWithMode");
            }

        }

        

        private bool mGetMenuInfoStack(UltraToolbar objToolBar, string strMenuInfo, ref List<object> lstStack)
        {
            Logger.logBegin(string.Format("mGetMenuInfoStack {0}", MarsUFTAddins.IMars.tiger.ReflectorForCSharp.MarsTigerUtility.CombinePara("MenuInfo", strMenuInfo)));
            if (objToolBar == null) return false;
            if (lstStack == null) lstStack = new List<object>();

            for (int i = 0; i < (objToolBar == null ? -1 : objToolBar.Tools.Count); i++)
            {
                ToolBase objToolItem = objToolBar.Tools[i];
                if (!(objToolItem is PopupMenuTool)) continue;

                PopupMenuTool objMenu = (PopupMenuTool)objToolItem;

                ToolBase objMenuItem = objMenu.Tools[0];
                //objMenuItem.OwningToolbar

                if (string.Compare(objToolItem.Key, strMenuInfo, true) == 0)
                {
                    //find
                    lstStack.Add(objToolItem);
                    return true;
                }
                //if (objToolItem.)
            }

            Logger.logEnd("mGetMenuInfoStack");
            return false;
        }

        private void mProcessMenuOldMode(params object[] strMenuInfo)
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

        public void MouseDown(int x, int y)
        {
            Logger.logBegin("MouseDonw");
        }

        private bool isValidateOfComp = false;

        public bool Wait4Validate()
        {
            Logger.logBegin("Wait4Validate");
            UltraToolbarsDockArea objToolbarArea = (UltraToolbarsDockArea)SourceControl;
            objToolbarArea.Validated += new EventHandler(objToolbarArea_Validated);
            int iTotalWait = 0;
            try
            {
                while (!isValidateOfComp)
                {
                    Thread.Sleep(100);
                    iTotalWait += 100;
                    if (iTotalWait >= (3 * 60 * 1000))
                    {
                        Logger.Error("Wait4Validate", "Waiting lasts over 3 minutes");
                        break;
                    }
                }
            }
            catch (Exception e)
            {
                Logger.Error("Wait4Validate",string.Format("Exceptions when waiting...., [{0}]", e.Message),e);
            }
            finally
            {
                objToolbarArea.Validated -= objToolbarArea_Validated;
            }
            Logger.logEnd("Wait4Validate");
            return isValidateOfComp ;
        }

        void objToolbarArea_Validated(object sender, EventArgs e)
        {
            isValidateOfComp = true;
        }

        public void WriteCursorStatus()
        {
            Logger.logBegin("WriteCursorStatus");
            UltraToolbarsDockArea objToolbarArea = (UltraToolbarsDockArea)SourceControl;
            
            for (int i = 0; i < 100; i++)
            {
                Console.WriteLine(objToolbarArea.Cursor.ToString());
                Thread.Sleep(50);
            }
            Logger.logEnd("WriteCursorStatus");
        }

        #endregion
    }
}
