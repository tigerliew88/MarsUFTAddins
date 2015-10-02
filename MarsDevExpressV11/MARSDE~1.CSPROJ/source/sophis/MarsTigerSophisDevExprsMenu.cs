using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Mercury.QTP.CustomServer;
using Route2NSEx.src.Marquis.systemUtil;
using DevExpress.XtraBars;
using MarsTestFrame.systemUtil;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;

namespace MarsUFTAddins.source.sophis
{
    
    [ReplayInterface]
    public interface MenuReplayC
    {
        #region Wizard generated sample code (commented)
        //		void  CustomMouseDown(int X, int Y);
        //void Press(string strItemToPress);
        void MouseDown(int x, int y);
        bool ClickMenuItem(object objMenuInfo);
        #endregion
    }

    public class MarsTigerSophisDevExprsMenu : CustomServerBase, MenuReplayC
    {
        private MLogger Logger = MLogger.GetLogger(typeof(MarsTigerSophisDevExprsMenu));

        public void MouseDown(int x, int y)
        {
            Logger.logBegin("MouseDown");
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

		/// <summary>
		/// To catch window messages, implement this method.
		/// This method is called only if the CustomServer is running
		/// in the QuickTest context.
		/// </summary>
		public override RecordStatus OnMessage(ref Message tMsg)
		{
			// TODO:  Add OnMessage implementation.
			return RecordStatus.RECORD_HANDLED;
		}
*/
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
        #endregion

        #region MarsFuncitons
        public bool ClickMenuItem(object objMenuInfo)
        {
            Logger.logBegin("ClickMenuItem");
            //MessageBox.Show("come");
            string strMenuInfo = objMenuInfo.ToString();
            Logger.Info("ClickMenuItem",string.Format("parameters:[{0}]", strMenuInfo));
            try
            {
                string[] arrMenuItems = strMenuInfo.Split(new string[]{";"}, StringSplitOptions.RemoveEmptyEntries);
                if (!(base.SourceControl is DevExpress.XtraBars.Controls.PopupMenuBarControl))
                {
                    Logger.Error("ClickMenuItem", string.Format("Control is not DevExpress.XtraBars.Controls.PopupMenuBarControl, but [{0}]", base.SourceControl.GetType().ToString()));
                    return false;
                }
                DevExpress.XtraBars.Controls.PopupMenuBarControl objMenu = (DevExpress.XtraBars.Controls.PopupMenuBarControl)base.SourceControl;
                //objMenu.Menu.ItemLinks[0].
                List<BarItemLink> lstMenuItem=new List<BarItemLink>() ;
                if (TraverseItems(objMenu.Menu.ItemLinks, arrMenuItems, 0, lstMenuItem))
                {
                    if (lstMenuItem.Count > 0)
                    {
                        /** click all items by order **/
                        for (int i = 0; i < lstMenuItem.Count; i++)
                        {
                            Rectangle rect = lstMenuItem[i].Bounds;
                            /** convert to screen Position **/
                            Point ptScreen = objMenu.PointToScreen(new Point(rect.X, rect.Y));
                            TigerMarsUtil.LeftMouseClick(ptScreen.X + rect.Width / 2, ptScreen.Y + rect.Height / 2);
                            Thread.Sleep(100);
                        }
                        return true;
                    }
                    Logger.Error("ClickMenuItem", "TraverseItems return true, but no items exists in Chain. ");
                    return false;
                }
                else
                {
                    Logger.Error("ClickMenuItem", "no menu Item matches. ");
                    return false;
                }

                
            }
            finally
            {
                Logger.logEnd("ClickMenuItem");
            }
        }

        private bool TraverseItems(BarItemLinkCollection col, string[] arrItems, int iCurrentLevel, List<BarItemLink> lstMenuChain )
        {
            if (iCurrentLevel >= arrItems.Length)
            {
                Logger.Error("TraverseItems", "Can't get the right menuItem when Navigating reaches the last item.");
                return false;
            }
            foreach (BarItemLink link in col)
            {
                //SomeAction(link);
                if (!(TigerMarsUtil.RegularTest(arrItems[iCurrentLevel], link.Caption))) continue;
                /** keep the current menuItem **/
                lstMenuChain.Add(link);
                if (link.Item is BarSubItem)
                {
                    
                    if (TraverseItems((link.Item as BarSubItem).ItemLinks, arrItems, iCurrentLevel + 1, lstMenuChain))
                    {
                        return true;
                    }
                    else
                    {
                        /** remove the last one **/
                        if (lstMenuChain.Count > 0)
                        {
                            Logger.Info("TraverseItems", string.Format("Its Subitems [{0}] are not match to [{1}], level:[{2}]. Remove last node of the chain", link.Caption, arrItems[iCurrentLevel], iCurrentLevel));
                            lstMenuChain.RemoveAt(lstMenuChain.Count - 1);
                            return false;
                        }
                    }
                }
                else
                {
                    return true;
                }
            }
            return false;
        }
        #endregion //MarsFuncitons
    }
}
