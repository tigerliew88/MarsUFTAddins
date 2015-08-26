using System;
using System.Collections.Generic;
#if tiger_dotNet4
using System.Linq;
#endif
using System.Text;
using Mercury.QTP.CustomServer;
using Route2NSEx.src.Marquis.systemUtil;
using Infragistics.Win.UltraWinTree;
using System.Drawing;
using System.Threading;

namespace MarsUFTAddins.IMars.tiger.infragistics.v12
{

    [ReplayInterface]
    public interface IMarsTigerTreeviewReplay
    {
        bool ClickAtSpecialNode(params object[] arrObj);
    }

    sealed class MouseClickInfo
    {
        public int miMouseId;
        public string NodePath;

        internal string[] ConertPathToArr()
        {
            if (NodePath == null) return null;
            string strTrim = NodePath.Trim();
            while (strTrim.StartsWith("\\"))
            {
                strTrim.Remove(0, 1);
            }
            return strTrim.Split('\\');
        }
    }

    public class MarsTigerTreeViewRep : MarsTigerServerBase, IMarsTigerTreeviewReplay
    {
        private static MLogger Logger = MLogger.GetLogger(typeof(MarsTigerTreeViewRep));

        private MouseClickInfo mobjMouseInfo = new MouseClickInfo();



        protected override void AddEvent()
        {
            Logger.logBegin("AddEvent");
#if _tigerDebugMouseDown
            //Delegate eMouseDown = new System.Windows.Forms.MouseEventHandler(this.tigerMouseDown);
            //base.AddHandler(MarsTigerServerConst.CNST_EVNT_MOUSEDOWN, eMouseDown);
#endif
            Logger.logEnd("AddEvent");
        }

        public bool ClickAtSpecialNode(params object[] arrParams)
        {
            Logger.logBegin("ClickAtSpecialNode");
            /** 
             * Parameters:
             * 1, Left or right mouse button,0, left,1-right,2 left-double click
             * 2, Node text to be clicked at with path
             * steps:
             * 1,
             * **/
            try
            {
                base.PrepareForReplay();
                string strError = "";
                if (!ParametersFormatCheckForClickSpecialNode(arrParams))
                {
                    strError = "parameters format aren't right!";
                    Logger.Error("ClickAtSpecialNode", strError);
                    base.ReplayReportStep("ClickAtSpecialNode", EventStatus.EVENTSTATUS_FAIL, new object[] { arrParams, strError });
                    return false;
                }
                UltraTree objTree = (UltraTree)this.SourceControl;
                string[] arrNodesPath = this.mobjMouseInfo.ConertPathToArr();
                UltraTreeNode objTargetNode = FindSpecialNode(objTree.Nodes, arrNodesPath, 0);
                if (objTargetNode == null)
                {
                    strError = string.Format("Can't find nodes by Nodepath:[{0}]", this.mobjMouseInfo.NodePath);
                    Logger.Error("ClickAtSpecialNode", strError);
                    base.ReplayReportStep("ClickAtSpecialNode", EventStatus.EVENTSTATUS_FAIL, new object[] { strError, this.mobjMouseInfo.NodePath });
                    return false;
                }

                /** find rectangle **/
                Rectangle rect = objTargetNode.Bounds;
                Point ptClient = objTree.PointToScreen(rect.Location);
                

                switch (this.mobjMouseInfo.miMouseId)
                {
                    case 2:
                        LeftMouseClick(ptClient.X + rect.Width / 2, ptClient.Y + rect.Height / 2);
                        Thread.Sleep(50);
                        LeftMouseClick(ptClient.X + rect.Width / 2, ptClient.Y + rect.Height / 2);
                        Thread.Sleep(50);
                        break;
                    case 0:
                        LeftMouseClick(ptClient.X + rect.Width / 2, ptClient.Y + rect.Height / 2);
                        Thread.Sleep(50);
                        break;
                    case 1:
                        
                        LeftMouseClick(ptClient.X + rect.Width / 2, ptClient.Y + rect.Height / 2);
                        Thread.Sleep(50);
                        RightMouseClick(ptClient.X + rect.Width / 2, ptClient.Y + rect.Height / 2);
                        break;
                    default:
                        strError = "Wrong mouse command, \r\n\t0-left mouse button click \r\n\t1- right mouse button click\r\n\t2-left mouse button double click" ;
                        Logger.Error("ClickAtSpecialNode", strError ) ;
                        base.ReplayReportStep("ClickAtSpecialNode", EventStatus.EVENTSTATUS_FAIL, new string[] { this.mobjMouseInfo.miMouseId.ToString(), this.mobjMouseInfo.NodePath, strError});
                        break;
                }
                /** to verify the active node is the one **/
                Thread.Sleep(100);
                if (objTree.ActiveNode == objTargetNode)
                {
                    return true;
                }
                else
                {
                    strError = string.Format("Actived Cell [{0}] is not the one", objTree.ActiveNode.Text);
                    base.ReplayReportStep("ClickAtSpecialNode", EventStatus.EVENTSTATUS_FAIL, new object[] { strError, this.mobjMouseInfo.NodePath });
                    return false;
                }
            }
            finally
            {
                Logger.logEnd("ClickAtSpecialNode");
            }



        }

        private UltraTreeNode FindSpecialNode(TreeNodesCollection lstNodes, string[] arrNodesPath, int iLevel)
        {
            Logger.logBegin("FindSpecialNode");
            if (arrNodesPath == null) return null;
            if (arrNodesPath.Length < iLevel) return null;
            try
            {
                string strCurrentCheckKey = arrNodesPath[iLevel];
                for (int i = 0; i < lstNodes.Count; i++)
                {
                    UltraTreeNode objCurrentNode = lstNodes[i];
                    Logger.Info("FindSpecialNode", string.Format("Trying to check node:[{0}], compare to :[{1}]", objCurrentNode.Text, strCurrentCheckKey));
                    if ((string.Compare(objCurrentNode.Text, strCurrentCheckKey, true) == 0) ||/**Regular express support**/MarsUFTAddins.IMars.tiger.ReflectorForCSharp.MarsTigerUtility.RegularExpressChecking(strCurrentCheckKey, objCurrentNode.Text))
                    {
                        if (iLevel == arrNodesPath.Length - 1)
                        {
                            /** the last node to check **/
                            Logger.Info("FindSpecialNode", "find the leaf Node!");
                            return objCurrentNode;
                        }
                        else
                        {
                            Logger.Info("FindSpecialNode", string.Format("Find branch nodes:[{0}], compare to :[{1}], nested to call FindSpecialNode", objCurrentNode.Text, strCurrentCheckKey));
                            return FindSpecialNode(objCurrentNode.Nodes, arrNodesPath, iLevel + 1);
                        }
                    }
                }
                Logger.Error("FindSpecialNode", string.Format("navigated all nodes information, but no such node [{1}] find in Level:[{0}]", iLevel, strCurrentCheckKey));
                return null;
            }
            finally
            {
                Logger.logEnd("FindSpecialNode");
            }

        }

        private Boolean ParametersFormatCheckForClickSpecialNode(params object[] arrParams)
        {
            Logger.logBegin("ParametersFormatCheckForClickSpecialNode");
            try
            {
                if (arrParams == null) return false;
                if (arrParams.Length != 2) return false;
                try
                {
                    int iMouseId = int.Parse(arrParams[0].ToString());
                    this.mobjMouseInfo.miMouseId = iMouseId;
                }
                catch (Exception)
                {
                    Logger.Error("ParametersFormatCheckForClickSpecialNode", string.Format("The first parameter should be number [0/1],but it is [{0}]", arrParams[0]));
                    return false;
                }
                this.mobjMouseInfo.NodePath = arrParams[1].ToString();
                return true;
            }
            finally
            {
                Logger.logEnd("ParametersFormatCheckForClickSpecialNode");
            }

        }
    }

}
