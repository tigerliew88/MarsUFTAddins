using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Mercury.QTP.CustomServer;
using Route2NSEx.src.Marquis.systemUtil;
using DevExpress.XtraTreeList;
using DevExpress.XtraTreeList.Nodes;
using MarsTestFrame.systemUtil;
using DevExpress.XtraTreeList.ViewInfo;
using System.Drawing;
using System.Threading;
using System.Collections;
using DevExpress.XtraTreeList.Columns;
using MarsUFTAddins.IMars.tiger;
using DevExpress.XtraTreeList.Data;
using Sophis.Util.GUI;
using sophis.portfolio;

namespace MarsUFTAddins.tiger.SophisDevExpress
{
    [ReplayInterface]
    public interface DevExpress_Addins
    {
        #region Wizard generated sample code (commented)
        //		void  CustomMouseDown(int X, int Y);
        #endregion

        bool ClickAtSpecialNode(params object[] arrObj);
        bool FetchDataFromCells(object strCommand, ref object objResult, ref object objError);
#if _demo_misys_
        bool FetchDataFromCellsMisys(object strCommand, ref object objResult, ref object objError);
#endif
    }
    /// <summary>
    /// Summary description for Mars.tiger.Sophis.DevExpress.Addins.TreeView.
    /// </summary>
    public class TreeViewSever :
        CustomServerBase,
        DevExpress_Addins
    {
        // Do not call Base class methods or properties in the constructor.
        // The services are not initialized.
        private MLogger Logger = MLogger.GetLogger(typeof(TreeViewSever));

        
        public TreeViewSever()
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


        #region Replay interface implementation
        #region Wizard generated sample code (commented)
        /*		public void CustomMouseDown(int X, int Y)
		{
			MouseClick(X, Y, MOUSE_BUTTON.LEFT_MOUSE_BUTTON);
		}
*/
        #endregion

        public bool ClickAtSpecialNode(params object[] arrObj)
        {
            Logger.logBegin("ClickAtSpecialNode");
            base.PrepareForReplay();     
            if (!(base.SourceControl is TreeList))
            {
                Logger.Error("ClickAtSpecialNode", string.Format("Type is [{0}],but [{1}] and any of its Descendants is required. ", base.SourceControl.GetType().ToString(), typeof(TreeList).ToString()));
                return false;
            }


            if (arrObj.Length < 2)
            {
                Logger.Error("ClickAtSpecialNode", string.Format("For Sophis(DevExpress based) application, at least two parameters required, 1st--[Left or right mouse button]"));
            }
            
            /** 
             * Parameters:
             * 1, Left or right mouse button,0, left,1-right,2 left-double click,3-expand 
             * 2, Node text to be clicked at with path            
             * **/
                   
            /** parameters check **/
            int iMouseMark = -1;
            if (!int.TryParse(arrObj[0].ToString(), out iMouseMark))
            {
                Logger.Error("ClickAtSpecialNode", string.Format("First parameters should be [0,2-Left Mouse button/1-Right Mouse button], but the value is [{0}]", arrObj[0].ToString()));
                return false;
            }
            string strNodeToSearch = arrObj[1].ToString();
            /** for DevExpress, the last parameter is the column Id, default value is the first column to search **/
            int iColumn = 0;
            if (arrObj.Length >= 3)
            {
                /** the column information could be string **/
                /** **/
                if (!int.TryParse(arrObj[2].ToString(), out iColumn))
                {
                    Logger.Info("ClickAtSpecialNode", string.Format("Can't convert the third paramter [{0}] to Int, default value 0 is used", arrObj[2].ToString()));
                    return false;
                }
            }

            try
            {
                /** get position **/
                string[] arrNodeSearchItems = strNodeToSearch.Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries);
                TreeList objTrLst = (TreeList)base.SourceControl;
                
                bool isFindNode = true ;
                object objDataSrcTmp = objTrLst.DataSource;
                if (objTrLst.Nodes.Count<=0)
                {
                    Logger.Error("ClickAtSpecialNode", "no Node exists ") ;
                    return false ;
                }
                TreeListNode objCurrentNode = objTrLst.Nodes[0] ;
                
                if (objCurrentNode == null)
                {
                    Logger.Error("ClickAtSpecialNode", "Root node is null ") ;
                    return false ;
                }
#if _sophis7_
                if (objTrLst is Sophis.Util.GUI.CustomTreeList)
                {
                    if (objTrLst.Columns.Count>1)
                        isFindNode = FindNodeBySettings_Sophis((Sophis.Util.GUI.CustomTreeList)objTrLst,arrNodeSearchItems.Length, 0, arrNodeSearchItems, ref objCurrentNode, iColumn);
                    else
                        isFindNode = FindNodeBySettings(objTrLst.Nodes, arrNodeSearchItems.Length, 0, arrNodeSearchItems, ref objCurrentNode, iColumn);
                }
                else
                {
                    isFindNode = FindNodeBySettings(objTrLst.Nodes, arrNodeSearchItems.Length, 0, arrNodeSearchItems, ref objCurrentNode, iColumn);
                }
#else
                isFindNode = FindNodeBySettings(objTrLst.Nodes, arrNodeSearchItems.Length, 0, arrNodeSearchItems, ref objCurrentNode, iColumn);
#endif
                
                if (!isFindNode) return false;
                
                Thread.Sleep(100);
                //objTrLst.VirtualDataHelper
                /** try to double or click node **/
                RowInfo objRow = objTrLst.ViewInfo.RowsInfo[objCurrentNode];                
                CellInfo objCell = objRow.Cells[0] as CellInfo;
                if (objCell==null)
                {
                    Logger.Error("ClickAtSpecialNode", "View for cell info is null");
                    return false;
                }

                Rectangle objRectTarget = objCell.Bounds;
                Logger.Info("ClickAtSpecialNode", string.Format("Bounds:x:[{0}],y:[{1}],w:[{2}],h:[{3}]", objRectTarget.X, objRectTarget.Y,objRectTarget.Width,objRectTarget.Height));

                /** convert 2 screen and move mouse to **/
                Point ptScreenLocation = objTrLst.PointToScreen(objRectTarget.Location);
                Point ptMouseMoveTo = new Point(ptScreenLocation.X + objRectTarget.Width / 2,
                    ptScreenLocation.Y + objRectTarget.Height / 2);
                Logger.Info("ClickAtSpecialNode", string.Format("New screen position:[{0},{1}] compare to ", ptMouseMoveTo.X, ptMouseMoveTo.Y));
                /** stimulate click **/
                switch (iMouseMark)
                {
                    case 0:
                    case 2:
                        TigerMarsUtil.LeftMouseClick(ptMouseMoveTo.X, ptMouseMoveTo.Y);
                        if (iMouseMark == 2)
                        {
                            Thread.Sleep(50);
                            TigerMarsUtil.LeftMouseClick(ptMouseMoveTo.X, ptMouseMoveTo.Y);
                        }
                        break;
                    case 1:
                        TigerMarsUtil.RightMouseClick(ptMouseMoveTo.X, ptMouseMoveTo.Y);
                        break;

                    default:
                        /** Expand **/
                        TigerMarsUtil.LeftMouseClick(ptMouseMoveTo.X, ptMouseMoveTo.Y);
                        Thread.Sleep(50);
                        objCurrentNode.ExpandAll();
                        break;
                }
                
                return true;
            }
            catch (Exception e)
            {
                Logger.Error("ClickAtSpecialNode", string.Format("Errors come when trying to click special Node:[{0}], Exceptions:[{1}]", strNodeToSearch,e.Message), e);
                return false;
            }
            finally
            {
                Logger.logEnd("ClickAtSpecialNode");
            }

           
        }

        private void ExpandParents(TreeListNode objCurrentNode)
        {
            if (objCurrentNode == null) return;
            if (objCurrentNode.ParentNode != null)
            {
                if (!objCurrentNode.ParentNode.Expanded)
                {
                    ExpandParents(objCurrentNode.ParentNode);
                    objCurrentNode.ParentNode.ExpandAll();
                }
            }
        }
#if _sophis7_
        private bool FindNodeBySettings_Sophis(Sophis.Util.GUI.CustomTreeList objTreeList, int iDepths, int iCurrentDep, string[] arrNodeSearchItems, ref TreeListNode objCurrentNode, int iColumn)
        {
            Logger.logBegin("FindNodeBySettings_Sophis");
            ReflectorForCSharp objReflector = new ReflectorForCSharp();
            try
            {
                TreeListData objData = objReflector.GetMember<TreeListData>(objTreeList, "Data");
                if (objData == null)
                {
                    Logger.Info("FindNodeBySettings_Sophis", "Can't get Data from object instance via reflection.");
                    return false;
                }
                if (objData.DataList == null)
                {
                    Logger.Info("FindNodeBySettings_Sophis", "DataList is null");
                    return false;
                }
                /** +		virtualDataHelper	{Sophis.Util.GUI.CustomTLVDH}	DevExpress.XtraTreeList.TreeList.TreeListVirtualDataHelper {Sophis.Util.GUI.CustomTLVDH}
                 * -		objData.DataHelper	{DevExpress.XtraTreeList.Data.TreeListDataHelper}	DevExpress.XtraTreeList.Data.TreeListDataHelper


                object o = objData.GetValue(0, 0);
                string strCellDis = objTreeList.Nodes[0].GetDisplayText(objTreeList.Columns[0]) ;
                object or = objData.GetDataRow(objTreeList.Nodes[0].Id);
                strCellDis = objData.GetDisplayText(objTreeList.Nodes[0].Id, objTreeList.Columns[0].AbsoluteIndex, objTreeList.Nodes[0]);
                object oc = objReflector.GetMember<object>(objTreeList, "CurrencyManager");
                object ov = objReflector.CallPrivateMethod<object>(objTreeList, "GetNodeValue", new object[] {objTreeList.Nodes[0],objTreeList.Columns[0] });
                object objHelper0 = objReflector.GetMember<object>(objTreeList, "VirtualDataHelper");
                object objHelper2 = objReflector.GetPrivateProperty<object>(objTreeList, "VirtualDataHelper");
                TreeList.TreeListVirtualDataHelper objDataHelper = (TreeList.TreeListVirtualDataHelper)(objHelper0) ;
                object objCellData = objDataHelper.GetCellDataViaEvent(((VirtualDataRow)objData.DataList[0]).Node, objTreeList.Columns[0]);
                objCellData = objDataHelper.GetCellDataViaEvent(objTreeList.Nodes[0], objTreeList.Columns[0]);
                objCellData = objDataHelper.GetCellDataViaInterface(((VirtualDataRow)objData.DataList[0]).Node, objTreeList.Columns[0], null);
                objCellData = objDataHelper.GetCellDataViaInterface(objTreeList.Nodes[0], objTreeList.Columns[0], null);
                object value = typeof(TreeListVirtualData).GetField("virtualNodeToDataRowCache", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(objData);
                 *  **/
                //DevExpress.XtraTreeList.TreeList.TreeListVirtualDataHelper objHelper1 = objReflector.GetPrivateProperty<DevExpress.XtraTreeList.TreeList.TreeListVirtualDataHelper>(objTreeList, "virtualDataHelper");
                 //GetCellDataViaEvent
                for (int i = 0; i < objData.DataList.Count; i++)
                {
                    object objRawItem = objData.DataList[i];
                    if (!(objRawItem is VirtualDataRow))
                    {
                        /** I don't know what it could be **/
                        continue;
                    }
                    VirtualDataRow objRow = (VirtualDataRow)objRawItem;
                    
                    if (objRow.VirtualNode == null) continue;
                    string strCaption = objRow.VirtualNode.ToString();
                    if (TigerMarsUtil.RegularTest(arrNodeSearchItems[iCurrentDep], strCaption))
                    {
                        /** find **/
                        Logger.Info("FindNodeBySettings_Sophis",string.Format("Find Level [{0}] node information [{1}], with node value is:[{2}]", iCurrentDep, arrNodeSearchItems[iCurrentDep],strCaption));
                        if (iCurrentDep == iDepths - 1)
                        {
                            objCurrentNode = objTreeList.Nodes[i];
                            objTreeList.FocusedNode = objTreeList.Nodes[i];
                        }
                        return true;
                    }
                }
                return false;
            }
            catch (Exception e)
            {
                Logger.Error("FindNodeBySettings_Sophis",string.Format("Exceptions when trying to get data via Reflection. Exception:[{0}]", e.Message),e);
                return false;
            }
            
            
        }
#endif
        private bool FindNodeBySettings(TreeListNodes treeListNodes, int iDepths, int iCurrentDep, string[] arrSearchItems, ref TreeListNode objFindNode, int iColumnId = 0)
        {
            Logger.logBegin(string.Format("FindNodeBySettings with level [{0}]", iCurrentDep));

            try
            {
                IEnumerator en = treeListNodes.GetEnumerator();
                en.Reset();
                while (en.MoveNext())
                {
                    TreeListNode objOneItem = (TreeListNode)en.Current;
                    if (objOneItem.GetValue(iColumnId) == null) continue;
                    string strNodeCaption = objOneItem.GetValue(iColumnId).ToString();
                    Logger.Info("FindNodeBySettings", string.Format("tring to compare node value [{0}] against [{1}]", strNodeCaption, arrSearchItems[iCurrentDep]));

                    if (TigerMarsUtil.RegularTest(arrSearchItems[iCurrentDep], strNodeCaption))
                    {
                        if (iCurrentDep == iDepths - 1)
                        {
                            objFindNode = objOneItem;
                            /** expand all **/
                            ExpandParents(objOneItem);
                            return true;
                        }

                        if (objOneItem.HasChildren)
                        {
                            bool isFind = FindNodeBySettings(objOneItem.Nodes, iDepths, iCurrentDep + 1, arrSearchItems, ref objFindNode,iColumnId);
                            if (isFind)
                            {
                                Logger.Info("FindNodeBySettings", "find node");                                
                                return true;
                            }
                            else
                            {
                                Logger.Info("FindNodeBySettings", string.Format("Not find for node:[{0}] and its Descendants", strNodeCaption));
                                continue;
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }
                    else
                    {
                        continue;
                    }
                }
                return false;
            }
            catch (Exception e)
            {
                Logger.Error("FindNodeBySettings",string.Format("Exceptions:[{0}]",e.Message),e);
                return false;
            }
            finally
            {
                Logger.logBegin(string.Format("FindNodeBySettings with level [{0}]", iCurrentDep));
            }


        }

#if _demo_misys_
        public bool FetchDataFromCellsMisys(object objCommand, ref object objResult, ref object objError)
        {
            Logger.logBegin("FetchDataFromCellsMisys");
            
            const string cnst_CONDITION = "CONDITION:";
            string strCommand = objCommand == null ?  null : objCommand.ToString();
            try
            {
                
                if (string.IsNullOrEmpty(strCommand))
                {
                    throw new Exception("objCommand should not be NULL or empty");
                }

                /** check Command **/
                /*  ' format :
                    '   ROWS_LIMIT:0:0;Code/Name
                    '   ALLROWS;ColumnsName
                    '   CONDITION:COLUMN=A;ColumnName;NOSUB, Example: CONDITION:0=@FO_test;Income;NOSUB
                    '       NOSUB means don't get all sub information, other wise get all sub expanded rows
                 * */
                if (!TigerMarsUtil.RegularTest("^CONDITION:", strCommand))
                {
                    throw new Exception(string.Format("No supported Command, [{0}]", strCommand));
                }
                //objReflect.GetMember<object>(objDataSrcTmp,"")

                if (TigerMarsUtil.RegularTest("^" + cnst_CONDITION, strCommand))
                {
                    #region CONDITION MODE
                    //CONDITION:COLUMN=A;ColumnName;NOSUB, Example: CONDITION:0=@FO_test;Income;NOSUB
                    if (!TigerMarsUtil.RegularTest(@"CONDITION:\S+=\S{1,};\b\w+\b(;NOSUB){0,1}", strCommand))
                    {
                        objError = string.Format("RC should match CONDITION:\\b\\w+\\b=\\w+;\\b\\w+\\b(;NOSUB), Command:{0}", strCommand);
                        Logger.Error("FetchDataFromCellsMisys", (string)objError);
                        return false;
                    }

                    string strCmdNoHeader = strCommand.Replace(cnst_CONDITION, "");
                    string[] arrTargetField = strCmdNoHeader.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                    string strTargetField = arrTargetField[1];
                    string[] arrCondition = arrTargetField[0].Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries);                    

                    TreeList objTr = (TreeList)base.SourceControl;

                    CustomTreeList objCTrlst = (CustomTreeList)objTr;
                    object objDataSrcTmp = objCTrlst.DataSource;
                    ReflectorForCSharp objReflect = new ReflectorForCSharp();
                    CSMPortfolio objPortC = null;
                    CSMPosition objPortCsub = null ;
                    //object objPortfolio = objReflect.GetMember<object>(objDataSrcTmp, "Portfolio");
                    //CSMPortfolio objPortC = (CSMPortfolio)objPortfolio;
                    //double dIncome = objPortC.GetIncome();
                    object tmpData = objReflect.GetMember<object>(objTr, "Data");
                    UnboundData unData = (UnboundData)tmpData;
                    VirtualDataRow objVDataRow = null;
                    bool isFind = false;
                    string strHeadName = "";
                    int iLevel = -1;
                    double dDat = 0.00;
                    string strResult = "", strTmp="";
                    for (int i = 0; i < unData.DataList.Count; i++)
                    {
                        objVDataRow = (VirtualDataRow)unData.DataList[i];
                        objPortC = (CSMPortfolio)(objReflect.GetMember<object>(objVDataRow.VirtualNode, "Portfolio"));
                        iLevel = objPortC.GetLevel();
                        if ((iLevel<=1)&&(isFind))
                        {
                            break;
                        }
                        strHeadName = objPortC.GetName()==null?"":objPortC.GetName().ToString();
                        if (!isFind)
                        {
                            /** Only the root level can be taken as comparison node **/
                            if (iLevel != 1) continue;
                            if (!TigerMarsUtil.RegularTest(arrCondition[1], strHeadName)) continue;
                            else
                                isFind = true;
                        }
                        
                        switch(strTargetField.ToUpper())
                        {
                            case "INCOME":
                                dDat = objPortC.GetIncome()*1000;
                                strTmp = dDat.ToString("0.");
                                break;
                            case "NAV":
                                dDat = objPortC.GetNetAssetValue() * 1000;
                                strTmp = dDat.ToString("0.");
                                break;
                            case "RESULT":
                                dDat = objPortC.GetResult() * 1000;
                                strTmp = dDat.ToString("0.");
                                break;
                            case "REALIZED":
                                dDat = objPortC.GetRealised() * 1000;
                                strTmp = dDat.ToString("0.");
                                break;
                            case "BALANCE":
                                dDat = objPortC.GetBalance() * 1000;
                                strTmp = dDat.ToString("0.");
                                break;
                            default:
                                throw new Exception(string.Format("Not supported field :[{0}]", strTargetField));
                        }
                        if (string.IsNullOrEmpty(strResult))
                        {
                            strResult = string.Format("[{0}{1}]", strHeadName,FormatStrWithTab(0,strTmp));
                        }
                        else
                        {
                            strResult = string.Format("{0}\n\r{1}",strResult, string.Format("[{0}{1}]", strHeadName, FormatStrWithTab(0, strTmp)));
                        }

                        for (int j = 0; j < objVDataRow.Children.Count; j++)
                        {
                            //objVDataRow.Children[j];
                            objPortCsub = (CSMPosition)objReflect.GetMember<object>(objVDataRow.Children[j], "Position");
                            strHeadName = objVDataRow.Children[j].ToString();
                            if (objPortCsub == null) continue;
                            switch (strTargetField.ToUpper())
                            {
                                case "INCOME":
                                    dDat = objPortCsub.GetIncome() * 1000;
                                    strTmp = dDat.ToString("0.");
                                    break;
                                case "NAV":
                                    dDat = objPortCsub.GetNetAssetValue() * 1000;
                                    strTmp = dDat.ToString("0.");
                                    break;
                                case "RESULT":
                                    dDat = objPortCsub.GetResult() * 1000;
                                    strTmp = dDat.ToString("0.");
                                    break;
                                case "REALIZED":
                                    dDat = objPortCsub.GetRealised() * 1000;
                                    strTmp = dDat.ToString("0.");
                                    break;
                                case "BALANCE":
                                    dDat = objPortCsub.GetBalance() * 1000;
                                    strTmp = dDat.ToString("0.");
                                    break;
                                default:
                                    throw new Exception(string.Format("Not supported field :[{0}]", strTargetField));
                            }
                            strResult = string.Format("{0}\n\r{1}", strResult, string.Format("[{0}    {1}]", strHeadName, FormatStrWithTab(0, strTmp)));
                        }
                        break;
                    }
                    objResult = strResult;
                    return true;
                    #endregion //CONDTION MODE
                }
               
                objError = string.Format("Unsupported Command formatter:[{0}]", strCommand);
                return false;
            }
            catch (Exception e)
            {
                objError = e.Message;
                Logger.Error("FetchDataFromCellsMisys", string.Format("Errors come across. Exception:[{0}]", e.Message), e);
                return false;
            }
            finally
            {
                Logger.logEnd("FetchDataFromCellsMisys");
            }
        }
        private string FormatStrWithTab(int iLevel, string strSrc)
        {
            string strTmp = "";
            for (int i=0;i<=iLevel;i++)
            {
                strTmp += "    ";
            }
            return string.Format("{0}{1}",strTmp, strSrc);
        }
#endif

        public bool FetchDataFromCells(object objCommand, ref object objResult, ref object objError)
        {
            Logger.logBegin("FetchDataFromCells");
            const string cnst_allrow = "ALLROWS;";
            const string cnst_CONDITION ="CONDITION:" ;
            const string cnst_ROWS_LIMIT="ROWS_LIMIT:" ;
            try
            {
                string strCommand = objCommand == null ? objCommand.ToString() : null;
                if (string.IsNullOrEmpty(strCommand))
                {
                    throw new Exception("objCommand should not be NULL or empty");
                }

                /** check Command **/
                /*  ' format :
                    '   ROWS_LIMIT:0:0;Code/Name
                    '   ALLROWS;ColumnsName
                    '   CONDITION:COLUMN=A;ColumnName;NOSUB, Example: CONDITION:0=@FO_test;Income;NOSUB
                    '       NOSUB means don't get all sub information, other wise get all sub expanded rows
                 * */
                if (!(TigerMarsUtil.RegularTest("^" + cnst_allrow, strCommand) || TigerMarsUtil.RegularTest("^ROWS_LIMIT", strCommand) || TigerMarsUtil.RegularTest("^CONDITION:", strCommand)))
                {
                    throw new Exception(string.Format("No supported Command, [{0}]", strCommand));
                }

                TreeList objTr = (TreeList)base.SourceControl;
                int iStartRow = 0, iEndRow = -1;
                string strColumn = "";

                if (TigerMarsUtil.RegularTest("^" + cnst_allrow, strCommand))
                {
                    #region AllROWS dealing
                    iStartRow = 0;
                    strColumn = strCommand.Replace(cnst_allrow, "");
                    int iColId = GetColumnIdByCaption(objTr, strColumn);

                    if (iColId == -1)
                    {
                        objError = string.Format("Can't get column info with caption:[{0}]", strColumn);
                        Logger.Error("FetchDataFromCells", (string)objError);
                        return false;
                    }

                    TreeListColumn objCol = objTr.Columns[0];

                    foreach (TreeListNode objOneNode in objTr.Nodes)
                    {
                        if (strColumn.Length == 0)
                            strColumn = objOneNode.GetValue(iColId).ToString();
                        else
                            string.Format("{1}", objOneNode.GetValue(iColId).ToString());
                    }
                    objResult = strColumn;
                    return true;
                    #endregion //AllROWS dealing
                }
                if (TigerMarsUtil.RegularTest("^" + cnst_CONDITION, strCommand))
                {
                    #region CONDITION MODE
                    //CONDITION:COLUMN=A;ColumnName;NOSUB, Example: CONDITION:0=@FO_test;Income;NOSUB
                    if (!TigerMarsUtil.RegularTest(@"CONDITION:\b\w+\b=\w+;\b\w+\b(;NOSUB){0,1}", strCommand))
                    {
                        objError = string.Format("RC should match CONDITION:\\b\\w+\\b=\\w+;\\b\\w+\\b(;NOSUB), Command:{0}", strCommand);
                        Logger.Error("FetchDataFromCells", (string)objError);
                        return false;
                    }
                    string strCmd = strCommand.Replace(cnst_CONDITION, "");
                    string[] arrCmdComma = strCmd.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                    int iColId = -1;
                    string[] arrCmdCondition = arrCmdComma[0].Split(new string[] { "=" }, StringSplitOptions.RemoveEmptyEntries);
                    bool isColIndex = int.TryParse(arrCmdCondition[0], out iColId);
                    if (!isColIndex)
                    {
                        iColId = this.GetColumnIdByCaption(objTr, arrCmdCondition[0]);
                        if (iColId == -1)
                        {
                            Logger.Error("FetchDataFromCells", (string)(objError = string.Format("Can't find condition column name like [{0}] from Treelist type spreadsheet", arrCmdCondition[0])));
                            return false;
                        }
                    }
                    /** GET TAGET COLUMN **/
                    int iTargetColId = this.GetColumnIdByCaption(objTr, arrCmdComma[1]);
                    if (iTargetColId == -1)
                    {
                        Logger.Error("FetchDataFromCells", (string)(objError = string.Format("Can't find target column name like [{0}] from Treelist type spreadsheet", arrCmdComma[1])));
                        return false;
                    }

                    /** find Condtion node **/
                    string strResult = "";
                    bool isSubNodeInclude = arrCmdComma.Length <= 2 ? true : string.Compare(arrCmdComma[2], "NOSUB", true) != 0;

                    foreach (TreeListNode objNode in objTr.Nodes)
                    {
                        string strCaption = objNode.GetValue(iColId).ToString();
                        string strTargetCaption = objNode.GetValue(iTargetColId).ToString();
                        if (TigerMarsUtil.RegularTest(arrCmdCondition[1], strCaption))
                        {
                            strResult = string.IsNullOrEmpty(strResult) ? strTargetCaption : string.Format("{0}\r\n{1}", strResult, strTargetCaption);
                            /** 判断是否有子 节点 and expanded ; arrCmdComma.Length>2 means sub**/
                            if (objNode.HasChildren && objNode.Expanded && isSubNodeInclude)
                            {
                                strTargetCaption = this.CaptureAllValuesOfSubNodes(objNode.Nodes, iTargetColId);
                                strResult = string.IsNullOrEmpty(strResult) ? strTargetCaption : string.Format("{0}\r\n{1}", strResult, strTargetCaption);
                            }
                        }
                    }
                    objResult = strResult;
                    return true;
                    #endregion //CONDTION MODE
                }
                if (TigerMarsUtil.RegularTest("^" + cnst_ROWS_LIMIT, strCommand))
                {
                    #region Rows Limit
                    //ROWS_LIMIT:0:0;Code/Name
                    string[] arrCmdComma = strCommand.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
                    if (arrCmdComma.Length < 2)
                    {
                        objError = string.Format("For Rows_limit model, command formatter should be Rows_limit:[rowStart]:[rowEnd];Column name or column Index, but current command string is [{0}]", strCommand);
                        Logger.Error("FetchDataFromCells", (string)objError);
                        return false;
                    }
                    string strCmd = arrCmdComma[0].Replace(cnst_ROWS_LIMIT, "");
                    string[] arrRows = strCmd.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                    if (arrRows.Length != 2)
                    {
                        objError = string.Format("For Rows_limit model, command formatter should be Rows_limit:[rowStart]:[rowEnd];Column name or column Index, but current RowStart and RowEnd string is [{0}]", strCmd);
                        Logger.Error("FetchDataFromCells", (string)objError);
                        return false;
                    }
                    bool isInt = int.TryParse(arrRows[0], out iStartRow) && int.TryParse(arrRows[1], out iEndRow);
                    if ((!isInt) || (iEndRow < iStartRow) || (iEndRow <= 0) || (iStartRow <= 0) || (iEndRow >= objTr.Nodes.Count) || (iStartRow >= objTr.Nodes.Count))
                    {
                        objError = string.Format("For Rows_limit model, command formatter should be Rows_limit:[rowStart]:[rowEnd];Column name or column Index. StartRow should less than EndRow, but current RowStart and RowEnd string is [{0}]", strCmd);
                        Logger.Error("FetchDataFromCells", (string)objError);
                        return false;
                    }
                    int iColId = -1;
                    bool isVisibleNode = true;
                    if (!int.TryParse(arrCmdComma[1], out iColId))
                    {
                        foreach (TreeListColumn item in objTr.Columns)
                        {
                            if (TigerMarsUtil.RegularTest(arrCmdComma[1], item.Caption))
                            {
                                iColId = item.AbsoluteIndex + 1;
                                isVisibleNode = false;
                                break;
                            }
                        }
                    }

                    if ((iColId < 0) || (iColId > objTr.Columns.Count))
                    {
                        objError = string.Format("No such column [{0}] exists or Exceed the total column count.");
                        Logger.Error("FetchDataFromCells", (string)objError);
                        return false;
                    }
                    string strResult = "";
                    for (int i = iStartRow; i < iEndRow; i++)
                    {
                        string strCell = objTr.Nodes[i - 1].GetValue(iColId).ToString();
                        strResult = string.IsNullOrEmpty(strResult) ? strCell : string.Format("{0}\r\n{1}", strResult, strCell);
                    }
                    objResult = strResult;
                    return true;
                    #endregion //Rows Limit
                }
                objError = string.Format("Unsupported Command formatter:[{0}]", strColumn);
                return false;
            }
            catch (Exception e)
            {
                objError = e.Message;
                Logger.Error("FetchDataFromCells", string.Format("Errors come across. Exception:[{0}]", e.Message), e);
                return false;
            }
            finally
            {
                Logger.logEnd("FetchDataFromCells");
            }
            
        }

        private string CaptureAllValuesOfSubNodes(TreeListNodes lstNodes, int iColId)
        {
            Logger.logBegin("CaptureAllValuesOfSubNodes");
            string strResult = "",strCaption;
            foreach (TreeListNode objNode in lstNodes)
            {
                strCaption=objNode.GetValue(iColId).ToString() ;
                strResult = string.IsNullOrEmpty(strResult) ? strCaption : string.Format("{0}\r\n{1}", strResult, strCaption);
                if (objNode.HasChildren && objNode.Expanded)
                {
                    strCaption = CaptureAllValuesOfSubNodes(objNode.Nodes, iColId);
                    strResult = string.IsNullOrEmpty(strResult) ? strCaption : string.Format("{0}\r\n{1}", strResult, strCaption);
                }
            }
            Logger.logEnd("CaptureAllValuesOfSubNodes");
            return strResult;
        }

        private int GetColumnIdByCaption(TreeList objTr, string strColumn)
        {
            foreach (TreeListColumn objOneColumn in objTr.VisibleColumns)
            {
                if (TigerMarsUtil.RegularTest(strColumn, objOneColumn.Caption))
                {
                    return objOneColumn.AbsoluteIndex;
                    
                }
            }
            return -1 ;
        }

        //private bool FindNodeBySettings(TreeListNode objOneItem, int iDepths, int iCurrentLevel, string[] arrSearchItems, int iColumnId)
        //{
        //    Logger.logBegin("FindNodeBySettings");
        //    if (objOneItem == null)
        //    {
        //        Logger.Error("FindNodeBySettings", "Sub-nodes are null");
        //        return false;
        //    }
        //    TreeListNode objTmp = null ;
        //    while ((objTmp = objOneItem.NextNode) != null)
        //    {
        //        if (TigerMarsUtil.RegularTest(arrSearchItems[iCurrentLevel], objTmp.GetValue(iColumnId).ToString()))
        //        {
        //            Logger.Info("FindNodeBySettings", string.Format("find matched value :[{0}], Level:[{1}]", objTmp.GetValue(iColumnId).ToString(), iCurrentLevel));
        //            iCurrentLevel += 1;
        //            if (iCurrentLevel == iDepths) return true;

        //            if (objTmp.HasChildren) return FindNodeBySettings(objTmp.Nodes
        //        }
        //    }
        //    Logger.logEnd("FindNodeBySettings");
        //    return false;
        //}
        #endregion
    }
}
