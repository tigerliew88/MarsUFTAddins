using System;
using System.Collections.Generic;
#if tiger_dotNet4
using System.Linq;
#endif
using System.Text;
using Mercury.QTP.CustomServer;
using Route2NSEx.src.Marquis.systemUtil;
using Infragistics.Win.UltraWinGrid;
using System.Windows.Forms;
using Infragistics.Win;
using System.Drawing;
using System.Threading;
using MarsTestFrame.systemUtil;

namespace MarsUFTAddins.IMars.tiger.infragistics.v12
{
    internal enum E_RowColumnType
    {
        eType_Number = 0x01,
        eType_String,
    }

    internal class Mars_RowColumnInfo
    {
        #region properties
        public E_RowColumnType CurrentDataType;
        public object currentData;
        #endregion

        public int convertData2Int()
        {
            return (int)currentData;
        }

        internal Mars_RowColumnInfo(object objData)
        {
            this.CurrentDataType = E_RowColumnType.eType_String;
            if (objData is int)
            {
                this.CurrentDataType = E_RowColumnType.eType_Number;
            }
            int iNumber;
            if (int.TryParse(objData.ToString(), out iNumber))
            {
                this.CurrentDataType = E_RowColumnType.eType_Number;
                currentData = iNumber;
                return;
            }

            this.currentData = objData;
        }
    }

    [ReplayInterface]
    public interface IMarsTigerTableServer : IMarsTigerReplayBase
    {
        void MarsTigerActiveCell(params object[] arrRC);
        void MarsTigerSetDropListModeValue(object strData);
        void MarsTigerSetCurrentCellText(object strData);
        void MarsTigerSendTabKeyToCurrent();
        void MarsTigerSendEnterkeyToCurrent();
        void MarsRightClick(object objRowNumber, object objColumnName);
        void MarsLeftClick(object objRowNumber, object objColumnName, object clickTimes = null);
        int SearchColumnId(object objColumnInfo);
        object StoreColumnsValue(object objColumnId, object objStartRow, object objEndRow, object objSplitPerRow, ref object isRight);
        object GetTotalRowCount();
    }


    public class MarsTigerTableServer : MarsTigerServerBase, IMarsTigerTableServer
    {
        private static MLogger Logger = MLogger.GetLogger(typeof(MarsTigerTableServer));




        #region inherited from QTP/UFT
        public override void InitEventListener()
        {
            // disable AddEvent
            Logger.logBegin("InitEventListener");
            //base.InitEventListener();
            Logger.logEnd("InitEventListener");
        }
        /*
        protected override void tigerMouseDown(object objSender, MouseEventArgs e)
        {
            Logger.logBegin("tigerMouseDown");
            UltraGrid objGrid = (UltraGrid)SourceControl;
            //base.RecordFunction("ActiveCell", RecordingMode.RECORD_SEND_LINE, new int[] {  objGrid.ActiveCell.Row.Index, objGrid.ActiveCell.Column.Index });

            Logger.logEnd("tigerMouseDown");
        }
        */
        public override void ReleaseEventListener()
        {
            //base.ReleaseEventListener();
        }
        #endregion
        /*
        protected override void AddEvent()
        {
            //base.AddEvent();
        }
        */
        #region for addins
        public int SearchColumnId(object objColumnInfo)
        {
            Logger.logBegin("SearchColumnId");
            int iColumnResult = -1 ;
            UltraGrid objGrid = (UltraGrid)base.SourceControl;
            if (objGrid.Rows.Count <= 0)
            {
                Logger.Error("SearchColumnId", "Row count is zero");
                return iColumnResult;
            }
            string strColumnInfo = objColumnInfo==null?"":objColumnInfo.ToString() ;
            try 
	        {	        
		        if (string.IsNullOrEmpty(strColumnInfo))
                {
                    Logger.Error("SearchColumnId","ColumnInfo passed is null") ;
                    return -1 ;
                }
                UltraGridRow objFirstRow = objGrid.Rows[0];
                string strColumn="", strHeader = "";
                for (int i = 0; i < objFirstRow.Cells.Count; i++)
                {
                    strColumn = objFirstRow.Cells[i].Column.Key.ToUpper();
                    strHeader = objFirstRow.Cells[i].Column.Header.Caption.ToUpper();
                    if ((strColumnInfo.ToUpper().CompareTo(strColumn)==0)||(strColumnInfo.ToUpper().CompareTo(strHeader)==0)
                        ||(TigerMarsUtil.RegularTest(strColumnInfo,strColumn))||(TigerMarsUtil.RegularTest(strColumnInfo,strHeader)))
                    {
                        Logger.logEnd("SearchColumnId");
                        return iColumnResult=i ;
                    }
                }
                return iColumnResult ;
            }
	        catch (Exception e)
	        {
                Logger.Error("SearchColumnId", string.Format("Exception:[{0}]",e.Message),e);
		        return iColumnResult ;
	        }finally{
                Logger.logEnd("SearchColumnId");
            }
        }

        public object GetTotalRowCount()
        {
            UltraGrid objGrid = (UltraGrid)base.SourceControl;
            return objGrid == null ? 0 : objGrid.Rows.Count;
        }

        public object StoreColumnsValue(object objColumnId, object objStartRow, object objEndRow, object objSplitPerRow, ref object isRight)
        {
            Logger.logBegin("StoreColumnsValue");
            Logger.Info("StoreColumnsValue", string.Format("parameter:\r\n\tobjColumnId:[{0}] \t objStartRow:[{1}], \tobjEndRow:[{2}] \tobjSplitPerRow[{3}]",
                objColumnId, objStartRow, objEndRow, objSplitPerRow));
            string strResult = "";
            UltraGrid objGrid = (UltraGrid)base.SourceControl;
            /** check parameters **/
            int iColumnId = TigerMarsUtil.ConvertObject2Int(objColumnId,-1);
            isRight = false;
            if (iColumnId < 0)
            {
                Logger.Error("StoreColumnsValue", string.Format("Column Id is < 0 or not a number:[{0}]", objColumnId));
                return strResult="";
            }
            int iStartRow = TigerMarsUtil.ConvertObject2Int(objStartRow, -1);
            if (iStartRow < 0)
            {
                Logger.Error("StoreColumnsValue", string.Format("iStartRow is < 0 or not a number:[{0}]", objStartRow));
                return "";
            }
            int iEndRow = TigerMarsUtil.ConvertObject2Int(objEndRow, -1);
            if (iEndRow < 0)
            {
                Logger.Error("StoreColumnsValue", string.Format("iEndRow is < 0 or not a number:[{0}]", objEndRow));
                return "";
            }
            if (objGrid.Rows.Count<=0)
            {
                isRight = true;
                Logger.Error("StoreColumnsValue", "No row exists for the Datagrid") ;
                return "" ;
            }
            if ((iStartRow >= objGrid.Rows.Count) || (iEndRow >= objGrid.Rows.Count)|| (iStartRow<0)||(iEndRow<0))
            {
                Logger.Error("StoreColumnsValue",string.Format("StartRow [{0}] or EndRow [{1}] is out of range", iStartRow, iEndRow));
                return "";
            }
            if (iStartRow > iEndRow)
            {
                Logger.Error("StoreColumnsValue", string.Format("StartRow [{0}] is greater than endRow [{1}] number.", iStartRow, iEndRow));
                return "";
            }
            string strSplit = objSplitPerRow == null ? "\n" : objSplitPerRow.ToString();
            string strCell = "";
            isRight = true;
            for (int i = iStartRow; i <= iEndRow; i++)
            {
                strCell = objGrid.Rows[i].Cells[iColumnId].Text;
                if (TigerMarsUtil.RegularTest("179,769,313,486,23.*", strCell))
                {
                    strCell = "";
                }
                strResult = string.Format("{0}{1}{2}", strResult,strSplit, strCell);
            }

            Logger.logEnd("StoreColumnsValue");
            return strResult;
        }
        #endregion
        /** set the current embeded text value **/
        public void MarsTigerSetCurrentCellText(object strData)
        {
            Logger.logBegin("MarsTigerSetCurrentCellText");
            base.SendKeys(strData.ToString());
            Logger.logEnd("MarsTigerSetCurrentCellText");
        }

        public void MarsTigerSendTabKeyToCurrent()
        {
            Logger.logBegin("MarsTigerSendTabKeyToCurrent");
            base.KeyDown(15);
            Logger.logEnd("MarsTigerSendTabKeyToCurrent");
        }
        public void MarsTigerSendEnterkeyToCurrent()
        {
            Logger.logBegin("MarsTigerSendEnterkeyToCurrent");
            base.KeyDown((byte)Keys.Enter);
            Logger.logEnd("MarsTigerSendEnterkeyToCurrent");
        }

        public void MarsTigerSetDropListModeValue(object strData)
        {
            Logger.logBegin("MarsTigerSetDropListModeValue");
            UltraGrid objGrid = (UltraGrid)base.SourceControl;
            UltraGridCell objCell = objGrid.ActiveCell;
            string strError = "";
            if (objCell == null)
            {
                strError = string.Format("No Active cell before change value for droplist:[{0}]", strData);
                QtpShowException(strError);
                return;
            }
            objCell.SetValue(strData, false);
            Logger.logEnd("MarsTigerSetDropListModeValue");
        }

        /**  **/
        public void MarsTigerActiveCell(params object[] arrRC)
        {

            Logger.logBegin("MarsTigerActiveCell");
            if (arrRC.Length != 2)
            {
                Logger.Error("MarsTigerActiveCell", "Only two parameters are accepted. ");
                return;
            }
            PrepareForReplay();

            object oRow = arrRC[0], oColumn = arrRC[1];
            UltraGrid objGrid = (UltraGrid)SourceControl;
            string strError = "";

            Mars_RowColumnInfo oRowInfo = new Mars_RowColumnInfo(oRow),
                oColInfo = new Mars_RowColumnInfo(oColumn);
            try
            {
                if ((objGrid.Rows == null ? -1 : objGrid.Rows.Count) <= 0)
                {
                    base.ReplayReportStep("MarsTigerActiveCell", EventStatus.EVENTSTATUS_FAIL, new object[] { oRow, oColumn, "RowCount is 0" });
                    return;
                }

                if (oRowInfo.CurrentDataType == E_RowColumnType.eType_String)
                {
                    // the first column of the row with fixed caption, which should be implement later
                    throw new NotImplementedException();
                }
                /** find cells **/
                UltraGridRow objRow = mFindRowByRowNumber(oRowInfo.convertData2Int(), objGrid);
                if (objRow == null)
                {
                    strError = string.Format("no such row number [{0}] is availble, row number stars 0", objRow);
                    base.ReplayReportStep("MarsTigerActiveCell", EventStatus.EVENTSTATUS_FAIL, new object[] { oRow, oColumn, strError });
                    return;
                }
                /** find cells by oColumn, int or header **/
                UltraGridCell objCell = mFindCellByRowAndColumnInfo(objRow, oColInfo);
                if (objCell == null)
                {
                    strError = string.Format("Can't find sepecial cell. \r\n\tRow:[{0}]\r\n\tColumn:[{1}]", oRow, oColumn);
                    base.ReplayReportStep("MarsTigerActiveCell", EventStatus.EVENTSTATUS_FAIL, new object[] { oRow, oColumn, strError });
                    return;
                }
                /** 获得 cell的rectangle, 然后调用mouse事件 **/
                UIElement objUIElement = objCell.GetUIElement();
                Thread.Sleep(200);
                if (objUIElement == null)
                {
                    strError = string.Format("Can't get UIElement: Cell is invisible,no such cell or exceptions? \r\n\tRow:[{0}]\r\n\tColumn:[{1}]", oRow, oColumn);
                    base.ReplayReportStep("MarsTigerActiveCell", EventStatus.EVENTSTATUS_FAIL, new object[] { oRow, oColumn, strError });
                    return;
                }
                Rectangle rect = objUIElement.Rect;
                Point pt = rect.Location;
                base.MouseClick(pt.X + rect.Width / 2, pt.Y + rect.Height / 2, MOUSE_BUTTON.LEFT_MOUSE_BUTTON);

                //LeftMouseClick(pt.X, pt.Y);
            }
            finally
            {
                Logger.logEnd("MarsTigerActiveCell");
            }


        }

        private UltraGridCell mFindCellByRowAndColumnInfo(UltraGridRow objRow, Mars_RowColumnInfo oColInfo)
        {
            Logger.logBegin("mFindCellByRowAndColumnInfo");
            try
            {
                if (objRow == null)
                {
                    Logger.Error("mFindCellByRowAndColumnInfo", "Row is null");
                    return null;
                }
                switch (oColInfo.CurrentDataType)
                {
                    case E_RowColumnType.eType_Number:
                        int iCurrentColId = 0;
                        for (int i = 0; i < objRow.Cells.Count; i++)
                        {
                            UltraGridCell objCurrentCell = objRow.Cells[i];
                            if (objCurrentCell.Column.Hidden) continue;
                            if (iCurrentColId == oColInfo.convertData2Int())
                            {
                                Logger.Info("mFindCellByRowAndColumnInfo", "find cell");
                                return objCurrentCell;
                            }
                            iCurrentColId++;
                        }
                        break;
                    case E_RowColumnType.eType_String:
                        string strColumnName = "", strColumnInnderName = "";
                        for (int i = 0; i < objRow.Cells.Count; i++)
                        {
                            UltraGridCell objCurrentCell = objRow.Cells[i];
                            if (objCurrentCell.Column.Hidden) continue;
                            strColumnName = objCurrentCell.Column.Header.Caption;
                            strColumnInnderName = objCurrentCell.Column.Key;
                            if ((string.Compare(strColumnName, oColInfo.currentData.ToString(), true) == 0) || ((string.Compare(strColumnInnderName, oColInfo.currentData.ToString(), true) == 0)))
                            {
                                Logger.Info("mFindCellByRowAndColumnInfo", "find cell");
                                return objCurrentCell;
                            }
                        }
                        break;
                }


                return null;
            }
            finally
            {
                Logger.logEnd("mFindCellByRowAndColumnInfo");
            }

        }

        private UltraGridRow mFindRowByRowNumber(int iRow, UltraGrid objGrid)
        {
            Logger.logBegin("mFindRowByRowNumber");
            try
            {
                if (iRow < objGrid.Rows.Count) return objGrid.Rows[iRow];
                return null;
            }
            finally
            {
                Logger.logEnd("mFindRowByRowNumber");
            }

        }

        private void ReportError(string strStep, string strError)
        {
            this.ReplayReportStep(strStep, EventStatus.EVENTSTATUS_FAIL, new object[]{strError});
            this.ReplayThrowError(strError);
        }

        public void MarsRightClick(object objRowNumber, object objColumnName)
        {

            mClickAtGrid("RightClick", objRowNumber, objColumnName, MOUSE_BUTTON.RIGHT_MOUSE_BUTTON);

        }

        private void mClickAtGrid(string strMethodName, object objRowNumber, object objColumnName, MOUSE_BUTTON clickButton, int iClickTimes = 1)
        {
            Logger.logBegin(strMethodName);
            PrepareForReplay();
            string strErr = "";
            try
            {
                if ((objRowNumber == null) || (objColumnName == null))
                {
                    this.ReplayReportStep(strMethodName, EventStatus.EVENTSTATUS_FAIL,
                        new object[] {
                            strErr = string.Format("parameter should not be null. Parameters:Row-[{0}], ColumnName-[{1}]", objRowNumber==null?"null":objRowNumber.ToString(),
                            objColumnName==null?"null":objColumnName.ToString())}
                    );
                    this.ReplayThrowError(string.Format(strErr));
                    return;
                }
                int iNumber = Int32.Parse(objRowNumber.ToString());
                string strColumn = objColumnName.ToString();

                UltraGrid objGrid = (UltraGrid)SourceControl;
                if (iNumber > objGrid.Rows.Count)
                {
                    strErr = string.Format("Row number [{0}] is larger than the current Grid contains :[{1}]", iNumber, objGrid.Rows.Count);
                    ReportError(strMethodName, strErr);
                    return;
                }
                if (objGrid.Rows.Count == 0)
                {
                    strErr = "Row number is Zero, can't rightClick on Content. ";
                    ReportError(strMethodName, strErr);
                    return;
                }
                UltraGridRow objRow = objGrid.Rows[iNumber];
                Mars_RowColumnInfo objRowInfo = new Mars_RowColumnInfo(objRowNumber),
                    objColunInfo = new Mars_RowColumnInfo(objColumnName);
                UltraGridCell objCell = this.mFindCellByRowAndColumnInfo(objRow, objColunInfo);
                if (objCell == null)
                {
                    strErr = string.Format("No such Column can be found on the table:[{0}]", strColumn);
                    ReportError(strMethodName, strErr);
                    return;
                }
                UIElement objUIElement = objCell.GetUIElement();
                Thread.Sleep(200);
                if (objUIElement == null)
                {
                    strErr = string.Format("Can't get UIElement: Cell is invisible,no such cell or exceptions? \r\n\tRow:[{0}]\r\n\tColumn:[{1}]", iNumber, strColumn);
                    ReportError(strMethodName, strErr);
                    return;
                }
                Rectangle rect = objUIElement.Rect;
                Point pt = rect.Location;
                if (iClickTimes != 1)
                    base.MouseDblClick(pt.X + rect.Width / 2, pt.Y + rect.Height / 2, clickButton);
                else
                    base.MouseClick(pt.X + rect.Width / 2, pt.Y + rect.Height / 2, clickButton);
                
                this.ReplayReportStep(strMethodName, EventStatus.EVENTSTATUS_GENERAL, new object[] { "Passed" });
            }
            catch (Exception e)
            {
                strErr = string.Format("Can't continue to RightClick, Exceptions :[{0}]", e.Message);
                Logger.Error(strMethodName, strErr, e);
                this.ReplayReportStep("RightClick", EventStatus.EVENTSTATUS_FAIL,
                    new Object[] { strErr });
                this.ReplayThrowError(string.Format(strErr));
            }
        }

        public void MarsLeftClick(object objRowNumber, object objColumnName, object clickTimes = null)
        {
            int iClickTimes = 1;
            string strErr = "";
            try
            {
                iClickTimes = int.Parse(clickTimes.ToString());
                if (iClickTimes < 0)
                {
                    Logger.Error("LeftClick", strErr=string.Format("Click times should be 1 or 2, but current value is:[{0}]", clickTimes==null?"NULL":clickTimes.ToString()));
                    this.ReportError("LeftClick", strErr);
                    return;
                }
            }
            catch (Exception)
            {
                Logger.Error("LeftClick", strErr =string.Format("Last parameter should be a number:[{0}]", clickTimes == null ? "NULL" : clickTimes.ToString()));
                this.ReportError("LeftClick", strErr);
                return;
            }

            mClickAtGrid("LeftClick", objRowNumber, objColumnName, MOUSE_BUTTON.LEFT_MOUSE_BUTTON, iClickTimes >= 2 ? 2 : iClickTimes);
        }
    }
}
