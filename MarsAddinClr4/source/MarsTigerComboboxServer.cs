using System;
using System.Collections.Generic;
#if tiger_dotNet4
using System.Linq;
#endif
using System.Text;
using Route2NSEx.src.Marquis.systemUtil;
using Infragistics.Win.UltraWinEditors;

namespace MarsUFTAddins.IMars.tiger.infragistics.v12
{
    public class MarsTigerComboboxServer:MarsTigerServerBase
    {
        private static MLogger Logger = MLogger.GetLogger(typeof(MarsTigerComboboxServer));

        
        protected override void AddEvent()
        {
            Logger.logBegin("AddEvent");
            //base.AddEvent();
            Delegate eValueChanged = new System.EventHandler(this.ComboboxValueChanged);
            
            //base.AddHandler(MarsTigerServerConst.CNST_EVNT_VALUECHANGED, eValueChanged);
            base.AddHandler(MarsTigerServerConst.CNST_EVNT_SELECTCHANGED, eValueChanged);
            Delegate eChangeCmmt = new System.EventHandler(this.ComboboxValueChangedCmmt);
            base.AddHandler(MarsTigerServerConst.CNST_EVNT_SELECTCHANGEDCOMMIT, eChangeCmmt);
            Logger.logEnd("AddEvent");
        }

        protected void ComboboxValueChanged(object sender, System.EventArgs e)
        {
            Logger.logBegin("ComboboxValueChanged");
            UltraComboEditor objCombobox = sender as UltraComboEditor;
#if _tigerDebug
            base.RecordFunction("test", Mercury.QTP.CustomServer.RecordingMode.RECORD_KEEP_LINE, "Genearted by tiger");
#endif
            base.RecordFunction(MarsTigerServerConst.CNST_EVNT_SELECT,Mercury.QTP.CustomServer.RecordingMode.RECORD_SEND_LINE,  objCombobox.Value);
            Logger.logEnd("ComboboxValueChanged");
        }

        protected void ComboboxValueChangedCmmt(object sender, System.EventArgs e)
        {
            Logger.logBegin("ComboboxValueChangedCmmt");
            UltraComboEditor objCombobox = sender as UltraComboEditor;
#if _tigerDebug
            base.RecordFunction("test", Mercury.QTP.CustomServer.RecordingMode.RECORD_KEEP_LINE, "Genearted by tiger");
#endif
            base.RecordFunction(MarsTigerServerConst.CNST_EVNT_SELECT, Mercury.QTP.CustomServer.RecordingMode.RECORD_SEND_LINE, objCombobox.Value);
            Logger.logEnd("ComboboxValueChangedCmmt");
        }
    }
}
