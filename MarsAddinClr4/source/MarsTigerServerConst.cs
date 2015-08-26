using System;
using System.Collections.Generic;
#if tiger_dotNet4
using System.Linq;
#endif
using System.Text;

namespace MarsUFTAddins.IMars.tiger
{
    public sealed class MarsTigerServerConst
    {
        public const string CNST_EVNT_MOUSEDOWN = "MouseDown";
        public const string CNST_EVNT_VALUECHANGED = "ValueChanged";
        public const string CNST_EVNT_SELECTCHANGED = "SelectionChanged";
        public const string CNST_EVNT_SELECTCHANGEDCOMMIT = "SelectionChangeCommitted";
        public const string CNST_EVNT_SELECT = "Select";

        public const string CNST_TOOLBAR_CLICK="ClickToolbarTool" ;
        public const string CNST_TOOLBAR_CLICK_PARA_PREFIX_BUTTON_REG = @"^Button:";
        public const string CNST_TOOLBAR_CLICK_PARA_PREFIX_BUTTON = "Button:";
        public const string CNST_TOOLBAR_CLICK_PARA_PREFIX_MENU_REG = @"^Menu;";
        public const string CNST_TOOLBAR_CLICK_PARA_PREFIX_MENU = @"Menu:";

        #region Infragistics properties
        public const string CNST_INFR_TOOLBAR_MANAGER = "ToolbarsManager";
        #endregion

    }
}
