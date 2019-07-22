using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace UIP_Power_BI
{
    public partial class Ribbon_UIP_BI
    {
        private void Ribbon_UIP_BI_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Add_New_Trade_Button_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Application.Run("NewSheetButton");
        }

        private void View_Settings_Button_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Application.Run("ViewSettings_Sub");
        }

        private void View_Index_Button_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Application.Run("ViewIndex_Sub");
        }
    }
}
