using Aceoffix;
using System;

namespace Aceoffix7_Net.SimpleExcel
{
    public partial class Excel : System.Web.UI.Page
    {
        public AceoffixCtrl aceCtrl = new AceoffixCtrl();
        protected void Page_Load(object sender, EventArgs e)

        {
            aceCtrl.SaveFilePage = "SaveFile.aspx";
            aceCtrl.WebOpen("doc/test.xlsx", OpenModeType.xlsNormalEdit, "Tom");
        }
    }
}