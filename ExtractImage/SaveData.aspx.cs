
using Aceoffix.Word;
using System;

namespace Aceoffix7_Net.ExtractImage
{
    public partial class SaveData : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            WordDocumentReader wordDoc = new WordDocumentReader();
            DataRegionReader dataRegion1 = wordDoc.OpenDataRegion("ACE_image");
            //将提取的图片保存到服务器上，图片的名称为:a.jpg
            dataRegion1.OpenShape(1).SaveAsJPG(Server.MapPath("doc/logo.png"));
            wordDoc.CustomSaveResult = "The save was successful. The image path is:" + Server.MapPath("doc/logo.png");
            wordDoc.Close();
        }
    }
}