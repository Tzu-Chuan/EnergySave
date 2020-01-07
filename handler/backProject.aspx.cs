using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;

public partial class handler_backProject : System.Web.UI.Page
{
    ProjectInfo_DB pj_db = new ProjectInfo_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        /*
            功能： 計畫退回草稿
            傳入參數： 
         */
        XmlDocument xDoc = new XmlDocument();
        try
        {
            string aid = (Request["aid"] != null) ? Request["aid"].ToString().Trim() : "";

            pj_db._I_ID = aid;
            pj_db.backProject();

            xDoc.LoadXml("<?xml version='1.0' encoding='utf-8'?><root><Response>success</Response></root>");
        }
        catch (Exception ex)
        {
            xDoc = ExceptionUtil.GetExceptionDocument(ex);
        }
        Response.ContentType = System.Net.Mime.MediaTypeNames.Text.Xml;
        xDoc.Save(Response.Output);
    }
}
