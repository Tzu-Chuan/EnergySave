using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;

public partial class handler_CancelReport : System.Web.UI.Page
{
    OtherManage_DB om_db = new OtherManage_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        ///-----------------------------------------------------
        ///功    能: 月/季報送審抽單
        ///說明:
        /// * Request["id"]: RC_Guid
        /// * Request["type"]: 報表類別
        ///-----------------------------------------------------
        XmlDocument xDoc = new XmlDocument();
        try
        {
            string id = (Request["id"] != null) ? Request["id"].ToString().Trim() : "";
            string type = (Request["type"] != null) ? Request["type"].ToString().Trim() : "";
            
            if (type == "Month")
                om_db.CancelMonthReport(id);
            else
                om_db.CancelSeasonReport(id);

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