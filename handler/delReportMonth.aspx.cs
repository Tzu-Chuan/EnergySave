using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;

public partial class handler_delReportMonth : System.Web.UI.Page
{
    ReportMonth_DB rm_db = new ReportMonth_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        ///-----------------------------------------------------
        ///功    能: 月報 刪除草稿
        ///說明:
        /// * Request["strReportGuid"]: RM_ReportGuid
        /// 
        ///-----------------------------------------------------
        XmlDocument xDoc = new XmlDocument();
        try
        {
            string strguid = (Request["strReportGuid"] != null) ? Request["strReportGuid"].ToString().Trim() : "";

            rm_db._RM_ReportGuid = strguid;
            rm_db.delReportMonthNotCheck();
            
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