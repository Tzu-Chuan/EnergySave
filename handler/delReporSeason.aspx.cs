using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;

public partial class handler_delReporSeason : System.Web.UI.Page
{
    ReportSeasonV2_DB rs_db = new ReportSeasonV2_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        ///-----------------------------------------------------
        ///功    能: 季報 刪除草稿
        ///說明:
        /// * Request["strReportGuid"]: RM_ReportGuid
        /// 
        ///-----------------------------------------------------
        ///
        XmlDocument xDoc = new XmlDocument();
        try
        {
            string strguid = (Request["strReportGuid"] != null) ? Request["strReportGuid"].ToString().Trim() : "";
            string year = (Request["year"] != null) ? Request["year"].ToString().Trim() : "";
            string season = (Request["season"] != null) ? Request["season"].ToString().Trim() : "";
            string stage = (Request["stage"] != null) ? Request["stage"].ToString().Trim() : "";

            rs_db._RS_Guid = strguid;
            rs_db._RS_Year = year;
            rs_db._RS_Season = season;
            rs_db._RS_Stage = stage;
            rs_db.delReportSeasonNotCheck();

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