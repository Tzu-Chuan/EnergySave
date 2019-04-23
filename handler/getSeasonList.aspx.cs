using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Xml;

public partial class handler_getSeasonList : System.Web.UI.Page
{
    ReportSeasonV2_DB rs_db = new ReportSeasonV2_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        ///-----------------------------------------------------
        ///功    能: 查詢季報列表
        ///說明:
        /// * Request["PageNo"]:欲顯示的頁碼, 由零開始
        /// * Request["PageSize"]:每頁顯示的資料筆數, 未指定預設10
        ///-----------------------------------------------------
        XmlDocument xDoc = new XmlDocument();
        try
        {
            string PageNo = (Request["PageNo"] != null) ? Request["PageNo"].ToString().Trim() : "";
            int PageSize = (Request["PageSize"] != null) ? int.Parse(Request["PageSize"].ToString().Trim()) : 10;
            string year = (Request["year"] != null) ? Request["year"].ToString().Trim() : "";
            string season = (Request["season"] != null) ?Request["season"].ToString().Trim() : "";
            string stage = (Request["stage"] != null) ? Request["stage"].ToString().Trim() : "";

            int pageEnd = (int.Parse(PageNo) + 1) * PageSize;
            int pageStart = pageEnd - PageSize + 1;

            rs_db._RS_Year = year;
            rs_db._RS_Season = season;
            rs_db._RS_Stage = stage;
            DataSet dt = rs_db.getSeasonList(LogInfo.mGuid, pageStart.ToString(), pageEnd.ToString());
            string xmlstr = string.Empty;
            string totalxml = "<total>" + dt.Tables[0].Rows[0]["total"].ToString() + "</total>";
            xmlstr = DataTableToXml.ConvertDatatableToXML(dt.Tables[1], "dataList", "data_item");
            xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + totalxml + xmlstr + "</root>";
            xDoc.LoadXml(xmlstr);
        }
        catch (Exception ex)
        {
            xDoc = ExceptionUtil.GetExceptionDocument(ex);
        }
        Response.ContentType = System.Net.Mime.MediaTypeNames.Text.Xml;
        xDoc.Save(Response.Output);
    }
}