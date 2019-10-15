<%@ WebHandler Language="C#" Class="getHistory_M" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Data;

public class getHistory_M : IHttpHandler,IRequiresSessionState {
    ReportCheck_DB rc_db = new ReportCheck_DB();
    Member_DB m_db = new Member_DB();
    public void ProcessRequest (HttpContext context) {
        try
        {
            if (LogInfo.mGuid == "")
            {
                context.Response.Write("reLogin");
                return;
            }

            string sday = (context.Request["sday"] != null) ? context.Request["sday"].ToString() : "";
            string eday = (context.Request["eday"] != null) ? context.Request["eday"].ToString() : "";
            string year = (context.Request["year"] != null) ? context.Request["year"].ToString() : "";
            string reporttype = (context.Request["reporttype"] != null) ? context.Request["reporttype"].ToString() : "";
            year = (year != "") ? (Int32.Parse(year) + 1911).ToString() : "";
            string month = (context.Request["month"] != null) ? context.Request["month"].ToString() : "";
            string CityStr = (context.Request["city"] != null) ? context.Request["city"].ToString() : "";
            string SearchStr = (context.Request["SearchStr"] != null) ? context.Request["SearchStr"].ToString() : "";
            string CurrentPage = (context.Request.Form["CurrentPage"] == null) ? "" : context.Request.Form["CurrentPage"].ToString().Trim();
            string stage = (context.Request["stage"] != null) ? context.Request["stage"].ToString() : "";

            int Paging = 10; //一頁幾筆
            int pageEnd = (int.Parse(CurrentPage) + 1) * Paging; // +1 是因為分頁從0開始算
            int pageStart = pageEnd - Paging + 1;

            string xmlStr = string.Empty;
            string xmlStr2 = string.Empty;
                
            string city = (LogInfo.competence != "SA") ? LogInfo.city : CityStr;

            rc_db._strKeyword = SearchStr;
            DataSet ds = rc_db.getHistoryMonth(pageStart.ToString(), pageEnd.ToString(), sday, eday, city, year.ToString(), month,reporttype,stage);
            DataTable dt = ds.Tables[1];

            xmlStr = "<total>" + ds.Tables[0].Rows[0]["total"].ToString() + "</total>";
            xmlStr += "<comp>" + LogInfo.competence+ "</comp>";
            if (dt.Rows.Count > 0)
            {
                xmlStr2 += "<dataList>";
                for(int i = 0; i < dt.Rows.Count; i++)
                {
                    xmlStr2 += "<data_item>";
                    xmlStr2 += "<RC_ReportGuid>" + dt.Rows[i]["RC_ReportGuid"].ToString() + "</RC_ReportGuid>";
                    xmlStr2 += "<enGuid>" + context.Server.UrlEncode(Common.Encrypt(dt.Rows[i]["RC_ReportGuid"].ToString())) + "</enGuid>";
                    xmlStr2 += "<RC_Year>" + dt.Rows[i]["RC_Year"].ToString() + "</RC_Year>";
                    xmlStr2 += "<RC_Month>" + dt.Rows[i]["RC_Month"].ToString() + "</RC_Month>";
                    xmlStr2 += "<MbName>" + dt.Rows[i]["MbName"].ToString() + "</MbName>";
                    xmlStr2 += "<RC_CheckDate>" + dt.Rows[i]["RC_CheckDate"].ToString() + "</RC_CheckDate>";
                    xmlStr2 += "<AdName>" + dt.Rows[i]["AdName"].ToString() + "</AdName>";
                    xmlStr2 += "<RC_ReportType>" + dt.Rows[i]["RC_ReportType"].ToString() + "</RC_ReportType>";
                    xmlStr2 += "<City>" + dt.Rows[i]["City"].ToString() + "</City>";//20190801新縣市名稱欄位
                    xmlStr2 += "</data_item>";
                }
                xmlStr2 += "</dataList>";
            }

            xmlStr = "<root>" + xmlStr + xmlStr2 + "</root>";
            context.Response.Write(xmlStr);
        }
        catch (Exception ex) { context.Response.Write("Error：" + ex.Message.Replace("'", "\"")); }
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}