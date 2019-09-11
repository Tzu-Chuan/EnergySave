<%@ WebHandler Language="C#" Class="getHistory_S" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Data;

public class getHistory_S : IHttpHandler,IRequiresSessionState {
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
            year = (year != "") ? (Int32.Parse(year) + 1911).ToString() : "";
            string season = (context.Request["season"] != null) ? context.Request["season"].ToString() : "";
            string stage = (context.Request["stage"] != null) ? context.Request["stage"].ToString() : "";
            string SearchStr = (context.Request["SearchStr"] != null) ? context.Request["SearchStr"].ToString() : "";
            string CityStr = (context.Request["city"] != null) ? context.Request["city"].ToString() : "";
            string CurrentPage = (context.Request.Form["CurrentPage"] == null) ? "" : context.Request.Form["CurrentPage"].ToString().Trim();

            int Paging = 10; //一頁幾筆
            int pageEnd = (int.Parse(CurrentPage) + 1) * Paging; // +1 是因為分頁從0開始算
            int pageStart = pageEnd - Paging + 1;

            string xmlStr = string.Empty;
            string xmlStr2 = string.Empty;

            string city = (LogInfo.competence != "SA") ? LogInfo.city : CityStr;


            rc_db._strKeyword = SearchStr;
            DataSet ds = rc_db.getHistorySeason(pageStart.ToString(), pageEnd.ToString(), sday, eday, city, year.ToString(), season, stage);
            DataTable dt = ds.Tables[1];

            xmlStr = "<total>" + ds.Tables[0].Rows[0]["total"].ToString() + "</total>";
            xmlStr += "<comp>" + LogInfo.competence+ "</comp>";
            //xmlStr2 = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
            if (dt.Rows.Count > 0)
            {
                xmlStr2 += "<dataList>";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    xmlStr2 += "<data_item>";
                    xmlStr2 += "<RS_ID>" + dt.Rows[i]["RS_ID"].ToString() + "</RS_ID>";
                    xmlStr2 += "<RC_ReportGuid>" + dt.Rows[i]["RC_ReportGuid"].ToString() + "</RC_ReportGuid>";
                    xmlStr2 += "<RC_Year>" + dt.Rows[i]["RC_Year"].ToString() + "</RC_Year>";
                    xmlStr2 += "<RC_Season>" + dt.Rows[i]["RC_Season"].ToString() + "</RC_Season>";
                    xmlStr2 += "<RC_Stage>" + dt.Rows[i]["RC_Stage"].ToString() + "</RC_Stage>";
                    xmlStr2 += "<MbName>" + dt.Rows[i]["MbName"].ToString() + "</MbName>";
                    xmlStr2 += "<RC_CheckDate>" + dt.Rows[i]["RC_CheckDate"].ToString() + "</RC_CheckDate>";
                    xmlStr2 += "<AdName>" + dt.Rows[i]["AdName"].ToString() + "</AdName>";
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