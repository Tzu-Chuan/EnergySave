<%@ WebHandler Language="C#" Class="getLogProjectInfo" %>

using System;
using System.Web;
using System.Data;

public class getLogProjectInfo : IHttpHandler {
    Log_DB l_db = new Log_DB();
    public void ProcessRequest (HttpContext context) {
        try
        {
            string sday = (context.Request["sday"] != null) ? context.Request["sday"].ToString() : "";
            string eday = (context.Request["eday"] != null) ? context.Request["eday"].ToString() : "";
            string SearchStr = (context.Request["SearchStr"] != null) ? context.Request["SearchStr"].ToString() : "";
            string CurrentPage = (context.Request.Form["CurrentPage"] == null) ? "" : context.Request.Form["CurrentPage"].ToString().Trim();

            int Paging = 10; //一頁幾筆
            int pageEnd = (int.Parse(CurrentPage) + 1) * Paging; // +1 是因為分頁從0開始算
            int pageStart = pageEnd - Paging + 1;

            string xmlStr = string.Empty;
            string xmlStr2 = string.Empty;
                
            l_db._KeyWord = SearchStr;
            DataSet ds = l_db.getProjectInfo(pageStart.ToString(), pageEnd.ToString(), sday, eday);
            DataTable dt = ds.Tables[1];

            xmlStr = "<total>" + ds.Tables[0].Rows[0]["total"].ToString() + "</total>";
            xmlStr2 = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
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