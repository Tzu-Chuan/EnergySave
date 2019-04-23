<%@ WebHandler Language="C#" Class="getMemberList" %>

using System;
using System.Web;
using System.Data;

public class getMemberList : IHttpHandler {
    Member_DB Member_Db = new Member_DB();
    public void ProcessRequest (HttpContext context) {
        try
        {
            string type = (context.Request["type"] != null) ? context.Request["type"].ToString() : "";
            string SearchStr = (context.Request["SearchStr"] != null) ? context.Request["SearchStr"].ToString() : "";
            string CurrentPage = (context.Request.Form["CurrentPage"] == null) ? "" : context.Request.Form["CurrentPage"].ToString().Trim();
            string sortMethod = (context.Request.Form["sortMethod"] == null) ? "" : context.Request.Form["sortMethod"].ToString().Trim();
            string sortName = (context.Request.Form["sortName"] == null) ? "" : context.Request.Form["sortName"].ToString().Trim();
            string sortStr = sortName + " " + sortMethod;

            int Paging = 10; //一頁幾筆
            int pageEnd = (int.Parse(CurrentPage) + 1) * Paging; // +1 是因為分頁從0開始算
            int pageStart = pageEnd - Paging + 1;

            string xmlStr = string.Empty;
            string xmlStr2 = string.Empty;

            Member_Db._KeyWord = SearchStr;
            DataSet ds = Member_Db.getMemberList(pageStart.ToString(), pageEnd.ToString(),sortStr);
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