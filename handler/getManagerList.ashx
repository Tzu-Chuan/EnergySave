<%@ WebHandler Language="C#" Class="getManagerList" %>

using System;
using System.Web;
using System.Data;

public class getManagerList : IHttpHandler {
    Member_DB Member_Db = new Member_DB();
    public void ProcessRequest (HttpContext context) {
        try
        {
            string SearchStr = (context.Request["SearchStr"] != null) ? context.Request["SearchStr"].ToString() : "";
            string City = (context.Request["City"] != null) ? context.Request["City"].ToString() : "";
            string CurrentPage = (context.Request.Form["CurrentPage"] == null) ? "" : context.Request.Form["CurrentPage"].ToString().Trim();

            int Paging = 10; //一頁幾筆
            int pageEnd = (int.Parse(CurrentPage) + 1) * Paging; // +1 是因為分頁從0開始算
            int pageStart = pageEnd - Paging + 1;

            string xmlStr = string.Empty;
            string xmlStr2 = string.Empty;

            Member_Db._M_City = City;
            Member_Db._KeyWord = SearchStr;
            DataSet ds = Member_Db.getManagerList(pageStart.ToString(), pageEnd.ToString());
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