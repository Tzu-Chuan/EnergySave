<%@ WebHandler Language="C#" Class="getProjectDay" %>

using System;
using System.Web;
using System.Data;

public class getProjectDay : IHttpHandler {
    ProjectDate_DB pd_db = new ProjectDate_DB();
    public void ProcessRequest (HttpContext context) {
        try
        {
            string type = (context.Request["type"] != null) ? context.Request["type"].ToString() : "";

            string xmlStr = string.Empty;

            pd_db._PD_Type = type;
            DataTable dt = pd_db.SelectList();
            xmlStr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
            xmlStr = "<root>" + xmlStr + "</root>";

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