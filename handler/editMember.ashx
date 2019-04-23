<%@ WebHandler Language="C#" Class="editMember" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Data;

public class editMember : IHttpHandler,IRequiresSessionState {
    Member_DB Member_Db = new Member_DB();
    public void ProcessRequest(HttpContext context) {
        try
        {
            string id = (context.Request["id"] != null) ? context.Request["id"].ToString() : "";
            string mode = (context.Request["mode"] != null) ? context.Request["mode"].ToString() : "";

            id = (mode == "pFile") ? LogInfo.id : id;
            switch (mode)
            {
                case "Edit":
                case "pFile":
                    string xmlStr = string.Empty;
                    Member_Db._M_ID = id;
                    DataTable dt = Member_Db.getMemberById();
                    xmlStr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                    context.Response.Write(xmlStr);
                    break;
                case "Delete":
                    Member_Db._M_ID = id;
                    Member_Db._M_ModId = LogInfo.mGuid;
                    Member_Db.DeleteMember();
                    break;
            }
        }
        catch (Exception ex) { context.Response.Write("Error：" + ex.Message.Replace("'", "\"")); }
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}