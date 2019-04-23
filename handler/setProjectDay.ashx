<%@ WebHandler Language="C#" Class="setProjectDay" %>

using System;
using System.Web;
using System.Web.SessionState;

public class setProjectDay : IHttpHandler,IRequiresSessionState {
    ProjectDate_DB pd_db = new ProjectDate_DB();
    public void ProcessRequest (HttpContext context) {
        try
        {
            if (LogInfo.mGuid == "")
            {
                context.Response.Write("reLogin");
                return;
            }

            string type = (context.Request["type"] != null) ? context.Request["type"].ToString() : "";
            string sday = (context.Request["sday"] != null) ? context.Request["sday"].ToString() : "";
            string eday = (context.Request["eday"] != null) ? context.Request["eday"].ToString() : "";

            pd_db._PD_Type = type;
            if (sday != "")
                pd_db._PD_StartDate = DateTime.Parse(sday);
            if (eday != "")
                pd_db._PD_EndDate = DateTime.Parse(eday);
            pd_db._PD_ModId = LogInfo.mGuid;
            pd_db.setData();

            context.Response.Write("succeed");
        }
        catch (Exception ex) { context.Response.Write("Error：" + ex.Message.Replace("'", "\"")); }
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}