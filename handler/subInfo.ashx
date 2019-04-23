<%@ WebHandler Language="C#" Class="subInfo" %>

using System;
using System.Web;
using System.Web.SessionState;

public class subInfo : IHttpHandler,IRequiresSessionState {
    ProjectInfo_DB p_db = new ProjectInfo_DB();
    Member_DB m_db = new Member_DB();
    public void ProcessRequest (HttpContext context) {
        try
        {
            if (LogInfo.mGuid == "")
            {
                context.Response.Write("reLogin");
                return;
            }

            string id = (context.Request["id"] != null) ? context.Request["id"].ToString() : "";

            m_db._M_ID = id;
            string ProjectID = m_db.getProgectGuidByPersonId();

            p_db._I_GUID = ProjectID;
            p_db._I_ModId = LogInfo.mGuid;
            p_db.submit_Info();


            context.Response.Write("succeed");
        }
        catch (Exception ex)
        {
            context.Response.Write("Error：" + ex.Message.Replace("'", "\""));
        }
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}