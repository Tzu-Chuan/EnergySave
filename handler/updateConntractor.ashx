<%@ WebHandler Language="C#" Class="updateConntractor" %>

using System;
using System.Web;
using System.Data;
using System.Web.SessionState;

public class updateConntractor : IHttpHandler,IRequiresSessionState {
    Member_DB m_db = new Member_DB();
    Log_DB l_db = new Log_DB();
    string name = string.Empty;
    string city = string.Empty;
    public void ProcessRequest (HttpContext context) {
        try
        {
            string OrgID = (context.Request["orgid"] != null) ? context.Request["orgid"].ToString() : "";
            string NewID = (context.Request["newid"] != null) ? context.Request["newid"].ToString() : "";

            m_db.changeConntractor(NewID, OrgID);

            l_db._L_Type = "11";
            l_db._L_Person = LogInfo.mGuid;
            l_db._L_IP = Common.GetIP4Address();
            l_db._L_Desc = getPersonName(OrgID, "City") + " " + getPersonName(OrgID, "M_Name") + " 更換為 " + getPersonName(NewID, "M_Name");
            l_db.addLog();

            context.Response.Write("succeed");
        }
        catch (Exception ex) { context.Response.Write("Error：" + ex.Message.Replace("'", "\"")); }
    }

    private string getPersonName(string mguid, string ColumnName)
    {
        string str = string.Empty;
        m_db._M_Guid = mguid;
        DataTable dt = m_db.getPersonByGuid();
        if (dt.Rows.Count > 0)
            str = dt.Rows[0][ColumnName].ToString();
        return str;
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}