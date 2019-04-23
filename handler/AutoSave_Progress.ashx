<%@ WebHandler Language="C#" Class="AutoSave_Progress" %>

using System;
using System.Web;
using System.Web.SessionState;

public class AutoSave_Progress : IHttpHandler,IRequiresSessionState {
    Member_DB m_db = new Member_DB();
    CheckPoint_DB cp_db = new CheckPoint_DB();
    public void ProcessRequest (HttpContext context) {
        try
        {
            string LoginPerson = "SessionTimeout-AutoSave";
            if (LogInfo.mGuid != "")
                LoginPerson = LogInfo.mGuid;
            
            string mid = string.IsNullOrEmpty(context.Request.Form["mid"]) ? "" : context.Request.Form["mid"].ToString().Trim();
            //get Project ID
            m_db._M_ID = mid;
            string ProjectId = m_db.getProgectGuidByPersonId();

            string[] wr_value = null;
            string[] wr_pid = null;
            string[] cp_process = null;
            string[] pv_pid = null;
            wr_value = (context.Request["wr_value"] != null) ? context.Request["wr_value"].ToString().Split(',') : null;
            wr_pid = (context.Request["wr_pid"] != null) ? context.Request["wr_pid"].ToString().Split(',') : null;
            cp_process = (context.Request["cp_process"] != null) ? context.Request["cp_process"].ToString().Split(',') : null;
            pv_pid = (context.Request["pv_pid"] != null) ? context.Request["pv_pid"].ToString().Split(',') : null;
                
            if (wr_pid != null)
            {
                //工作比重
                for (int i = 0; i < wr_pid.Length; i++)
                {
                    cp_db._P_Guid = wr_pid[i];
                    cp_db._P_WorkRatio = wr_value[i];
                    cp_db._CP_ModId = LoginPerson;
                    cp_db.setWorkRatio();
                }
            }

            if (pv_pid != null)
            {
                //累計預定進度
                for (int i = 0; i < pv_pid.Length; i++)
                {
                    cp_db._CP_Guid = pv_pid[i];
                    cp_db._CP_Process = cp_process[i];
                    cp_db._CP_ModId = LoginPerson;
                    cp_db.setCP_Process();
                }
            }

            context.Response.Write("<script type='text/JavaScript'>parent.autofeedback('succeed');</script>");
        }
        catch (Exception ex)
        {
            context.Response.Write("<script type='text/JavaScript'>parent.autofeedback('Error：" + ex.Message.Replace("'", "\"") + "');</script>");
        }
    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}