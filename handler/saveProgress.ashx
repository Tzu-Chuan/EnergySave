<%@ WebHandler Language="C#" Class="saveProgress" %>

using System;
using System.Web;
using System.Web.SessionState;

public class saveProgress : IHttpHandler,IRequiresSessionState {
    Log_DB l_db = new Log_DB();
    Member_DB m_db = new Member_DB();
    CheckPoint_DB cp_db = new CheckPoint_DB();
    public void ProcessRequest (HttpContext context) {
        try
        {
            if (LogInfo.mGuid == "")
            {
                context.Response.Write("reLogin");
                return;
            }

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

            bool logstatus = false;
            if (wr_pid != null)
            {
                logstatus = true;
                //工作比重
                for (int i = 0; i < wr_pid.Length; i++)
                {
                    cp_db._P_Guid = wr_pid[i];
                    cp_db._P_WorkRatio = wr_value[i];
                    cp_db._CP_ModId = LogInfo.mGuid;
                    cp_db.setWorkRatio();
                }
            }

            if (pv_pid != null)
            {
                logstatus = true;
                //累計預定進度
                for (int i = 0; i < pv_pid.Length; i++)
                {
                    cp_db._CP_Guid = pv_pid[i];
                    cp_db._CP_Process = cp_process[i];
                    cp_db._CP_ModId = LogInfo.mGuid;
                    cp_db.setCP_Process();
                }
            }

            if (logstatus == true)
            {
                //Log
                l_db._L_Type = "07";
                l_db._L_Person = LogInfo.mGuid;
                l_db._L_IP = Common.GetIP4Address();
                l_db._L_ModItemGuid = ProjectId;
                l_db._L_Desc = "預定工作進度";
                l_db.addLog();
            }

            context.Response.Write("<script type='text/JavaScript'>parent.feedback('succeed');</script>");
        }
        catch (Exception ex)
        {
            context.Response.Write("<script type='text/JavaScript'>parent.feedback('Error：" + ex.Message.Replace("'", "\"") + "');</script>");
        }
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}