<%@ WebHandler Language="C#" Class="addMember" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Data;

public class addMember : IHttpHandler,IRequiresSessionState {
    Member_DB Member_Db = new Member_DB();
    Log_DB l_db = new Log_DB();
    public void ProcessRequest (HttpContext context) {
        try
        {
            if (LogInfo.mGuid == "")
            {
                context.Response.Write("reLogin");
                return;
            }

            string mode = (context.Request["mode"] != null) ? context.Request["mode"].ToString() : "";
            string id = (context.Request["id"] != null) ? context.Request["id"].ToString() : "";
            string M_Account = (context.Request["M_Account"] != null) ? context.Request["M_Account"].ToString() : "";
            string Old_Pwd = (context.Request["Old_Pwd"] != null) ? context.Request["Old_Pwd"].ToString() : "";
            string M_Pwd = (context.Request["M_Pwd"] != null) ? context.Request["M_Pwd"].ToString() : "";
            string M_Name = (context.Request["M_Name"] != null) ? context.Request["M_Name"].ToString() : "";
            string M_JobTitle = (context.Request["M_JobTitle"] != null) ? context.Request["M_JobTitle"].ToString() : "";
            string M_Tel = (context.Request["M_Tel"] != null) ? context.Request["M_Tel"].ToString() : "";
            string M_Ext = (context.Request["M_Ext"] != null) ? context.Request["M_Ext"].ToString() : "";
            string M_Fax = (context.Request["M_Fax"] != null) ? context.Request["M_Fax"].ToString() : "";
            string M_Phone = (context.Request["M_Phone"] != null) ? context.Request["M_Phone"].ToString() : "";
            string M_Email = (context.Request["M_Email"] != null) ? context.Request["M_Email"].ToString() : "";
            string M_Addr = (context.Request["M_Addr"] != null) ? context.Request["M_Addr"].ToString() : "";
            string M_City = (context.Request["M_City"] != null) ? context.Request["M_City"].ToString() : "";
            string M_Office = (context.Request["M_Office"] != null) ? context.Request["M_Office"].ToString() : "";
            string M_Competence = (context.Request["M_Competence"] != null) ? context.Request["M_Competence"].ToString() : "";
            string M_Manager_ID = (context.Request["M_Manager_ID"] != null) ? context.Request["M_Manager_ID"].ToString() : "";
            string aprEmail = (context.Request["aprEmail"] != null) ? context.Request["aprEmail"].ToString() : "";

            string gid = Guid.NewGuid().ToString("N");
            if (id != "")
            {
                Member_Db._M_ID = id;
                gid = Member_Db.getGuidById();
            }

            Member_Db._M_ID = id;
            Member_Db._M_Email = M_Email;
            DataTable dt = Member_Db.checkEmailRepeat();
            if (dt.Rows.Count > 0)
            {
                context.Response.Write("MailRepeat");
                return;
            }

            //設定帳密同email
            if (aprEmail == "true")
            {
                M_Account = M_Email;
                M_Pwd = M_Email;
            }

            switch (mode)
            {
                case "New":
                    Member_Db._M_Guid = gid;
                    Member_Db._M_Account = M_Account;
                    Member_Db._M_Pwd = Common.sha1en(M_Pwd);
                    Member_Db._M_Name = M_Name;
                    Member_Db._M_JobTitle = M_JobTitle;
                    Member_Db._M_Tel = M_Tel;
                    Member_Db._M_Ext = M_Ext;
                    Member_Db._M_Fax = M_Fax;
                    Member_Db._M_Phone = M_Phone;
                    Member_Db._M_Email = M_Email;
                    Member_Db._M_Addr = M_Addr;
                    Member_Db._M_City = M_City;
                    Member_Db._M_Office = M_Office;
                    Member_Db._M_Competence = M_Competence;
                    Member_Db._M_Manager_ID = M_Manager_ID;
                    Member_Db._M_CreateId = LogInfo.mGuid;
                    Member_Db._M_ModId = LogInfo.mGuid;
                    Member_Db.addMember();

                    //LOG
                    l_db._L_Type = "03";
                    l_db._L_Person = LogInfo.mGuid;
                    l_db._L_IP = Common.GetIP4Address();
                    l_db._L_ModItemGuid = gid;
                    l_db._L_Desc = "建立密碼";
                    l_db.addLog();
                    break;
                case "Mod":
                    Member_Db._M_ID = id;
                    Member_Db._M_Account = M_Account;
                    if (Old_Pwd != M_Pwd)
                    {
                        Member_Db._M_Pwd = Common.sha1en(M_Pwd);

                        //LOG
                        l_db._L_Type = "03";
                        l_db._L_Person = LogInfo.mGuid;
                        l_db._L_IP = Common.GetIP4Address();
                        l_db._L_ModItemGuid = gid;
                        l_db._L_Desc = "修改密碼";
                        l_db.addLog();
                    }
                    else
                        Member_Db._M_Pwd = M_Pwd;
                    Member_Db._M_Name = M_Name;
                    Member_Db._M_JobTitle = M_JobTitle;
                    Member_Db._M_Tel = M_Tel;
                    Member_Db._M_Ext = M_Ext;
                    Member_Db._M_Fax = M_Fax;
                    Member_Db._M_Phone = M_Phone;
                    Member_Db._M_Email = M_Email;
                    Member_Db._M_Addr = M_Addr;
                    Member_Db._M_City = M_City;
                    Member_Db._M_Office = M_Office;
                    Member_Db._M_Competence = M_Competence;
                    Member_Db._M_Manager_ID = M_Manager_ID;
                    Member_Db._M_ModId = LogInfo.mGuid;
                    Member_Db.modMember();
                    break;
                case "pFile":
                    Member_Db._M_ID = LogInfo.id;
                    Member_Db._M_Account = M_Account;
                    if (Old_Pwd != M_Pwd)
                    {
                        Member_Db._M_Pwd = Common.sha1en(M_Pwd);

                        //LOG
                        l_db._L_Type = "03";
                        l_db._L_Person = LogInfo.mGuid;
                        l_db._L_IP = Common.GetIP4Address();
                        l_db._L_ModItemGuid = gid;
                        l_db._L_Desc = "修改密碼";
                        l_db.addLog();
                    }
                    else
                        Member_Db._M_Pwd = M_Pwd;
                    Member_Db._M_Name = M_Name;
                    Member_Db._M_JobTitle = M_JobTitle;
                    Member_Db._M_Tel = M_Tel;
                    Member_Db._M_Ext = M_Ext;
                    Member_Db._M_Fax = M_Fax;
                    Member_Db._M_Phone = M_Phone;
                    Member_Db._M_Email = M_Email;
                    Member_Db._M_Addr = M_Addr;
                    Member_Db._M_City = M_City;
                    Member_Db._M_Office = M_Office;
                    Member_Db._M_ModId = LogInfo.mGuid;
                    Member_Db.modPersonalFile();
                    break;
            }
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