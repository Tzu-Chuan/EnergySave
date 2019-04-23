<%@ WebHandler Language="C#" Class="Login" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Data;

public class Login : IHttpHandler,IRequiresSessionState {
    Account ac = new Account();
    Log_DB l_db = new Log_DB();
    public void ProcessRequest (HttpContext context) {
        try
        {
            string acc = (context.Request["acc"] != null) ? context.Request["acc"].ToString() : "";
            string pw = (context.Request["pw"] != null) ? context.Request["pw"].ToString() : "";
            string vCode = (context.Request["vCode"] != null) ? context.Request["vCode"].ToString() : "";

            //判斷驗証碼是否正確
            //if (context.Session["ValidateNumber"] != null && context.Session["ValidateNumber"] != "")
            //{
            //    if (context.Session["ValidateNumber"].ToString() != vCode.ToUpper())
            //    {
            //        context.Response.Write("CodeFailed");
            //        return;
            //    }
            //}

            AccountInfo accInfo = new Account().ExecLogon(acc, Common.sha1en(pw));
            if (accInfo != null)
            {
                context.Session["AccountInfo"] = accInfo;
                l_db._L_Person = LogInfo.mGuid;
                l_db._L_Type = "01";
                l_db._L_IP = Common.GetIP4Address();
                l_db.addLog();
                context.Response.Write("succeed");
            }
            else
            {
                context.Response.Write("Failed");
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