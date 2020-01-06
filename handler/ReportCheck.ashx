<%@ WebHandler Language="C#" Class="ReportCheck" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Data;

public class ReportCheck : IHttpHandler, IRequiresSessionState
{
    ReportCheck_DB rc = new ReportCheck_DB();
    Log_DB logdb = new Log_DB();
    MailUtil mg = new MailUtil();
    ReportMonth_DB rm_db = new ReportMonth_DB();
    public void ProcessRequest (HttpContext context) {
        try
        {
            if (LogInfo.mGuid == "")
            {
                context.Response.Write("timeout");
                return;
            }
            string gcRGuid = string.IsNullOrEmpty(context.Request.Form["str_reportGuid"]) ? "" : context.Request.Form["str_reportGuid"].ToString().Trim();
            string gcChkType = string.IsNullOrEmpty(context.Request.Form["str_checkType"]) ? "" : context.Request.Form["str_checkType"].ToString().Trim();
            string gcReportType = string.IsNullOrEmpty(context.Request.Form["str_ReportType"]) ? "" : context.Request.Form["str_ReportType"].ToString().Trim();
            string strchk = "";
            string getseason = string.IsNullOrEmpty(context.Request.Form["season"]) ? "" : context.Request.Form["season"].ToString().Trim();//季
            string getyear = string.IsNullOrEmpty(context.Request.Form["yyyy"]) ? "" : context.Request.Form["yyyy"].ToString().Trim();//年
            string getmonth = string.IsNullOrEmpty(context.Request.Form["mm"]) ? "" : context.Request.Form["mm"].ToString().Trim();//月
            string mailTo = string.Empty;//收件者
            string mailBCC = string.Empty;//BCC
            rc._strCheckType = gcChkType;
            rc._RC_ReportGuid = gcRGuid;
            rc._RC_Boss = LogInfo.mGuid;
            rc.ReportCheck();
            if (gcChkType=="Y") {
                strchk = "通過";
            }
            if (gcChkType=="N") {
                strchk = "不通過";
            }
            //紀錄LOG
            logdb._L_Person = LogInfo.mGuid;// "登入者Guid"
            logdb._L_IP = Common.GetIP4Address();
            logdb._L_ModItemGuid = gcRGuid;  //"該季報ReportGuid"
            if (gcReportType=="01" || gcReportType=="03") {
                logdb._L_Type = "06";
                logdb._L_Desc = strchk;
            }
            if (gcReportType=="02") {
                logdb._L_Type = "10";
                logdb._L_Desc = strchk;
            }
            logdb.addLog();

            //更新累計資料
            if (gcReportType=="01") {
                rc.updateReportMonthNum();
            }
            //審核通過發信
            //if (gcChkType=="Y") {
            //    //撈出承辦人 & 主管MAIL
            //    DataSet ds = rm_db.selectMailByReportguid(gcRGuid);
            //    if (ds.Tables[0].Rows.Count > 0)
            //    {
            //        mailTo = ds.Tables[0].Rows[0]["M_Email"].ToString().Trim();
            //    }
            //    else {
            //        mailTo = "";
            //    }
            //    mailBCC = "yhc@itri.org.tw";
            //    //測試先寫死
            //    //mailTo = "wcc@bestitmaster.com";
            //    //mailBCC = "wang770418@gmail.com";
            //    if (gcReportType == "01") {
            //        //月報審核通過
            //        mg.MailTo(mailTo,"",mailBCC, getyear + "年" + getmonth + "月月報審核通過", "<br /><br /><br />此為系統自動寄發信件，請勿回信");
            //    }
            //    if (gcReportType == "02")
            //    {
            //        //季報審核通過
            //        mg.MailTo(mailTo,"",mailBCC, getyear + "年第" + getseason + "季季報審核通過", "<br /><br /><br />此為系統自動寄發信件，請勿回信");
            //    }
            //}
            context.Response.Write("success");
        }
        catch (Exception ex) {
            context.Response.Write("Error：" + ex.Message.Replace("'", "\""));
        }
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}