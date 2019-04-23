using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class handler_SendMailCheck : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //MailGun mg = new MailGun();
        MailUtil mg = new MailUtil();
        ReportMonth_DB rm_db = new ReportMonth_DB();
        try {
            string getReportGuid = string.IsNullOrEmpty(Request.Form["reportGuid"]) ? "" : Common.Decrypt(Request.Form["reportGuid"]).ToString().Trim();
            string getmailtype = string.IsNullOrEmpty(Request.Form["mailtype"]) ? "" : Request.Form["mailtype"].ToString().Trim(); //SnotO 季報不通過 KMnotOK 月報不通過
            string getmailbody = string.IsNullOrEmpty(Request.Form["mailbody"]) ? "" : Request.Form["mailbody"].ToString().Trim();
            string getseason = string.IsNullOrEmpty(Request.Form["season"]) ? "" : Request.Form["season"].ToString().Trim();//季
            string getyear = string.IsNullOrEmpty(Request.Form["yyyy"]) ? "" : Request.Form["yyyy"].ToString().Trim();//年
            string getmonth = string.IsNullOrEmpty(Request.Form["mm"]) ? "" : Request.Form["mm"].ToString().Trim();//月
            string mailTo = string.Empty;//收件者
            string mailCC = string.Empty;//CC 簽核主管自己
            string mailBCC = string.Empty;//BCC
            if (getReportGuid != "") {
                //撈出承辦人 & 主管MAIL
                DataSet ds = rm_db.selectMailByReportguid(getReportGuid);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    mailTo = ds.Tables[0].Rows[0]["M_Email"].ToString().Trim();
                }
                else {
                    mailTo = "";
                }
                if (ds.Tables[1].Rows.Count > 0)
                {
                    mailCC = ds.Tables[1].Rows[0]["M_Email"].ToString().Trim();
                }
                else {
                    mailCC = "";
                }
                mailBCC = "yhc@itri.org.tw";
                //測試先寫死
                //mailTo = "wcc@bestitmaster.com";
                //mailCC = "nicklai@bestitmaster.com";
                //mailBCC = "wang770418@gmail.com";
                //說明
                //MailUtil.cs  MailTo function
                //toMail: 收件者  *多人請以逗點隔開
                //ccMail: 副本 *多人請以逗點隔開
                //bccMail: 密件副本 *多人請以逗點隔開
                //subject: 主旨
                //body: 內容
                if (getmailtype == "MnotOK") {
                    //月報審核不通過
                    mg.MailTo(mailTo,mailCC,mailBCC, getyear + "年" + getmonth + "月月報審核不通過", "不通過原因：<br />" + getmailbody + "<br /><br /><br /> 此為系統自動寄發信件，請勿回信");
                }
                if (getmailtype == "SnotOK")
                {
                    //季報審核不通過
                    mg.MailTo(mailTo, mailCC, mailBCC, getyear + "年第" + getseason + "季季報審核不通過", "不通過原因：<br />" + getmailbody + "<br /><br /><br /> 此為系統自動寄發信件，請勿回信");
                }

                Response.Write("success");
            }
        }
        catch (Exception ex) { Response.Write("<script> alert();Error：" + ex.Message.Replace("'", "\"") +"</script>"); }


    }
}