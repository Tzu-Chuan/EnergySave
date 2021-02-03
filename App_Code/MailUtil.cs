using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Net.Mail;
using System.Net;

/// <summary>
/// MailUtil 的摘要描述
/// </summary>
public class MailUtil
{
    public void MailTo(string toMail, string subject, string body)
    {
        SendMail(toMail, "", "", subject, body);
    }
    public void MailTo(string toMail, string ccMail, string subject, string body)
    {
        SendMail(toMail, ccMail, "", subject, body);
    }
    public void MailTo(string toMail, string ccMail, string bccMail, string subject, string body)
    {
        SendMail(toMail, ccMail, bccMail, subject, body);
    }

    public void SendMail(string toMail, string ccMail, string bccMail, string subject, string body)
    {
        SmtpClient client = new SmtpClient();
        client.Host = ConfigurationManager.AppSettings["SmtpServer"];
        client.Port = 587;
        
        NetworkCredential basicCredential = new NetworkCredential(ConfigurationManager.AppSettings["SmtpAcc"], ConfigurationManager.AppSettings["SmtpPw"]);
        client.UseDefaultCredentials = false;
        client.Credentials = basicCredential;

        MailMessage message = new MailMessage();
        message.From = new MailAddress("energysave.ddns.net@ddns.net", "縣市共推住商節電行動計畫填報平台");
        //TO
        string[] toAddr = toMail.Split(',');
        for (int i = 0; i < toAddr.Length; i++)
        {
            if (!string.IsNullOrEmpty(toAddr[i]))
            {
                message.To.Add(new MailAddress(toAddr[i]));
            }
        }

        //CC
        string[] ccAddr = ccMail.Split(',');
        for (int i = 0; i < ccAddr.Length; i++)
        {
            if (!string.IsNullOrEmpty(ccAddr[i]))
            {
                message.CC.Add(new MailAddress(ccAddr[i]));
            }
        }

        //BCC
        string[] bccAddr = bccMail.Split(',');
        for (int i = 0; i < bccAddr.Length; i++)
        {
            if (!string.IsNullOrEmpty(bccAddr[i]))
            {
                message.Bcc.Add(new MailAddress(bccAddr[i]));
            }
        }

        message.Subject = subject;
        message.SubjectEncoding = System.Text.Encoding.UTF8;
        message.Body = body;
        message.IsBodyHtml = true;
        message.BodyEncoding = System.Text.Encoding.UTF8;

        if (ConfigurationManager.AppSettings["MailStatus"] == "true")
            client.Send(message);
    }
}