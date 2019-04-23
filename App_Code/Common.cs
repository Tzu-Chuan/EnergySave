using System;
using System.Collections.Generic;
using System.Web;
using System.Configuration;
using System.Text;
using System.Text.RegularExpressions;
using System.Data;
using System.Security.Cryptography;
using System.IO;
using System.Data.SqlClient;
//using FlexCel.Core;
using System.Net;
using System.Web.UI;

/// <summary>
/// Common 的摘要描述
/// </summary>
public class Common
{

    #region Get IPv4 Adress
    public static string GetIP4Address()
    {
        System.Web.HttpContext context = System.Web.HttpContext.Current;
        string sIPAddress = context.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];
        if (string.IsNullOrEmpty(sIPAddress))
        {
            string[] ipstr = context.Request.ServerVariables["REMOTE_ADDR"].Split(':');
            if (ipstr[0].Trim() != "")
                return context.Request.ServerVariables["REMOTE_ADDR"];
            else
                return "LOCAL-Name：" + Environment.MachineName;
        }
        else
        {
            string[] ipArray = sIPAddress.Split(new Char[] { ',' });
            return ipArray[0];
        }
    }
    #endregion

    #region 加解密
    /// <summary>
    /// 加密
    /// </summary>
    public static string Encrypt(string strSource)
    {
        //把字符串放到byte数组中  
        byte[] bytIn = System.Text.Encoding.Default.GetBytes(strSource);
        //建立加密对象的密钥和偏移量          
        byte[] iv = { 102, 16, 93, 156, 78, 4, 218, 32 };//定义偏移量  
        byte[] key = { 55, 103, 246, 79, 36, 99, 167, 3 };//定义密钥
        //实例DES加密类  
        DESCryptoServiceProvider mobjCryptoService = new DESCryptoServiceProvider();
        mobjCryptoService.Key = iv;
        mobjCryptoService.IV = key;
        ICryptoTransform encrypto = mobjCryptoService.CreateEncryptor();
        //实例MemoryStream流加密密文件  
        System.IO.MemoryStream ms = new System.IO.MemoryStream();
        CryptoStream cs = new CryptoStream(ms, encrypto, CryptoStreamMode.Write);
        cs.Write(bytIn, 0, bytIn.Length);
        cs.FlushFinalBlock();
        return System.Convert.ToBase64String(ms.ToArray());
    }

    /// <summary>
    /// 解密
    /// </summary>
    public static string Decrypt(string Source)
    {
        string str = "";
        try
        {
            //将解密字符串转换成字节数组  
            byte[] bytIn = System.Convert.FromBase64String(Source);
            //给出解密的密钥和偏移量，密钥和偏移量必须与加密时的密钥和偏移量相同  
            byte[] iv = { 102, 16, 93, 156, 78, 4, 218, 32 };//定义偏移量  
            byte[] key = { 55, 103, 246, 79, 36, 99, 167, 3 };//定义密钥  
            DESCryptoServiceProvider mobjCryptoService = new DESCryptoServiceProvider();
            mobjCryptoService.Key = iv;
            mobjCryptoService.IV = key;
            //实例流进行解密  
            System.IO.MemoryStream ms = new System.IO.MemoryStream(bytIn, 0, bytIn.Length);
            ICryptoTransform encrypto = mobjCryptoService.CreateDecryptor();
            CryptoStream cs = new CryptoStream(ms, encrypto, CryptoStreamMode.Read);
            StreamReader strd = new StreamReader(cs, Encoding.Default);
            str = strd.ReadToEnd();
        }
        catch
        {

        }
        return str;
    }

    /// <summary>
    /// SHA1加密
    /// </summary>
    public static string sha1en(string unCodeString)
    {
        string enCodeString;
        SHA1CryptoServiceProvider sha1en = new SHA1CryptoServiceProvider();
        enCodeString = BitConverter.ToString(sha1en.ComputeHash(UTF8Encoding.Default.GetBytes(unCodeString)), 4, 8);
        enCodeString = enCodeString.Replace("-", "");
        return enCodeString;
    }
    #endregion

    #region 匯出EXCEL
    //public static void excuteExcel(ExcelFile Xls, string ShowfileName)
    //{
    //    HttpContext.Current.Response.ContentType = "application/ms-excel";

    //    string fileName = HttpContext.Current.Server.UrlPathEncode(ShowfileName);
    //    string strContentDisposition = String.Format("{0}; filename=\"{1}\"", "attachment", fileName);
    //    HttpContext.Current.Response.AddHeader("Content-Disposition", strContentDisposition);
    //    MemoryStream ms = new MemoryStream();
    //    Xls.Save(ms, TFileFormats.Xls);
    //    ms.Position = 0;

    //    BinaryReader br = new BinaryReader(ms);
    //    BinaryWriter bw = new BinaryWriter(HttpContext.Current.Response.OutputStream);

    //    for (int i = 0; i < ms.Length; i++)
    //    {
    //        bw.Write(br.ReadByte());
    //    }
    //    bw.Close();
    //    ms.Close();

    //    bw = null;
    //    ms = null;

    //    HttpContext.Current.Response.OutputStream.Flush();
    //    HttpContext.Current.Response.OutputStream.Close();
    //    return;
    //}
    #endregion

    #region sqlInjection
    /// <summary>
    /// 檢查特殊字元
    /// </summary>
    /// <param name="checkValue">欲檢查的值</param>
    /// <returns></returns>
    public static bool CheckSQLInjection(string checkValue)
    {
        //「%27」:「'」(單引號)
        //「%2B」:「+」(加號)
        //「%3D」:「=」(等號)
        //「%7C」:「|」(｜)
        //「ALERT(」
        //「--」
        //「%2F*」:「/*」
        //「*%2F」:「*/」
        //「%26」:「&」
        //「%40」:「@」
        //「%25」:「%」
        //「%3B」:「;」
        //「%24」:「$」
        //「%26」:「*」
        //「%22」:「"」
        //「%2C」:「,」
        //「%2f」:「/」
        //「%5c」:「\」
        //「%22」:「"」
        //「%3C」:「<」
        //「%3E」:「>」


         
        if (checkValue.Length > 0 && (checkValue.ToUpper().IndexOf("%27") >= 0 || checkValue.ToUpper().IndexOf("%2B") >= 0
         || checkValue.ToUpper().IndexOf("'") >= 0) || checkValue.ToUpper().IndexOf("ALERT(") >= 0
         || checkValue.ToUpper().IndexOf("%3C") >= 0 || checkValue.ToUpper().IndexOf("%3E") >= 0
         || checkValue.ToUpper().IndexOf("%3D") >= 0 || checkValue.ToUpper().IndexOf("=") >= 0
         || checkValue.ToUpper().IndexOf("--") >= 0 || checkValue.ToUpper().IndexOf("%7C") >= 0
         || checkValue.ToUpper().IndexOf("%2F*") >= 0 || checkValue.ToUpper().IndexOf("*%2F") >= 0
         || checkValue.ToUpper().IndexOf("%26") >= 0
         || checkValue.ToUpper().IndexOf("%25") >= 0 || checkValue.ToUpper().IndexOf("%3B") >= 0
         || checkValue.ToUpper().IndexOf("%24") >= 0 || checkValue.ToUpper().IndexOf("*") >= 0
         || checkValue.ToUpper().IndexOf("%22") >= 0 || checkValue.ToUpper().IndexOf("%2C") >= 0
         || checkValue.ToUpper().IndexOf("%2F") >= 0 || checkValue.ToUpper().IndexOf("%5C") >= 0
         || checkValue.ToUpper().IndexOf("%40") >= 0)
        {
            return false;
        }
        else
        {
            return true;
        }
    }

    public static bool sqlInjection(string CheckString)
    {
        string SQL_injdata = @":|;|>|<|--|sp_|xp_|\|dir|cmd|^|(|)|+|$|'|copy|format|and|exec|insert|select|delete|update|count|*|%|chr|mid|master|truncate|char|declare";
        string[] SQL_inj = SQL_injdata.Split('|');
        bool returnBool = false;

        for (int i = 0; i < SQL_inj.Length; i++)
        {
            if (CheckString.IndexOf(SQL_inj[i]) != -1)
            {
                returnBool = true;
                break;
            }
            else
            {
                returnBool = false;
            }
        }

        return returnBool;
    }
    #endregion

    #region 檢查參數
    public void CheckParameters(System.Data.SqlClient.SqlCommand oCmd)
    {
        //檢查危險字元
        for (int i = 0; i < oCmd.Parameters.Count; i++)
        {
            if (!CheckSQLInjection(oCmd.Parameters[i].Value.ToString()))
            {
                //throw new Exception("危險字元");
                //導引至錯誤網頁
                System.Web.HttpContext.Current.Response.Redirect("Error.aspx?err=par");

            }
        }
    }
    #endregion

    #region 清除html...
    /// <summary>
    /// 輸入html後刪除html標簽...
    /// </summary>
    public static string NoHTML(string Htmlstring)
    {
        //删除脚本
        Htmlstring = Htmlstring.Replace("\r\n", "");
        Htmlstring = Regex.Replace(Htmlstring, @"<script.*?</script>", "", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"<style.*?</style>", "", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"<.*?>", "", RegexOptions.IgnoreCase);
        //删除HTML
        Htmlstring = Regex.Replace(Htmlstring, @"<(.[^>]*)>", "", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"([\r\n])[\s]+", "", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"-->", "", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"<!--.*", "", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(quot|#34);", "\"", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(amp|#38);", "&", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(lt|#60);", "<", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(gt|#62);", ">", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(nbsp|#160);", "", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(iexcl|#161);", "\xa1", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(cent|#162);", "\xa2", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(pound|#163);", "\xa3", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&(copy|#169);", "\xa9", RegexOptions.IgnoreCase);
        Htmlstring = Regex.Replace(Htmlstring, @"&#(\d+);", "", RegexOptions.IgnoreCase);
        Htmlstring = Htmlstring.Replace("<", "");
        Htmlstring = Htmlstring.Replace(">", "");
        Htmlstring = Htmlstring.Replace("\r\n", "");
        Htmlstring = HttpContext.Current.Server.HtmlEncode(Htmlstring).Trim();
        return Htmlstring;
    }
    #endregion

}


#region 會員登入使用
/// <summary>
/// MbrAccount 的摘要描述。
/// </summary>
public class Account
{
    public AccountInfo ExecLogon(string account,string pw)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@" select * from Member where M_Account=@M_Account and M_Pwd=@M_Pwd and M_Status='A' ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();
        oCmd.Parameters.AddWithValue("@M_Account", account);
        oCmd.Parameters.AddWithValue("@M_Pwd", pw);
        oda.Fill(dt);

        if (dt.Rows.Count > 0)
        {
            return new AccountInfo(dt);
        }
        else
        {
            return null;
        }
    }

}
/*-------------------------------------------------------------------------------------------------------------------------*/

/// <summary>
/// 會員的摘要描述。
/// </summary>
public class AccountInfo
{
    public AccountInfo(DataTable dr)
    {
        if (dr.Rows.Count > 0)
        {
            LogInfo.id = dr.Rows[0]["M_ID"].ToString();
            LogInfo.mGuid = dr.Rows[0]["M_Guid"].ToString();
            LogInfo.name = dr.Rows[0]["M_Name"].ToString();
            LogInfo.jobtitle = dr.Rows[0]["M_JobTitle"].ToString();
            LogInfo.tel = dr.Rows[0]["M_Tel"].ToString();
            LogInfo.ext = dr.Rows[0]["M_Ext"].ToString();
            LogInfo.fax = dr.Rows[0]["M_Fax"].ToString();
            LogInfo.phone = dr.Rows[0]["M_Phone"].ToString();
            LogInfo.email = dr.Rows[0]["M_Email"].ToString();
            LogInfo.addr = dr.Rows[0]["M_Addr"].ToString();
            LogInfo.city = dr.Rows[0]["M_City"].ToString();
            LogInfo.office = dr.Rows[0]["M_Office"].ToString();
            LogInfo.competence = dr.Rows[0]["M_Competence"].ToString();
            LogInfo.manager = dr.Rows[0]["M_Manager_ID"].ToString();
        }
    }
}
#endregion


#region JavaScript Alert
public class JavaScript
{
    /// <summary>
    /// AlertMessage
    /// </summary>
    public static void AlertMessage(System.Web.UI.Page objPage, string strMessage)
    {
        strMessage = strMessage.Replace("\r\n", "\\r");
        StringBuilder sb = new StringBuilder();
        sb.AppendFormat(@"<Script language=""javascript"" type=""text/javascript"">");
        sb.AppendFormat(@"alert(""{0}"");", strMessage);
        sb.AppendFormat(@"</Script>");

        //objPage.ClientScript.RegisterClientScriptBlock(objPage.GetType(), "", strJS, false);
        objPage.ClientScript.RegisterStartupScript(objPage.GetType(), "", sb.ToString(), false);
    }

    /// <summary>
    /// AlertMessageClose
    /// </summary>
    public static void AlertMessageClose(System.Web.UI.Page objPage, string strMessage)
    {
        string strJS = "";
        strMessage = strMessage.Replace("\r\n", "\\r");
        strJS = @"<Script language='javascript' type='text/javascript' >";
        strJS += "alert('" + strMessage + "');";
        strJS += "window.close();";
        strJS += "</Script>";
        //objPage.ClientScript.RegisterClientScriptBlock(objPage.GetType(), "", strJS, false);
        objPage.ClientScript.RegisterStartupScript(objPage.GetType(), "", strJS, false);
    }

    /// <summary>
    /// AlertMessageRedirect
    /// </summary>
    public static void AlertMessageRedirect(System.Web.UI.Page objPage, string strMessage, string strRedirectPage)
    {
        AlertMessageRedirect(objPage, strMessage, strRedirectPage, false);
    }

    public static void AlertMessageRedirect(System.Web.UI.Page objPage, string strMessage, string strRedirectPage, bool IsDisplayData)
    {
        string strJS = "";
        strMessage = strMessage.Replace("\r\n", "\\r");
        strJS = @"<Script language='javascript' type='text/javascript'>";
        strJS += "alert('" + strMessage + "');";
        strJS += "window.location ='" + strRedirectPage + "'; ";
        strJS += "</Script>";

        if (IsDisplayData)
            objPage.ClientScript.RegisterStartupScript(objPage.GetType(), "", strJS, false);
        else
            objPage.ClientScript.RegisterClientScriptBlock(objPage.GetType(), "", strJS, false);
    }
}
#endregion

#region 抓網頁圖片
public class CaptureURL
{
    public string Capture(string url)
    {
        try
        {
            string strHTML = string.Empty;

            if (url.IndexOf("https://") < 0)
            {
                string E = System.IO.Path.GetExtension(url);

                if (!E.Trim().ToLower().Equals(".html") && !string.IsNullOrEmpty(E.Trim()))
                {
                    string param = "hl=zh-CN&newwindow=1";

                    byte[] bs = System.Text.Encoding.ASCII.GetBytes(param);

                    HttpWebRequest req = (HttpWebRequest)HttpWebRequest.Create(url);

                    req.Method = "POST";

                    req.ContentType = "application/x-www-form-urlencoded";

                    req.ContentLength = bs.Length;


                    using (Stream reqStream = req.GetRequestStream())
                    {
                        reqStream.Write(bs, 0, bs.Length);
                    }

                    using (WebResponse wr = req.GetResponse())
                    {
                        //在這裡對接收到的頁面內容進行處理

                        using (Stream myStream = wr.GetResponseStream())
                        {
                            using (StreamReader myStreamReader = new StreamReader(myStream, System.Text.Encoding.UTF8))
                            {
                                strHTML = myStreamReader.ReadToEnd();
                            }
                        }
                    }
                }
                else
                {
                    Uri myUri = new Uri(url);

                    // Create a 'HttpWebRequest' object for the specified url. 

                    HttpWebRequest myHttpWebRequest = (HttpWebRequest)WebRequest.Create(myUri);

                    // Set the user agent as if we were a web browser

                    myHttpWebRequest.UserAgent = @"Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.0.4) Gecko/20060508 Firefox/1.5.0.4";

                    HttpWebResponse myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();

                    var stream = myHttpWebResponse.GetResponseStream();

                    var reader = new StreamReader(stream);

                    var html = reader.ReadToEnd();

                    // Release resources of response object.

                    myHttpWebResponse.Close();

                    return html;
                }
            }
            else
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

                //request.Method = "HEAD";

                //request.AllowAutoRedirect = false;

                request.Credentials = CredentialCache.DefaultCredentials;

                // Ignore Certificate validation failures (aka untrusted certificate + certificate chains)

                ServicePointManager.ServerCertificateValidationCallback = ((sender, certificate, chain, sslPolicyErrors) => true);

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                Stream resStream = response.GetResponseStream();

                StreamReader reader = new StreamReader(resStream, System.Text.Encoding.UTF8);

                strHTML = reader.ReadToEnd();
            }
            return strHTML;
        }
        catch (Exception ex) { throw ex; }
    }
}
#endregion
