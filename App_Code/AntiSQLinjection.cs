using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Text.RegularExpressions;
using System.Collections.Specialized;

/// <summary>
/// AntiSQLinjection 的摘要描述
/// </summary>
public class AntiSQLinjection : IHttpModule
{
    public void Dispose()
    {

    }

    public void Init(HttpApplication context)
    {
        //在此方法中调用一个委托

        //context_AcquireRequestState为委托的方法
        context.BeginRequest += new EventHandler(context_AcquireRequestState);
    }


    /// <summary>
    /// 委托中的方法

    /// </summary>
    /// <param ></param>
    /// <returns></returns>

    private void context_AcquireRequestState(object sender, EventArgs e)
    {

        //一定要将sender对象的Context转化为HttpContext的对象

        //其中包含请求的基本对象，如request,response等
        HttpContext context = ((HttpApplication)sender).Context;

        //错误处理的页面，里面可以是一个温馨提示
        string errorPage = "~/errorPage.aspx";
        //string errorPage = "~/msgpage.aspx";
        string keys = "";       //保存传参过来的键值 
        string values = "";    //保存传参过来的值
        //我们的请求可以是多种形式的，如表单提交，url传值等

        //我们就要对各种情况分类处理

        //如果是url传参传过来的情况
        if (context.Request.QueryString != null)
        {
            for (int i = 0; i < context.Request.QueryString.Count; i++)
            {
                //得到键值
                keys = context.Request.QueryString.Keys[i];
                //得到值
                values = context.Server.UrlDecode(context.Request.QueryString[keys]);
                //如果有非法字符串，则跳转到错误提示页面
                if (!this.ProcessSqlStrGET(values))
                {
                    context.Application["ErrorMsg"] = "請勿輸入非法字元";
                    context.Response.Redirect(errorPage);
                    context.Response.End();
                    break;
                }
            }
        }
        //如果是表单提交过来打值
        if (context.Request.Form != null)
        {
            for (int i = 0; i < context.Request.Form.Count; i++)
            {
                keys = context.Request.Form.Keys[i];
                values = context.Server.HtmlDecode(context.Request.Form[i]);
                if (keys == "__VIEWSTATE") continue;
                if (keys == "__EVENTTARGET") continue;
                if (keys == "__EVENTARGUMENT") continue;
                if (keys == "__LASTFOCUS") continue;
                if (keys == "__VIEWSTATEGENERATOR") continue;
                if (keys == "__EVENTVALIDATION") continue;
                if (keys == null) continue;//reportviewer 會有null
                if (keys.ToLower().IndexOf("desc") !=-1) continue;//排除tineymce插入圖片的路徑被阻擋的問題("=")

                //如果有非法字符串，则跳转到错误提示页面
                if (!this.ProcessSqlStrPOST(values))
                {
                    if (System.IO.Path.GetExtension(context.Request.Url.LocalPath.ToString()).Equals(".aspx"))
                    {
                        context.Response.Redirect(errorPage);
                        context.Response.End();
                        break;
                    }
                    else
                    {
                        throw new Exception("請勿輸入非法字元");
                    }
                }
            }
        }
    }

    /// <summary>
    /// 截取字符串的方法
    /// </summary>
    /// <param name="str"></param>
    /// <returns></returns>
    private bool ProcessSqlStrGET(string str)
    {
        bool bResult = true;
        try
        {
            str = Regex.Replace(str, "[\\s]{1,}", "");    //two or more spaces
            str = Regex.Replace(str, "(<[b|B][r|R]/*>)+|(<[p|P](.|\\n)*?>)", "\n");    //<br>

            string[] UnSafeArray = new string[28];
            UnSafeArray[0] = "'";
            UnSafeArray[1] = " xp_cmdshell ";
            UnSafeArray[2] = " declare ";
            UnSafeArray[3] = " netlocalgroupadministrators ";
            UnSafeArray[4] = " delete ";
            UnSafeArray[5] = " truncate ";
            UnSafeArray[6] = " netuser ";
            UnSafeArray[7] = "/add";
            UnSafeArray[8] = " drop ";
            UnSafeArray[9] = " update ";
            UnSafeArray[10] = " select ";
            UnSafeArray[11] = " union ";
            UnSafeArray[12] = " exec ";
            UnSafeArray[13] = " create ";
            UnSafeArray[14] = " insertinto ";
            UnSafeArray[15] = "sp_";
            UnSafeArray[16] = " exec ";
            UnSafeArray[17] = " create ";
            UnSafeArray[18] = " masterdbo ";
            UnSafeArray[19] = "sp_";
            UnSafeArray[20] = ";--";
            UnSafeArray[21] = "1=";
            UnSafeArray[22] = " and ";
            UnSafeArray[23] = " alert";
            UnSafeArray[24] = "\"";
            UnSafeArray[25] = "--";
            UnSafeArray[26] = "||";
            UnSafeArray[27] = "eval";
            foreach (string strValue in UnSafeArray)
            {
                if (str.ToLower().IndexOf(strValue) > -1)
                {
                    bResult = false;
                    break;
                }
            }

        }
        catch
        {
            bResult = false;
        }
        return bResult;
    }



    private bool ProcessSqlStrPOST(string str)
    {
        bool bResult = true;
        str = Regex.Replace(str, "[\\s]{1,}", "");    //two or more spaces
        str = Regex.Replace(str, "(<[b|B][r|R]/*>)+|(<[p|P](.|\\n)*?>)", "\n");    //<br>

        string[] UnSafeArray = new string[28];
        UnSafeArray[0] = "'";
        UnSafeArray[1] = " xp_cmdshell ";
        UnSafeArray[2] = " declare ";
        UnSafeArray[3] = " netlocalgroupadministrators ";
        UnSafeArray[4] = " delete ";
        UnSafeArray[5] = " truncate ";
        UnSafeArray[6] = " netuser ";
        UnSafeArray[7] = "/add";
        UnSafeArray[8] = " drop ";
        UnSafeArray[9] = " update ";
        UnSafeArray[10] = " select ";
        UnSafeArray[11] = " union ";
        UnSafeArray[12] = " exec ";
        UnSafeArray[13] = " create ";
        UnSafeArray[14] = " insertinto ";
        UnSafeArray[15] = "sp_";
        UnSafeArray[16] = " exec ";
        UnSafeArray[17] = " create ";
        UnSafeArray[18] = " masterdbo ";
        UnSafeArray[19] = "sp_";
        UnSafeArray[20] = ";--";
        UnSafeArray[21] = "1=";
        UnSafeArray[22] = " and ";
        UnSafeArray[23] = "eval";
        UnSafeArray[24] = " alert";
        UnSafeArray[25] = "\"";
        UnSafeArray[26] = "--";
        UnSafeArray[27] = "||";

        foreach (string strValue in UnSafeArray)
        {

            if (str.ToLower().IndexOf(strValue) > -1)
            {
                bResult = false;
                break;
            }
        }
        return bResult;
    }
}