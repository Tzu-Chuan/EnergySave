using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
/// LogInfo 的摘要描述
/// </summary>
public class LogInfo
{


    /// <summary>
    /// id。
    /// </summary>
    public static string id
    {
        get
        {
            return (HttpContext.Current.Session["id"] != null) ?
                 (!string.IsNullOrEmpty(HttpContext.Current.Session["id"].ToString())) ? HttpContext.Current.Session["id"].ToString() : "" : "";
        }
        set
        {
            HttpContext.Current.Session["id"] = value;
        }
    }

    /// <summary>
    /// Guid。
    /// </summary>
    public static string mGuid
    {
        get
        {
            return (HttpContext.Current.Session["mGuid"] != null) ?
                 (!string.IsNullOrEmpty(HttpContext.Current.Session["mGuid"].ToString())) ? HttpContext.Current.Session["mGuid"].ToString() : "" : "";
        }
        set
        {
            HttpContext.Current.Session["mGuid"] = value;
        }
    }

    /// <summary>
    /// 姓名。
    /// </summary>
    public static string name
    {
        get
        {
            return (HttpContext.Current.Session["name"] != null) ?
                 (!string.IsNullOrEmpty(HttpContext.Current.Session["name"].ToString())) ? HttpContext.Current.Session["name"].ToString() : "" : "";
        }
        set
        {
            HttpContext.Current.Session["name"] = value;
        }
    }

    /// <summary>
    /// 職稱。
    /// </summary>
    public static string jobtitle
    {
        get
        {
            return (HttpContext.Current.Session["jobtitle"] != null) ?
                 (!string.IsNullOrEmpty(HttpContext.Current.Session["jobtitle"].ToString())) ? HttpContext.Current.Session["jobtitle"].ToString() : "" : "";
        }
        set
        {
            HttpContext.Current.Session["jobtitle"] = value;
        }
    }

    /// <summary>
    /// 電話。
    /// </summary>
    public static string tel
    {
        get
        {
            return (HttpContext.Current.Session["tel"] != null) ?
                 (!string.IsNullOrEmpty(HttpContext.Current.Session["tel"].ToString())) ? HttpContext.Current.Session["tel"].ToString() : "" : "";
        }
        set
        {
            HttpContext.Current.Session["tel"] = value;
        }
    }

    /// <summary>
    /// 分機。
    /// </summary>
    public static string ext
    {
        get
        {
            return (HttpContext.Current.Session["ext"] != null) ?
                 (!string.IsNullOrEmpty(HttpContext.Current.Session["ext"].ToString())) ? HttpContext.Current.Session["ext"].ToString() : "" : "";
        }
        set
        {
            HttpContext.Current.Session["ext"] = value;
        }
    }

    /// <summary>
    /// 傳真。
    /// </summary>
    public static string fax
    {
        get
        {
            return (HttpContext.Current.Session["fax"] != null) ?
                 (!string.IsNullOrEmpty(HttpContext.Current.Session["fax"].ToString())) ? HttpContext.Current.Session["fax"].ToString() : "" : "";
        }
        set
        {
            HttpContext.Current.Session["fax"] = value;
        }
    }

    /// <summary>
    /// 手機。
    /// </summary>
    public static string phone
    {
        get
        {
            return (HttpContext.Current.Session["phone"] != null) ?
                 (!string.IsNullOrEmpty(HttpContext.Current.Session["phone"].ToString())) ? HttpContext.Current.Session["phone"].ToString() : "" : "";
        }
        set
        {
            HttpContext.Current.Session["phone"] = value;
        }
    }

    /// <summary>
    /// E-Mail。
    /// </summary>
    public static string email
    {
        get
        {
            return (HttpContext.Current.Session["email"] != null) ?
                 (!string.IsNullOrEmpty(HttpContext.Current.Session["email"].ToString())) ? HttpContext.Current.Session["email"].ToString() : "" : "";
        }
        set
        {
            HttpContext.Current.Session["email"] = value;
        }
    }

    /// <summary>
    /// 地址。
    /// </summary>
    public static string addr
    {
        get
        {
            return (HttpContext.Current.Session["addr"] != null) ?
                 (!string.IsNullOrEmpty(HttpContext.Current.Session["addr"].ToString())) ? HttpContext.Current.Session["addr"].ToString() : "" : "";
        }
        set
        {
            HttpContext.Current.Session["addr"] = value;
        }
    }

    /// <summary>
    /// 執行機關。
    /// </summary>
    public static string city
    {
        get
        {
            return (HttpContext.Current.Session["city"] != null) ?
                 (!string.IsNullOrEmpty(HttpContext.Current.Session["city"].ToString())) ? HttpContext.Current.Session["city"].ToString() : "" : "";
        }
        set
        {
            HttpContext.Current.Session["city"] = value;
        }
    }


    /// <summary>
    /// 承辦局處。
    /// </summary>
    public static string office
    {
        get
        {
            return (HttpContext.Current.Session["office"] != null) ?
                 (!string.IsNullOrEmpty(HttpContext.Current.Session["office"].ToString())) ? HttpContext.Current.Session["office"].ToString() : "" : "";
        }
        set
        {
            HttpContext.Current.Session["office"] = value;
        }
    }

    /// <summary>
    /// 身份/權限。
    /// </summary>
    public static string competence
    {
        get
        {
            return (HttpContext.Current.Session["competence"] != null) ?
                 (!string.IsNullOrEmpty(HttpContext.Current.Session["competence"].ToString())) ? HttpContext.Current.Session["competence"].ToString() : "" : "";
        }
        set
        {
            HttpContext.Current.Session["competence"] = value;
        }
    }

    /// <summary>
    /// 承辦主管。
    /// </summary>
    public static string manager
    {
        get
        {
            return (HttpContext.Current.Session["manager"] != null) ?
                 (!string.IsNullOrEmpty(HttpContext.Current.Session["manager"].ToString())) ? HttpContext.Current.Session["manager"].ToString() : "" : "";
        }
        set
        {
            HttpContext.Current.Session["manager"] = value;
        }
    }
}