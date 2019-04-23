using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class WebPage_MoneyExecution : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string city = LogInfo.city;
        string mGuid = LogInfo.mGuid;
        string competence = LogInfo.competence;
        //只有有登入狀況而且是01承辦人身分才能進入這個頁面
        if (city=="" || mGuid=="" || competence=="") {
            Response.Write("<script>alert('請重新登入');location.href='Login.aspx';</script>");
        }
        if (competence!="01") {
            Response.Write("<script>alert('沒有權限進入本頁面');location.href='Login.aspx';</script>");
        }
    }
}