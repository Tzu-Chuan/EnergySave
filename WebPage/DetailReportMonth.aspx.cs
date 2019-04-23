using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class WebPage_DetailReportMonth : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (LogInfo.mGuid == "")
            {
                Response.Write("<script type='text/javascript'>alert(\'請重新登入\');location.href=\'Login.aspx\';</script>");
            }
            //if (LogInfo.mGuid != "")
            //{
            //    if (LogInfo.competence != "02")
            //    {
            //        Response.Write("<script type='text/javascript'>alert(\'您沒有權限進入該頁面\');location.href=\'ProjectList.aspx\';</script>");
            //        return;
            //    }
            //}
            //else
            //{
            //    Response.Write("<script type='text/javascript'>alert(\'請重新登入\');location.href=\'Login.aspx\';</script>");
            //    return;
            //}
        }
    }
}