using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class Manage_Admin : System.Web.UI.MasterPage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (LogInfo.mGuid != "")
        {
            if (LogInfo.competence.ToUpper() != "SA")
            {
                Response.Write("<script>alert('您沒有權限進入此頁面');window.location='../WebPage/Login.aspx';</script>");
            }
            mName.Text = LogInfo.name;
        }
        else
        {
            Response.Write("<script>alert('請先登入');window.location='../WebPage/Login.aspx';</script>");
        }
    }
}
