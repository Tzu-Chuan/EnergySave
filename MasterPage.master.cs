using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class MasterPage : System.Web.UI.MasterPage
{
    public string MgBtn = "Show";
    protected void Page_Load(object sender, EventArgs e)
    {
        if (LogInfo.mGuid != "")
        {
            switch (LogInfo.competence.ToUpper())
            {
                case "SA":
                    report_li.Visible = false;
                    money_li.Visible = false;
                    admin_li.Visible = false;
                    break;
                case "01":
                    MgBtn = "Hide";
                    sa_form.Visible = false;
                    admin_li.Visible = false;
                    break;
                case "02":
                    MgBtn = "Hide";
                    sa_form.Visible = false;
                    money_li.Visible = false;
                    report_li.Visible = false;
                    break;
            }
            mName.Text = LogInfo.name;
        }
        else
        {
            Response.Write("<script>alert('請先登入');window.location='Login.aspx';</script>");
        }
    }

    protected void pfbtn(object sender, EventArgs e)
    {
        Response.Redirect("~/WebPage/PersonalFile.aspx?v=" + LogInfo.id);
    }
}
