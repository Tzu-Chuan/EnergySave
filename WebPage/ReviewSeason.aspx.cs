using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class WebPage_ReviewSeason : System.Web.UI.Page
{
    public string city;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (LogInfo.mGuid != "")
            {
                city = LogInfo.city;
                if (LogInfo.competence != "02")
                {
                    JavaScript.AlertMessageRedirect(this.Page, "您沒有權限進入此頁面", "ProjectList.aspx");
                }
            }
        }
    }
}