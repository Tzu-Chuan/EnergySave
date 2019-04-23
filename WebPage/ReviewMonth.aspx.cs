using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class WebPage_ReviewMonth : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (LogInfo.mGuid != "")
            {
                if (LogInfo.competence != "02")
                {
                    JavaScript.AlertMessageRedirect(this.Page, "您沒有權限進入此頁面", "ProjectList.aspx");
                }
            }
        }
    }
}