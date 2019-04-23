using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class WebPage_total_real : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (LogInfo.mGuid != "")
            {
                if (LogInfo.competence != "SA")
                {
                    JavaScript.AlertMessageRedirect(this.Page, "您沒有權限進入此頁面", "ProjectList.aspx");
                }
            }
        }
    }
}