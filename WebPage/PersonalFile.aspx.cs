using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class WebPage_PersonalFile : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (Request.QueryString["v"] == null || Request.QueryString["v"] == "")
            {
                JavaScript.AlertMessageRedirect(this.Page, "參數錯誤", "ProjectList.aspx");
            }
        }
    }
}