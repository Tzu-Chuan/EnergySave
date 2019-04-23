using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class WebPage_HistorySeason : System.Web.UI.Page
{
    public string showcity,city;
    protected void Page_Load(object sender, EventArgs e)
    {
        city = LogInfo.city;
        if (LogInfo.competence == "SA")
        {
            city = "";
            showcity = "Y";
        }
    }
}