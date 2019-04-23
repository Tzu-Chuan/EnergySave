using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class WebPage_LogOut : System.Web.UI.Page
{
    Log_DB l_db = new Log_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        l_db._L_Person = LogInfo.mGuid;
        l_db._L_Type = "02";
        l_db._L_IP = Common.GetIP4Address();
        l_db.addLog();
        Session.Abandon();
        Response.Redirect("~/WebPage/Login.aspx");
    }
}