using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class handler_delMoney_Execution : System.Web.UI.Page
{
    MoneyExecute_DB me_db = new MoneyExecute_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            if (LogInfo.mGuid == "" || LogInfo.city == "")
            {
                Response.Write("reLogin");
                return;
            }
            string PR_ID = (Request["PR_ID"] != null) ? Request["PR_ID"].ToString().Trim() : "";

            me_db._PR_ID = PR_ID;
            me_db._PR_ModId = LogInfo.mGuid;
            me_db.delMoney();
            
            Response.Write("success");
        }
        catch (Exception ex)
        {
            Response.Write("Error:" + ex.Message);
        }
    }
}