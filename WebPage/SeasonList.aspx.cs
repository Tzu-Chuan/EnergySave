using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class WebPage_SeasonList : System.Web.UI.Page
{
    ReportSeasonV2_DB rs_db = new ReportSeasonV2_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        if (LogInfo.mGuid != "")
        {
            if (!IsPostBack)
            {
                DataTable dt = rs_db.checkProjectFlag(LogInfo.mGuid);
                if (dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["I_Flag"].ToString() != "Y")
                        Response.Write("<script type='text/javascript'>alert('您沒有權限填寫季報');location.href='ProjectList.aspx';</script>");
                }
                else
                    Response.Write("<script type='text/javascript'>alert('您沒有權限填寫季報');location.href='ProjectList.aspx';</script>");
            }
        }
        else
        {
            Response.Write("<script type='text/javascript'>alert('請重新登入');location.href='Login.aspx';</script>");
        }
    }
}