using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
public partial class WebPage_DetailReportSeason : System.Web.UI.Page
{
    ReportSeasonV2_DB rs_db = new ReportSeasonV2_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (LogInfo.mGuid == "" || LogInfo.city == "")
            {
                Response.Write("<script type='text/javascript'>alert(\'您沒有權限進入該頁面\');location.href=\'Login.aspx\';</script>");
            }

            string strv = "";//RS_ID
            string strSCity = "";//該季報所屬的機關
            if (Request.QueryString["v"] != null)
            {
                strv = Request.QueryString["v"].ToString().Trim();
                //撈出該季報是哪個執行機關
                DataTable dt = rs_db.getICityByRSID(strv);
                if (dt.Rows.Count > 0)
                {
                    strSCity = dt.Rows[0]["I_City"].ToString().Trim();
                }
                //如果不是管理者 去看別的機關的季報明細 就導調
                if (strSCity != LogInfo.city && LogInfo.competence != "SA")
                {
                    Response.Write("<script type='text/javascript'>alert(\'您沒有權限進入該頁面\');location.href=\'Login.aspx\';</script>");
                }

            }
            else {
                Response.Write("<script type='text/javascript'>alert(\'您沒有權限進入該頁面\');location.href=\'Login.aspx\';</script>");
            }
        }
    }
}