using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class WebPage_ReportSeason : System.Web.UI.Page
{
    public string sGuid,AutoSaveStatus;
    ReportSeasonV2_DB rs_db = new ReportSeasonV2_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        if (LogInfo.mGuid != "")
        {
            sGuid = Guid.NewGuid().ToString("N").ToLower();
            /// 自動存檔開關
            /// open: 開; close: 關;
            AutoSaveStatus = "close";
            if (!IsPostBack)
            {
                if (!string.IsNullOrEmpty(Request["year"]) && !string.IsNullOrEmpty(Request["season"]) && !string.IsNullOrEmpty(Request["stage"]))
                {
                    // 確認有無季報資料
                    DataTable checkDt = rs_db.getSeasonInfo(LogInfo.mGuid, Request["year"].ToString(), Request["season"].ToString(), Request["stage"].ToString());
                    if (checkDt.Rows.Count > 0)
                        sGuid = checkDt.Rows[0]["RS_Guid"].ToString().Trim();

                    // 確認季報有無送審
                    rs_db._RS_Guid = sGuid;
                    DataTable dt = rs_db.getSeasonReview();
                    if (dt.Rows.Count > 0)
                        Response.Write("<script type='text/javascript'>alert('很抱歉，該季報已送審！');location.href='SeasonList.aspx';</script>");
                }
                else
                {
                    Response.Write("<script type='text/javascript'>alert('參數錯誤');location.href='SeasonList.aspx';</script>");
                }
            }
        }
        else
        {
            Response.Write("<script type='text/javascript'>alert('請重新登入');location.href='Login.aspx';</script>");
        }
    }
}