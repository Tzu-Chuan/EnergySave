using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class WebPage_ReportMonthEx : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Member_DB mb = new Member_DB();
        ReportMonth_DB rmdb = new ReportMonth_DB();
        ProjectInfo_DB pj = new ProjectInfo_DB();
        if (!IsPostBack)
        {
            if (LogInfo.mGuid != "")
            {
                if (!IsPostBack)
                {
                    if (!string.IsNullOrEmpty(Request["year"]) && !string.IsNullOrEmpty(Request["month"]) && !string.IsNullOrEmpty(Request["stage"]))
                    {
                        // 確認月報是不是送審中或審核通過
                        DataTable checkDt = rmdb.chkReportMonth(LogInfo.mGuid, Request["year"].ToString(), Request["month"].ToString(), Request["stage"].ToString(),"02");
                        if (checkDt.Rows.Count > 0)
                            Response.Write("<script type='text/javascript'>alert('很抱歉，該月報已送審！');location.href='ReportMonthList.aspx';</script>");
                    }
                    else
                    {
                        Response.Write("<script type='text/javascript'>alert('參數錯誤');location.href='ReportMonthList.aspx';</script>");
                    }
                }
            }
            else
            {
                Response.Write("<script type='text/javascript'>alert('請重新登入');location.href='Login.aspx';</script>");
            }
        }
    }
}