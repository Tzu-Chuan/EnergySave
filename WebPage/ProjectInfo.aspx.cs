using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class WebPage_ProjectInfo : System.Web.UI.Page
{
    Member_DB m_db = new Member_DB();
    ProjectInfo_DB pj = new ProjectInfo_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (LogInfo.mGuid != "")
            {
                if ((Request.QueryString["v"] == null || Request.QueryString["v"] == "") && LogInfo.competence != "SA")
                {
                    Response.Write("<script type='text/javascript'>alert(\'參數錯誤\');location.href=\'ProjectList.aspx\';</script>");
                    return;
                }
                if (LogInfo.competence == "02")
                {
                    //承辦主管不能進
                    Response.Write("<script type='text/javascript'>alert(\'您沒有權限進入該頁面\');location.href=\'ProjectList.aspx\';</script>");
                    return;
                }
                string mid = LogInfo.id;//登入者ID
                string mcity = LogInfo.city;//登入者的機關
                string pid = Request.QueryString["v"].ToString().Trim();//該計畫基本資料擁有者M_ID
                string pcity = "";//該計畫基本資料擁有者的M_City
                DataTable dt = new DataTable();
                m_db._M_ID = pid;
                dt = m_db.getMemberById();//該計畫基本資料擁有者的基本資料
                if (dt.Rows.Count>0) {
                    pcity = dt.Rows[0]["M_City"].ToString().Trim();
                }
                if (mcity!= pcity && LogInfo.competence!="SA") {
                    //如果登入者跟看的資料是不同縣市
                    Response.Write("<script type='text/javascript'>alert(\'您沒有權限進入該頁面\');location.href=\'ProjectList.aspx\';</script>");
                    return;
                }
                
                m_db._M_ID = (Request.QueryString["v"] == null || Request.QueryString["v"] == "")?LogInfo.id: Request.QueryString["v"].ToString().Trim();
                string pgid = m_db.getProgectGuidByPersonId();
                //如果沒有填寫基本資料
                //只有承辦人可以進去計畫基本資料畫面填寫
                if (pgid=="" && LogInfo.competence != "01") {
                    Response.Write("<script type='text/javascript'>alert(\'您沒有權限進入該頁面\');location.href=\'ProjectList.aspx\';</script>");
                }
                tmpguid.Value = (pgid != "") ? pgid : Guid.NewGuid().ToString("N");
            }
        }
    }
}