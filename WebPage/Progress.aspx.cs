using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class WebPage_Progress : System.Web.UI.Page
{
    ProjectInfo_DB p_db = new ProjectInfo_DB();
    Member_DB M_Db = new Member_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            if (LogInfo.mGuid != "")
            {
                if (Request.QueryString["v"] == null || Request.QueryString["v"] == "")
                {
                    JavaScript.AlertMessageRedirect(this.Page, "參數錯誤", "ProjectList.aspx");
                }
                else
                {
                    M_Db._M_ID = Request.QueryString["v"].ToString();
                    string project_id = M_Db.getProgectGuidByPersonId();
                    if (project_id != "")
                    {
                        if (LogInfo.competence == "01")
                        {
                            //check 是否為同縣市承辦人
                            p_db._I_GUID = project_id;
                            DataTable ccy = p_db.checkCity();
                            if (ccy.Rows.Count > 0)
                            {
                                if (ccy.Rows[0]["I_City"].ToString() != LogInfo.city)
                                {
                                    JavaScript.AlertMessageRedirect(this.Page, "您沒有權限進入該頁面", "ProjectList.aspx");
                                }
                            }
                        }
                        else if (LogInfo.competence == "02")
                            JavaScript.AlertMessageRedirect(this.Page, "您沒有權限進入該頁面", "ProjectList.aspx");
                    }
                    else
                        JavaScript.AlertMessageRedirect(this.Page, "參數錯誤", "ProjectList.aspx");
                }
            }
        }
    }
}