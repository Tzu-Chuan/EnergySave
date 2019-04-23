using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class handler_deleteNews : System.Web.UI.Page
{
    News_DB n_db = new News_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        ///-----------------------------------------------------
        ///功    能: 撈公告列表
        ///說明:
        /// * Request["N_Date"]: 日期
        /// * Request["N_Title"]: 關鍵字
        /// * Request["N_Content"]:欲顯示的頁碼, 由零開始
        ///-----------------------------------------------------
        try
        {
            if (LogInfo.mGuid == "" || LogInfo.city == "")
            {
                Response.Write("reLogin");
                return;
            }
            string N_ID = (Request["N_ID"] != null) ? Request["N_ID"].ToString().Trim() : "";

            n_db._N_ID = N_ID;
            //n_db._N_ModId = LogInfo.mGuid;
            //n_db._N_ModDate = DateTime.Now;
            n_db.deleteNews();

            Response.Write("success");
        }
        catch (Exception ex)
        {
            Response.Write("Error:" + ex.Message);
        }
    }
}