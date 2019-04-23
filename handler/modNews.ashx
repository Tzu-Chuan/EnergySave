<%@ WebHandler Language="C#" Class="modNews" %>

using System;
using System.Web;
using System.Web.SessionState;
public class modNews : IHttpHandler, IRequiresSessionState {

    News_DB n_db = new News_DB();
    public void ProcessRequest (HttpContext context) {
        ///-----------------------------------------------------
        ///功    能: 撈公告列表
        ///說明:
        /// * Request["N_Guid"]: Guid
        /// * Request["N_Date"]: 日期
        /// * Request["N_Title"]: 標題
        /// * Request["N_Content"]:內容
        /// * Request["N_Mod"]:add / mod  (新增/修改)
        /// * Request["N_ID"]:修改的那筆資料ID(只有修改會有)
        ///-----------------------------------------------------
        try
        {
            if (LogInfo.mGuid == "" || LogInfo.city == "")
            {
                context.Response.Write("reLogin");
                return;
            }
            string N_Guid = (context.Request["N_Guid"] != null) ? context.Request["N_Guid"].ToString().Trim() : "";
            string N_Date = (context.Request["N_Date"] != null) ? context.Request["N_Date"].ToString().Trim() : "";
            string N_Title = (context.Request["N_Title"] != null) ? context.Request["N_Title"].ToString().Trim() : "";
            string N_Content = (context.Request["N_Content"] != null) ? context.Request["N_Content"].ToString().Trim() : "";
            string modType = (context.Request["N_Mod"] != null) ? context.Request["N_Mod"].ToString().Trim() : "";
            string mguid = LogInfo.mGuid;
            string N_ID= (context.Request["N_ID"] != null) ? context.Request["N_ID"].ToString().Trim() : "";
            DateTime dtnow = DateTime.Now;

            //新增
            if (modType == "add")
            {
                n_db._N_Date = N_Date;
                n_db._N_Title = N_Title;
                n_db._N_Content = context.Server.UrlDecode(N_Content);
                n_db._N_CreateId = LogInfo.mGuid;
                n_db._N_CreateDate = dtnow;
                n_db._N_Guid = N_Guid;
                n_db.addNews();
            }
            else {
                //修改
                n_db._N_ID = N_ID;
                n_db._N_Date = N_Date;
                n_db._N_Title = N_Title;
                n_db._N_Content = context.Server.UrlDecode(N_Content);
                n_db._N_ModId = LogInfo.mGuid;
                n_db._N_ModDate = dtnow;
                n_db.updateNews();
            }

            if (modType=="add") {
                modType = "新增成功";
            }
            if (modType == "mod")
            {
                modType = "修改成功";
            }

            context.Response.Write("<script type='text/JavaScript'>parent.feedback('" + modType + "'); </script>");
        }
        catch (Exception ex)
        {
            context.Response.Write("<script type='text/JavaScript'>parent.feedback('Error:" + ex.Message.Replace("\r\n", "") + "'); </script>");
        }
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}