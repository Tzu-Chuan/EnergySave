<%@ WebHandler Language="C#" Class="getProjectList" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Data;

public class getProjectList : IHttpHandler,IRequiresSessionState {
    ProjectInfo_DB pi_db = new ProjectInfo_DB();
    Member_DB m_db = new Member_DB();
    public void ProcessRequest (HttpContext context) {
        try
        {
            if (LogInfo.mGuid == "")
            {
                context.Response.Write("reLogin");
                return;
            }

            string SearchStr = (context.Request["SearchStr"] != null) ? context.Request["SearchStr"].ToString() : "";
            string CurrentPage = (context.Request.Form["CurrentPage"] == null) ? "" : context.Request.Form["CurrentPage"].ToString().Trim();

            int Paging = 10; //一頁幾筆
            int pageEnd = (int.Parse(CurrentPage) + 1) * Paging; // +1 是因為分頁從0開始算
            int pageStart = pageEnd - Paging + 1;

            //判斷該承辦人有無資料
            string checkInfo = string.Empty;
            pi_db._I_People = LogInfo.mGuid;
            DataTable cInfoDt = pi_db.selProjectByPeople();
            if (cInfoDt.Rows.Count > 0)
                checkInfo = "<checkinfo>Y</checkinfo>";

            string xmlStr = string.Empty;
            string xmlStr2 = string.Empty;
            string xmlComp = "<comp>" + LogInfo.competence + "</comp>";
            string cityflag = string.Empty;
            DataTable dt = new DataTable();

            switch (LogInfo.competence)
            {
                //系統管理者
                case "SA":
                    pi_db._str_keyword = SearchStr;
                    DataSet ds = pi_db.getProjectListBySA(pageStart.ToString(), pageEnd.ToString());
                    dt = ds.Tables[1];
                    xmlStr = "<total>" + ds.Tables[0].Rows[0]["total"].ToString() + "</total>";
                    xmlStr2 = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                    xmlStr = "<root>" + xmlStr + xmlStr2 + xmlComp + checkInfo + "</root>";
                    context.Response.Write(xmlStr);
                    break;
                //承辦人
                case "01":
                    pi_db._str_keyword = SearchStr;
                    pi_db._I_City = LogInfo.city;
                    pi_db._I_People = LogInfo.mGuid;
                    dt = pi_db.getProjectListByPerson();
                    xmlStr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                    //判斷該城市有無定稿資料
                    pi_db._I_City = LogInfo.city;
                    DataTable dt2 = pi_db.getFlagByCity();
                    if (dt2.Rows.Count > 0)
                        cityflag = "<cityflag>Y</cityflag>";
                    xmlStr = "<root>" + xmlStr + xmlComp + cityflag + checkInfo + "<myGuid>" + LogInfo.mGuid + "</myGuid></root>";
                    context.Response.Write(xmlStr);
                    break;
                //承辦主管
                case "02":
                    pi_db._I_City = LogInfo.city;
                    dt = pi_db.getProjectListByManager();
                    xmlStr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                    xmlStr = "<root>" + xmlStr + xmlComp + checkInfo + "</root>";
                    context.Response.Write(xmlStr);
                    break;
            }
                
        }
        catch (Exception ex)
        {
            context.Response.Write("Error：" + ex.Message);
        }
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}