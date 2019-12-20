<%@ WebHandler Language="C#" Class="ReportMonth" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Data;
using System.Collections.Generic;
using System.Xml;

public class ReportMonth : IHttpHandler,IRequiresSessionState
{
    Member_DB mb = new Member_DB();
    ReportMonth_DB rm = new ReportMonth_DB();
    ProjectInfo_DB pj = new ProjectInfo_DB();
    ReportCheck_DB rc = new ReportCheck_DB();
    Log_DB logdb = new Log_DB();
    public void ProcessRequest (HttpContext context) {
        try {
            if (LogInfo.mGuid == "")
            {
                context.Response.Write("timeout");
                return;
            }

            string str_func = string.IsNullOrEmpty(context.Request.Form["func"]) ? "" : context.Request.Form["func"].ToString().Trim();
            string xmlStr = string.Empty;
            string xmlStr1 = string.Empty;
            string xmlStr2 = string.Empty;
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            switch (str_func)
            {
                case "load_projectbyperson":
                    //string mid = string.IsNullOrEmpty(context.Request.Form["str_mid"]) ? "" : context.Request.Form["str_mid"].ToString().Trim();
                    string mid = LogInfo.id;
                    rm._M_ID = mid;
                    dt = rm.selectMemberById();
                    xmlStr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                    xmlStr = "<root>" + xmlStr + "</root>";
                    context.Response.Write(xmlStr);
                    break;

                //撈期程起訖日
                case "load_stagedate":
                    string sstage = string.IsNullOrEmpty(context.Request.Form["str_stage"]) ? "" : context.Request.Form["str_stage"].ToString().Trim();
                    //string smid = string.IsNullOrEmpty(context.Request.Form["str_mid"]) ? "" : context.Request.Form["str_mid"].ToString().Trim();
                    string smid = LogInfo.id;
                    mb._M_ID = smid;
                    string iguid = mb.getProgectGuidByPersonId();
                    rm._I_Guid = iguid;
                    rm._str_stage = sstage;
                    dt = rm.selectStageDate();
                    if (dt.Rows.Count>0) {
                        xmlStr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                        xmlStr = "<root>" + xmlStr + "</root>";
                        context.Response.Write(xmlStr);
                    }
                    break;

                //撈月報資料
                case "load_MonthReportData":
                    //string lmrdmid = string.IsNullOrEmpty(context.Request.Form["str_mid"]) ? "" : context.Request.Form["str_mid"].ToString().Trim();
                    string lmrdmid = LogInfo.id;
                    string lmrdstage = string.IsNullOrEmpty(context.Request.Form["str_stage"]) ? "" : context.Request.Form["str_stage"].ToString().Trim();
                    string lmrdyear = string.IsNullOrEmpty(context.Request.Form["str_year"]) ? "" : context.Request.Form["str_year"].ToString().Trim();
                    string lmrdmonth = string.IsNullOrEmpty(context.Request.Form["str_month"]) ? "" : context.Request.Form["str_month"].ToString().Trim();
                    string lmrreporttype = string.IsNullOrEmpty(context.Request.Form["str_rmptype"]) ? "" : context.Request.Form["str_rmptype"].ToString().Trim();
                    mb._M_ID = lmrdmid;
                    string lmrdiguid = mb.getProgectGuidByPersonId();
                    rm._RM_ProjectGuid = lmrdiguid;
                    rm._M_ID = lmrdmid;
                    rm._I_Guid = lmrdiguid;
                    rm._P_Period = lmrdstage;
                    rm._RM_Stage = lmrdstage;
                    rm._RM_Year = lmrdyear;
                    rm._RM_Month = lmrdmonth;
                    rm._RM_ReportType = lmrreporttype;
                    DataSet dtc = rm.selectMonthReportBefore();
                    if (dtc.Tables[0].Rows.Count > 0)
                    {
                        //原本有填過資料 但未審核
                        //先判斷有沒有在上一次填寫月報受刪除掉的"推動項目"
                        //推動項目如果有刪掉 就先UPDATE 月報status='D'
                        if (dtc.Tables[1].Rows.Count > 0 && dtc.Tables[1].Rows[0]["RC_CheckType"].ToString().Trim()!="Y")
                        {
                            for (int i = 0; i < dtc.Tables[1].Rows.Count; i++)
                            {
                                //如果 P_Status='D' 表示是刪除掉的推動項目
                                //要去看該推動項目是否有在月報裡面 如果有 就將月報該資料射程 RM_Status='D'
                                if ((dtc.Tables[1].Rows[i]["P_Status"].ToString().Trim() == "D" && dtc.Tables[1].Rows[i]["RM_ID"] != null))
                                {
                                    rm._RM_ID = dtc.Tables[1].Rows[i]["RM_ID"].ToString().Trim();
                                    rm._RM_ModId = LogInfo.mGuid;
                                    rm._RM_PGuid = dtc.Tables[1].Rows[i]["P_Guid"].ToString().Trim();
                                    rm.deleteMonthReport();
                                }
                            }
                        }
                    }

                    if (lmrreporttype == "01")
                    {
                        ds = rm.selectMonthReport();
                    }
                    if (lmrreporttype=="02") {
                        ds = rm.selectMonthReportEx();
                    }
                    //ds = rm.selectMonthReport();
                    if (ds.Tables[0].Rows.Count>0 && ds.Tables[1].Rows.Count>0) {
                        xmlStr1 = DataTableToXml.ConvertDatatableToXML(ds.Tables[0], "dataList", "people_item");
                        xmlStr2 = DataTableToXml.ConvertDatatableToXML(ds.Tables[1], "dataList", "data_item");
                        xmlStr = "<root>" + xmlStr1 + xmlStr2 + "</root>";
                        context.Response.Write(xmlStr);
                    }
                    break;

                //撈年份下拉選單
                case "load_year":
                    //string lymid = string.IsNullOrEmpty(context.Request.Form["str_mid"]) ? "" : context.Request.Form["str_mid"].ToString().Trim();
                    string lymid = LogInfo.id;
                    mb._M_ID = lymid;
                    string lyiguid = mb.getProgectGuidByPersonId();
                    pj._I_GUID = lyiguid;
                    dt = pj.getProjectInfo();
                    if (dt.Rows.Count > 0) {
                        xmlStr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                        xmlStr = "<root>" + xmlStr + "</root>";
                        context.Response.Write(xmlStr);
                    }
                    break;

                //送審
                case "goCheck":
                    string gcmid = LogInfo.id;
                    string gcstage = string.IsNullOrEmpty(context.Request.Form["str_stage"]) ? "" : context.Request.Form["str_stage"].ToString().Trim();
                    string gcyear = string.IsNullOrEmpty(context.Request.Form["str_year"]) ? "" : context.Request.Form["str_year"].ToString().Trim();
                    string gcmonth = string.IsNullOrEmpty(context.Request.Form["str_month"]) ? "" : context.Request.Form["str_month"].ToString().Trim();
                    string gcrcreporttype = string.IsNullOrEmpty(context.Request.Form["str_rcreporttype"]) ? "" : context.Request.Form["str_rcreporttype"].ToString().Trim();
                    string chk_reportGuid = "";
                    mb._M_ID = gcmid;
                    string gciguid = mb.getProgectGuidByPersonId();
                    rm._M_ID = gcmid;
                    rm._I_Guid = gciguid;
                    rm._P_Period = gcstage;
                    rm._RM_Year = gcyear;
                    rm._RM_Month = gcmonth;
                    if (gcrcreporttype=="01") {
                        ds = rm.selectMonthReport();
                    }
                    if (gcrcreporttype=="03") {
                            ds = rm.selectMonthReportEx();
                    }

                    if (ds.Tables[1].Rows.Count>0) {
                        chk_reportGuid = ds.Tables[1].Rows[0]["RM_ReportGuid"].ToString().Trim();
                    }
                    //先檢查有沒有資料 要先儲存過才能送審
                    rm._RM_ReportGuid = chk_reportGuid;
                    DataTable dtchk = rm.selectMonthReportByGuid();
                    if (dtchk.Rows.Count == 0)
                    {
                        context.Response.Write("nodata");
                        return;
                    }
                    rc._RC_ReportType = gcrcreporttype;//01月報(設備汰換) 02季報 03月報(擴大補助)
                    rc._RC_ReportGuid = chk_reportGuid;//審核用ReportGuid
                    rc._RC_PeopleGuid = LogInfo.mGuid;//承辦人GUID
                    rc._RC_Stage = gcstage;//期
                    rc._RC_Year = gcyear;//年
                    rc._RC_Month = gcmonth;//月
                    rc._RC_Season = "";//季(季報才有)
                    rc.addMonth();
                    //紀錄LOG
                    logdb._L_Type = "04";
                    logdb._L_Person = LogInfo.mGuid;// "登入者Guid"
                    logdb._L_IP = Common.GetIP4Address();
                    logdb._L_ModItemGuid = chk_reportGuid;  //"該月報Guid"
                    logdb._L_Desc = "第" + gcstage + "期" + gcyear + "年" + gcmonth + "月月報送審";// 可以不用帶 
                    logdb.addLog();
                    context.Response.Write("success");
                    break;
            }
        }
        catch (Exception ex) {
            context.Response.Write("Error：" + ex.Message.Replace("'", "\""));
        }

    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}