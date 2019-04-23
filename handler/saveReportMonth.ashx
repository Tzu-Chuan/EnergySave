<%@ WebHandler Language="C#" Class="saveReportMonth" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Data;

public class saveReportMonth : IHttpHandler,IRequiresSessionState
{
    Member_DB mb = new Member_DB();
    ReportMonth_DB rm = new ReportMonth_DB();
    Log_DB logdb = new Log_DB();
    public void ProcessRequest (HttpContext context) {
        try
        {
            if (LogInfo.mGuid == "")
            {
                context.Response.Write("timeout");
                return;
            }

            //string mid = (context.Request["mid"] != null) ? context.Request["mid"].ToString() : "";
            string mid = LogInfo.id;
            string stage = (context.Request["stage"] != null) ? context.Request["stage"].ToString() : "";
            string year = (context.Request["year"] != null) ? context.Request["year"].ToString() : "";
            string month = (context.Request["month"] != null) ? context.Request["month"].ToString() : "";
            mb._M_ID = mid;
            string iguid = mb.getProgectGuidByPersonId();
            if (iguid=="") {
                context.Response.Write("<script type='text/JavaScript'>parent.feedback('參數錯誤');</script>");
                return;
            }
            //小分類
            string[] report_type = null;
            string[] report_P_Guid = null;
            string[] report_Guid = null;
            //值
            string RM_Finish = null;
            string RM_Finish01 = null;//無風管的累計完成數(台)
            string RM_Planning = null;
            string RM_Type1Value1 = null;
            string RM_Type1Value2 = null;
            string RM_Type1Value3 = null;
            string RM_Type1ValueSum = null;
            string RM_Type2Value1 = null;
            string RM_Type2Value2 = null;
            string RM_Type2Value3 = null;
            string RM_Type2ValueSum = null;
            string RM_Type3Value1 = null;
            string RM_Type3Value2 = null;
            string RM_Type3Value3 = null;
            string RM_Type3ValueSum = null;
            string RM_Type4Value1 = null;
            string RM_Type4Value2 = null;
            string RM_Type4Value3 = null;
            string RM_Type4ValueSum = null;
            string RM_PreVal = null;
            string RM_ChkVal = null;
            string RM_NotChkVal = null;
            string RM_Remark = null;
            string RM_ReportGuid = Guid.NewGuid().ToString("N");
            string mod_RM_ReportGuid = null;
            report_type = (context.Request["report_type"] != null) ? context.Request["report_type"].ToString().Split(',') : null;
            report_P_Guid = (context.Request["report_P_Guid"] != null) ? context.Request["report_P_Guid"].ToString().Split(',') : null;
            report_Guid = (context.Request["report_Guid"] != null) ? context.Request["report_Guid"].ToString().Split(',') : null;
            mod_RM_ReportGuid= (report_Guid[0] != null) ? report_Guid[0].ToString() : null;

            for (int i=0;i<report_type.Length;i++) {
                RM_Finish= (context.Request[""+report_type[i]+"RM_Finish"] != null) ? context.Request[""+report_type[i]+"RM_Finish"].ToString() : null;
                //無風管的累計完成數(台)
                RM_Finish01= (context.Request[""+report_type[i]+"RM_Finish01"] != null) ? context.Request[""+report_type[i]+"RM_Finish01"].ToString() : null;
                RM_Planning= (context.Request[""+report_type[i]+"RM_Planning"] != null) ? context.Request[""+report_type[i]+"RM_Planning"].ToString() : null;
                RM_Type1Value1= (context.Request[""+report_type[i]+"RM_Type1Value1"] != null) ? context.Request[""+report_type[i]+"RM_Type1Value1"].ToString() : null;
                RM_Type1Value2= (context.Request[""+report_type[i]+"RM_Type1Value2"] != null) ? context.Request[""+report_type[i]+"RM_Type1Value2"].ToString() : null;
                RM_Type1Value3= (context.Request[""+report_type[i]+"RM_Type1Value3"] != null) ? context.Request[""+report_type[i]+"RM_Type1Value3"].ToString() : null;
                RM_Type1ValueSum= (context.Request[""+report_type[i]+"RM_Type1ValueSum"] != null) ? context.Request[""+report_type[i]+"RM_Type1ValueSum"].ToString() : null;
                RM_Type2Value1= (context.Request[""+report_type[i]+"RM_Type2Value1"] != null) ? context.Request[""+report_type[i]+"RM_Type2Value1"].ToString() : null;
                RM_Type2Value2= (context.Request[""+report_type[i]+"RM_Type2Value2"] != null) ? context.Request[""+report_type[i]+"RM_Type2Value2"].ToString() : null;
                RM_Type2Value3= (context.Request[""+report_type[i]+"RM_Type2Value3"] != null) ? context.Request[""+report_type[i]+"RM_Type2Value3"].ToString() : null;
                RM_Type2ValueSum= (context.Request[""+report_type[i]+"RM_Type2ValueSum"] != null) ? context.Request[""+report_type[i]+"RM_Type2ValueSum"].ToString() : null;
                RM_Type3Value1= (context.Request[""+report_type[i]+"RM_Type3Value1"] != null) ? context.Request[""+report_type[i]+"RM_Type3Value1"].ToString() : null;
                RM_Type3Value2= (context.Request[""+report_type[i]+"RM_Type3Value2"] != null) ? context.Request[""+report_type[i]+"RM_Type3Value2"].ToString() : null;
                RM_Type3Value3= (context.Request[""+report_type[i]+"RM_Type3Value3"] != null) ? context.Request[""+report_type[i]+"RM_Type3Value3"].ToString() : null;
                RM_Type3ValueSum= (context.Request[""+report_type[i]+"RM_Type3ValueSum"] != null) ? context.Request[""+report_type[i]+"RM_Type3ValueSum"].ToString() : null;
                RM_Type4Value1= (context.Request[""+report_type[i]+"RM_Type4Value1"] != null) ? context.Request[""+report_type[i]+"RM_Type4Value1"].ToString() : null;
                RM_Type4Value2= (context.Request[""+report_type[i]+"RM_Type4Value2"] != null) ? context.Request[""+report_type[i]+"RM_Type4Value2"].ToString() : null;
                RM_Type4Value3= (context.Request[""+report_type[i]+"RM_Type4Value3"] != null) ? context.Request[""+report_type[i]+"RM_Type4Value3"].ToString() : null;
                RM_Type4ValueSum= (context.Request[""+report_type[i]+"RM_Type4ValueSum"] != null) ? context.Request[""+report_type[i]+"RM_Type4ValueSum"].ToString() : null;
                RM_PreVal= (context.Request[""+report_type[i]+"RM_PreVal"] != null) ? context.Request[""+report_type[i]+"RM_PreVal"].ToString() : null;
                RM_ChkVal= (context.Request[""+report_type[i]+"RM_ChkVal"] != null) ? context.Request[""+report_type[i]+"RM_ChkVal"].ToString() : null;
                RM_NotChkVal= (context.Request[""+report_type[i]+"RM_NotChkVal"] != null) ? context.Request[""+report_type[i]+"RM_NotChkVal"].ToString() : null;
                RM_Remark= (context.Request[""+report_type[i]+"RM_Remark"] != null) ? context.Request[""+report_type[i]+"RM_Remark"].ToString() : null;

                rm._RM_ProjectGuid = iguid;
                rm._RM_PGuid = report_P_Guid[i];
                rm._RM_Stage = stage;
                rm._RM_Year = year;
                rm._RM_Month = month;
                rm._RM_CPType = report_type[i].ToString().Trim();
                rm._RM_ModId = LogInfo.mGuid;
                rm._RM_Finish = (RM_Finish != null && RM_Finish != "") ? Convert.ToDecimal(RM_Finish):0;
                rm._RM_Finish01 = (RM_Finish01 != null && RM_Finish01 != "") ? Convert.ToInt32(RM_Finish01):0;//無風管的累計完成數(台)
                rm._RM_Planning =(RM_Planning != null && RM_Planning != "") ? Convert.ToDecimal(RM_Planning):0;
                rm._RM_Type1Value1 =(RM_Type1Value1 != null && RM_Type1Value1 != "") ? Convert.ToInt32(RM_Type1Value1):0;
                rm._RM_Type1Value2 =(RM_Type1Value2 != null && RM_Type1Value2 != "") ? Convert.ToInt32(RM_Type1Value2):0;
                rm._RM_Type1Value3 =(RM_Type1Value3 != null && RM_Type1Value3 != "") ? Convert.ToInt32(RM_Type1Value3):0;
                rm._RM_Type1ValueSum =(RM_Type1ValueSum != null && RM_Type1ValueSum != "") ? Convert.ToInt32(RM_Type1ValueSum):0;
                rm._RM_Type2Value1 =(RM_Type2Value1 != null && RM_Type2Value1 != "") ? Convert.ToInt32(RM_Type2Value1):0;
                rm._RM_Type2Value2 =(RM_Type2Value2 != null && RM_Type2Value2 != "") ? Convert.ToInt32(RM_Type2Value2):0;
                rm._RM_Type2Value3 =(RM_Type2Value3 != null && RM_Type2Value3 != "") ? Convert.ToInt32(RM_Type2Value3):0;
                rm._RM_Type2ValueSum =(RM_Type2ValueSum != null && RM_Type2ValueSum != "") ? Convert.ToInt32(RM_Type2ValueSum):0;
                rm._RM_Type3Value1 =(RM_Type3Value1 != null && RM_Type3Value1 != "") ? Convert.ToDecimal(RM_Type3Value1):0;
                rm._RM_Type3Value2 =(RM_Type3Value2 != null && RM_Type3Value2 != "") ? Convert.ToDecimal(RM_Type3Value2):0;
                rm._RM_Type3Value3 =(RM_Type3Value3 != null && RM_Type3Value3 != "") ? Convert.ToDecimal(RM_Type3Value3):0;
                rm._RM_Type3ValueSum =(RM_Type3ValueSum != null && RM_Type3ValueSum != "") ? Convert.ToDecimal(RM_Type3ValueSum):0;
                rm._RM_Type4Value1 =(RM_Type4Value1 != null && RM_Type4Value1 != "") ? Convert.ToDecimal(RM_Type4Value1):0;
                rm._RM_Type4Value2 =(RM_Type4Value2 != null && RM_Type4Value2 != "") ? Convert.ToDecimal(RM_Type4Value2):0;
                rm._RM_Type4Value3 =(RM_Type4Value3 != null && RM_Type4Value3 != "") ? Convert.ToDecimal(RM_Type4Value3):0;
                rm._RM_Type4ValueSum =(RM_Type4ValueSum != null && RM_Type4ValueSum != "") ? Convert.ToDecimal(RM_Type4ValueSum):0;
                rm._RM_PreVal =(RM_PreVal != null && RM_PreVal != "") ? Convert.ToDecimal(RM_PreVal):0;
                rm._RM_ChkVal =(RM_ChkVal != null && RM_ChkVal != "") ? Convert.ToDecimal(RM_ChkVal):0;
                rm._RM_NotChkVal =(RM_NotChkVal != null && RM_NotChkVal != "") ? Convert.ToDecimal(RM_NotChkVal):0;
                rm._RM_Remark =RM_Remark;
                rm._RM_ReportType = "01";//01設備汰換 02擴大補助
                //檢查有沒有資料已經存在
                DataTable chkMR = rm.selectMonthReportByPjno();
                if (chkMR.Rows.Count > 0)
                {
                    //有資料=>修改
                    rm.modMonthReport("mod");
                }
                else {
                    //沒資料=>新增
                    rm._RM_ReportGuid = RM_ReportGuid;
                    rm.modMonthReport("add");
                }

            }
            //紀錄LOG
            logdb._L_Type = "08";
            logdb._L_Person = LogInfo.mGuid;// "登入者Guid"
            logdb._L_IP = Common.GetIP4Address();
            if (mod_RM_ReportGuid != "")
            {
                logdb._L_ModItemGuid = mod_RM_ReportGuid;  //"該月報Guid"
                logdb._L_Desc = "修改第" + stage + "期" + year + "-" + month + "月報";
            }
            else {
                logdb._L_ModItemGuid = RM_ReportGuid;  //"該月報Guid"
                logdb._L_Desc = "新增第" + stage + "期" + year + "-" + month + "月報";
            }


            logdb.addLog();

            context.Response.Write("<script type='text/JavaScript'>parent.feedback('succeed');</script>");
        }
        catch (Exception ex)
        {
            context.Response.Write("<script type='text/JavaScript'>parent.feedback('Error：" + ex.Message.Replace("'", "\"") + "');</script>");
        }
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}