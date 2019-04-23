<%@ WebHandler Language="C#" Class="DetailReport" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Data;
using System.Collections.Generic;
using System.Xml;

public class DetailReport : IHttpHandler,IRequiresSessionState
{
    ReportMonth_DB rm = new ReportMonth_DB();
    //ReportSeason_DB rs = new ReportSeason_DB();
    public void ProcessRequest (HttpContext context) {
        try {
            if (LogInfo.mGuid == "")
            {
                context.Response.Write("timeout");
                return;
            }

            string str_func = string.IsNullOrEmpty(context.Request.Form["func"]) ? "" : context.Request.Form["func"].ToString().Trim();
                string mrguid = string.IsNullOrEmpty(context.Request.Form["str_mid"]) ? "" :  Common.Decrypt(context.Request.Form["str_mid"].ToString().Trim());
            string xmlStr = string.Empty;
            string xmlStr1 = string.Empty;
            string xmlStr2 = string.Empty;
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            switch (str_func)
            {
                //撈月報資料(設備汰換)
                case "load_MonthReportData":
                    //mrguid = "ba8aec0a63d2462ab49759731ee0bfd1";
                    rm._RM_ReportGuid = mrguid;
                    dt = rm.selectMonthReportDetail();
                    if (dt.Rows.Count>0) {
                        xmlStr1 = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                        xmlStr2 = "<data_acc><competence>" + LogInfo.competence + "</competence></data_acc>";
                        xmlStr = "<root>" + xmlStr1 + xmlStr2 + "</root>";
                        context.Response.Write(xmlStr);
                    }
                    break;

                //撈月報資料(擴大補助)
                case "load_MonthReportDataEx":
                    //mrguid = "ba8aec0a63d2462ab49759731ee0bfd1";
                    rm._RM_ReportGuid = mrguid;
                    dt = rm.selectMonthReportDetailEx();
                    if (dt.Rows.Count>0) {
                        xmlStr1 = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                        xmlStr2 = "<data_acc><competence>" + LogInfo.competence + "</competence></data_acc>";
                        xmlStr = "<root>" + xmlStr1 + xmlStr2 + "</root>";
                        context.Response.Write(xmlStr);
                    }
                    break;

                    ////撈季報資料
                    //case "load_SeasonReportData":
                    //    string srguid = string.IsNullOrEmpty(context.Request.Form["str_mid"]) ? "" :Common.Decrypt(context.Request.Form["str_mid"].ToString().Trim());
                    //    //srguid = "93DD7F93-E5FC-4812-8032-AD36E566CFA7";
                    //    rs._RS_ReportGuid = srguid;
                    //    dt = rs.selectSeasonReportDetail();
                    //    if (dt.Rows.Count>0) {
                    //        xmlStr1 = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                    //        xmlStr2 = "<data_acc><competence>" + LogInfo.competence + "</competence></data_acc>";
                    //        //xmlStr2 = "<data_acc><competence>02</competence></data_acc>";
                    //        xmlStr = "<root>" + xmlStr1 + xmlStr2 + "</root>";
                    //        context.Response.Write(xmlStr);
                    //    }
                    //    break;
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