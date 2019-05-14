<%@ WebHandler Language="C#" Class="accReport" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Data;
using System.Collections.Generic;
using System.Xml;

public class accReport : IHttpHandler,IRequiresSessionState
{
    Chart_DB ch_db = new Chart_DB();
    public void ProcessRequest (HttpContext context) {
        try {
            if (LogInfo.mGuid == "")
            {
                context.Response.Write("timeout");
                return;
            }
            if (LogInfo.competence != "SA")
            {
                context.Response.Write("noacc");
                return;
            }

            string str_func = string.IsNullOrEmpty(context.Request.Form["func"]) ? "" : context.Request.Form["func"].ToString().Trim();
            string strCodeType = string.Empty;
            string strStage = string.Empty;
            string strYear = string.Empty;
            string strMonth = string.Empty;
            string strClass = string.Empty;
            string xmlStr = string.Empty;
            string xmlStr1 = string.Empty;
            string xmlStr2 = string.Empty;
            string xmlStr3 = string.Empty;
            DataTable dt = new DataTable();
            DataSet ds = new DataSet();
            switch (str_func)
            {
                //撈"節電基礎及因地制宜工作進度摘要"資料 
                case "load_reportApply":
                    strStage = string.IsNullOrEmpty(context.Request.Form["strStage"]) ? "" :  context.Request.Form["strStage"].ToString().Trim();
                    ch_db._strStage = strStage;
                    dt = ch_db.getReportApply();
                    if (dt.Rows.Count>0) {
                        xmlStr1 = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                        xmlStr = "<root>" + xmlStr1 + "</root>";
                        context.Response.Write(xmlStr);
                    }
                    break;

                //撈"計畫執行進度遭遇困難"資料
                case "load_reportTotalLog":
                    strStage = string.IsNullOrEmpty(context.Request.Form["strStage"]) ? "" :  context.Request.Form["strStage"].ToString().Trim();
                    ch_db._strStage = strStage;
                    dt = ch_db.getReportTotalLog();
                    if (dt.Rows.Count>0) {
                        xmlStr1 = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                        xmlStr = "<root>" + xmlStr1 + "</root>";
                        context.Response.Write(xmlStr);
                    }
                    break;
                //撈"各縣市申請數"資料
                case "load_reportTotalBehind":
                    strStage = string.IsNullOrEmpty(context.Request.Form["strStage"]) ? "" :  context.Request.Form["strStage"].ToString().Trim();
                    ch_db._strStage = strStage;
                    dt = ch_db.getReportTotalBehindg();
                    if (dt.Rows.Count>0) {
                        xmlStr1 = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                        xmlStr = "<root>" + xmlStr1 + "</root>";
                        context.Response.Write(xmlStr);
                    }
                    break;
                //撈"計劃書更動紀錄"資料
                case "load_reportPlanChange":
                    strCodeType = string.IsNullOrEmpty(context.Request.Form["strCodeType"]) ? "" :  context.Request.Form["strCodeType"].ToString().Trim();
                    ch_db._strLType = strCodeType;
                    dt = ch_db.getReportLog();
                    if (dt.Rows.Count>0) {
                        xmlStr1 = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                        xmlStr = "<root>" + xmlStr1 + "</root>";
                        context.Response.Write(xmlStr);
                    }
                    break;
                //撈"當期累計執行進度"資料
                case "load_reportProcess":
                    strStage = string.IsNullOrEmpty(context.Request.Form["strStage"]) ? "" :  context.Request.Form["strStage"].ToString().Trim();
                    ch_db._strStage = strStage;
                    dt = ch_db.getReportProcess();
                    if (dt.Rows.Count>0) {
                        xmlStr1 = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                        xmlStr = "<root>" + xmlStr1 + "</root>";
                        context.Response.Write(xmlStr);
                    }
                    break;
                //撈"工作經費狀態及實支率"資料
                case "load_reportReal":
                    strStage = string.IsNullOrEmpty(context.Request.Form["strStage"]) ? "" :  context.Request.Form["strStage"].ToString().Trim();
                    ch_db._strStage = strStage;
                    dt = ch_db.getReportReal();
                    if (dt.Rows.Count>0) {
                        xmlStr1 = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                        xmlStr = "<root>" + xmlStr1 + "</root>";
                        context.Response.Write(xmlStr);
                    }
                    break;
                //撈"各縣市申請數"資料 BY 月累計
                case "load_reportTotalBehindByM":
                    strStage = string.IsNullOrEmpty(context.Request.Form["strStage"]) ? "" :  context.Request.Form["strStage"].ToString().Trim();
                    ch_db._strStage = strStage;
                    dt = ch_db.getReportTotalBehindgByMonth();
                    if (dt.Rows.Count>0) {
                        xmlStr1 = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                        xmlStr = "<root>" + xmlStr1 + "</root>";
                        context.Response.Write(xmlStr);
                    }
                    break;
                //撈"各月全縣市申請數"資料
                //case "load_reportTotalAllCityByM":
                //    strStage = string.IsNullOrEmpty(context.Request.Form["strStage"]) ? "" :  context.Request.Form["strStage"].ToString().Trim();
                //    strYear = string.IsNullOrEmpty(context.Request.Form["strYear"]) ? "" :  context.Request.Form["strYear"].ToString().Trim();
                //    strMonth = string.IsNullOrEmpty(context.Request.Form["strMonth"]) ? "" :  context.Request.Form["strMonth"].ToString().Trim();
                //    ch_db._strStage = strStage;
                //    ch_db._strYear = strYear;
                //    ch_db._strMonth = strMonth;
                //    dt = ch_db.getAllCityByM();
                //    if (dt.Rows.Count>0) {
                //        xmlStr1 = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
                //        xmlStr = "<root>" + xmlStr1 + "</root>";
                //        context.Response.Write(xmlStr);
                //    }
                //    break;

                //撈"各縣市申請數" 擴大補助 資料
                case "load_reportTotalBehindForEx":
                    strStage = string.IsNullOrEmpty(context.Request.Form["strStage"]) ? "" :  context.Request.Form["strStage"].ToString().Trim();
                    strClass = string.IsNullOrEmpty(context.Request.Form["strClass"]) ? "" :  context.Request.Form["strClass"].ToString().Trim();
                    ch_db._strStage = strStage;
                    ch_db._strExType= strClass;
                    ds = ch_db.getReportTotalBehindgForEx();
                    if (ds.Tables[0].Rows.Count>0) {
                        xmlStr1 = DataTableToXml.ConvertDatatableToXML(ds.Tables[0], "dataList", "dataHead");
                        //xmlStr2 = DataTableToXml.ConvertDatatableToXML(ds.Tables[1], "dataList", "data_item");
                        xmlStr2 = DataXML(ds.Tables[1]);
                        xmlStr = "<root>" + xmlStr1 + xmlStr2 + "</root>";
                        context.Response.Write(xmlStr);
                    }
                    break;

                //撈"各縣市申請數" 擴大補助 資料
                case "load_reportTotalBehindByMForEx":
                    strStage = string.IsNullOrEmpty(context.Request.Form["strStage"]) ? "" :  context.Request.Form["strStage"].ToString().Trim();
                    strClass = string.IsNullOrEmpty(context.Request.Form["strClass"]) ? "" :  context.Request.Form["strClass"].ToString().Trim();
                    ch_db._strStage = strStage;
                    ch_db._strExType= strClass;
                    ds = ch_db.getReportTotalBehindByMForEx();
                    if (ds.Tables[0].Rows.Count>0) {
                        xmlStr1 = DataTableToXml.ConvertDatatableToXML(ds.Tables[0], "dataList", "dataHead");
                        //xmlStr2 = DataTableToXml.ConvertDatatableToXML(ds.Tables[1], "dataList", "data_item");
                        xmlStr2 = DataXMLByM(ds.Tables[1]);
                        xmlStr = "<root>" + xmlStr1 + xmlStr2 + "</root>";
                        context.Response.Write(xmlStr);
                    }
                    break;
            }
        }
        catch (Exception ex) {
            context.Response.Write("Error：" + ex.Message.Replace("'", "\""));
        }
    }

    //擴大補助 各縣市申請數
    private string DataXML(DataTable dt)
    {
        string rVal = string.Empty;
        DataView dv = dt.DefaultView;
        if (dv.Count > 0)
        {
            XmlDocument doc = new XmlDocument();
            /// 根節點
            XmlElement dataList = doc.CreateElement("dataList");
            doc.AppendChild(dataList);
            XmlElement City = doc.DocumentElement;
            for (int i = 0; i < dv.Count; i++)
            {
                /// Node - City 
                if (!(i != 0 && dv[i - 1]["city_I_Guid"].ToString() == dv[i]["city_I_Guid"].ToString() && dv[i - 1]["city_Year"].ToString() == dv[i]["city_Year"].ToString() && dv[i - 1]["city_Season"].ToString() == dv[i]["city_Season"].ToString()))
                {
                    City = doc.CreateElement("City");
                    City.SetAttribute("city_Item", dv[i]["city_Item"].ToString());
                    City.SetAttribute("city_Item_cn", dv[i]["city_Item_cn"].ToString());
                    City.SetAttribute("city_I_Guid", dv[i]["city_I_Guid"].ToString());
                    City.SetAttribute("city_Stage", dv[i]["city_Stage"].ToString());
                    City.SetAttribute("city_Year", dv[i]["city_Year"].ToString());
                    City.SetAttribute("city_Season", dv[i]["city_Season"].ToString());
                    dataList.AppendChild(City);
                }

                /// Node - Data Item 
                XmlElement data_item = doc.CreateElement("data_item");
                /// CheckPoint各欄位
                XmlElement RM_ProjectGuid = doc.CreateElement("RM_ProjectGuid");
                RM_ProjectGuid.InnerText = VerificationString(dv[i]["RM_ProjectGuid"].ToString());
                data_item.AppendChild(RM_ProjectGuid);
                XmlElement RM_Stage = doc.CreateElement("RM_Stage");
                RM_Stage.InnerText = VerificationString(dv[i]["RM_Stage"].ToString());
                data_item.AppendChild(RM_Stage);
                XmlElement RM_Year = doc.CreateElement("RM_Year");
                RM_Year.InnerText = VerificationString(dv[i]["RM_Year"].ToString());
                data_item.AppendChild(RM_Year);
                XmlElement RM_Season = doc.CreateElement("RM_Season");
                RM_Season.InnerText = VerificationString(dv[i]["RM_Season"].ToString());
                data_item.AppendChild(RM_Season);
                XmlElement RM_Planning = doc.CreateElement("RM_Planning");
                RM_Planning.InnerText = VerificationString(dv[i]["RM_Planning"].ToString());
                data_item.AppendChild(RM_Planning);
                XmlElement sumFinsh = doc.CreateElement("sumFinsh");
                sumFinsh.InnerText = VerificationString(dv[i]["sumFinsh"].ToString());
                data_item.AppendChild(sumFinsh);
                XmlElement sumFinsh01 = doc.CreateElement("sumFinsh01");
                sumFinsh01.InnerText = VerificationString(dv[i]["sumFinsh01"].ToString());
                data_item.AppendChild(sumFinsh01);
                XmlElement RM_CPType = doc.CreateElement("RM_CPType");
                RM_CPType.InnerText = VerificationString(dv[i]["RM_CPType"].ToString());
                data_item.AppendChild(RM_CPType);
                XmlElement sum1 = doc.CreateElement("sum1");
                sum1.InnerText = VerificationString(dv[i]["sum1"].ToString());
                data_item.AppendChild(sum1);
                XmlElement sum2 = doc.CreateElement("sum2");
                sum2.InnerText = VerificationString(dv[i]["sum2"].ToString());
                data_item.AppendChild(sum2);
                XmlElement sum3 = doc.CreateElement("sum3");
                sum3.InnerText = VerificationString(dv[i]["sum3"].ToString());
                data_item.AppendChild(sum3);
                XmlElement sum4 = doc.CreateElement("sum4");
                sum4.InnerText = VerificationString(dv[i]["sum4"].ToString());
                data_item.AppendChild(sum4);
                City.AppendChild(data_item);

            }
            rVal = doc.OuterXml.ToString();
        }
        return rVal;
    }

    //擴大補助 各縣市申請數 月累計
    private string DataXMLByM(DataTable dt)
    {
        string rVal = string.Empty;
        DataView dv = dt.DefaultView;
        if (dv.Count > 0)
        {
            XmlDocument doc = new XmlDocument();
            /// 根節點
            XmlElement dataList = doc.CreateElement("dataList");
            doc.AppendChild(dataList);
            XmlElement City = doc.DocumentElement;
            for (int i = 0; i < dv.Count; i++)
            {
                /// Node - City 
                if (!(i != 0 && dv[i - 1]["city_I_Guid"].ToString() == dv[i]["city_I_Guid"].ToString() && dv[i - 1]["city_Year"].ToString() == dv[i]["city_Year"].ToString() && dv[i - 1]["city_Month"].ToString() == dv[i]["city_Month"].ToString()))
                {
                    City = doc.CreateElement("City");
                    City.SetAttribute("city_Item", dv[i]["city_Item"].ToString());
                    City.SetAttribute("city_Item_cn", dv[i]["city_Item_cn"].ToString());
                    City.SetAttribute("city_I_Guid", dv[i]["city_I_Guid"].ToString());
                    City.SetAttribute("city_Stage", dv[i]["city_Stage"].ToString());
                    City.SetAttribute("city_Year", dv[i]["city_Year"].ToString());
                    City.SetAttribute("city_Month", dv[i]["city_Month"].ToString());
                    dataList.AppendChild(City);
                }

                /// Node - Data Item 
                XmlElement data_item = doc.CreateElement("data_item");
                /// CheckPoint各欄位
                XmlElement RM_Planning = doc.CreateElement("RM_Planning");
                RM_Planning.InnerText = VerificationString(dv[i]["RM_Planning"].ToString());
                data_item.AppendChild(RM_Planning);
                XmlElement sumFinsh = doc.CreateElement("sumFinsh");
                sumFinsh.InnerText = VerificationString(dv[i]["sumFinsh"].ToString());
                data_item.AppendChild(sumFinsh);
                XmlElement sumFinsh01 = doc.CreateElement("sumFinsh01");
                sumFinsh01.InnerText = VerificationString(dv[i]["sumFinsh01"].ToString());
                data_item.AppendChild(sumFinsh01);
                XmlElement RM_CPType = doc.CreateElement("RM_CPType");
                RM_CPType.InnerText = VerificationString(dv[i]["RM_CPType"].ToString());
                data_item.AppendChild(RM_CPType);
                XmlElement sum1 = doc.CreateElement("sum1");
                sum1.InnerText = VerificationString(dv[i]["sum1"].ToString());
                data_item.AppendChild(sum1);
                XmlElement sum2 = doc.CreateElement("sum2");
                sum2.InnerText = VerificationString(dv[i]["sum2"].ToString());
                data_item.AppendChild(sum2);
                XmlElement sum3 = doc.CreateElement("sum3");
                sum3.InnerText = VerificationString(dv[i]["sum3"].ToString());
                data_item.AppendChild(sum3);
                XmlElement sum4 = doc.CreateElement("sum4");
                sum4.InnerText = VerificationString(dv[i]["sum4"].ToString());
                data_item.AppendChild(sum4);
                City.AppendChild(data_item);

            }
            rVal = doc.OuterXml.ToString();
        }
        return rVal;
    }

    private string VerificationString(string str)
    {
        string rVal = string.Empty;
        if (string.IsNullOrEmpty(str))
            rVal = " ";
        else
            rVal = str;
        return rVal;
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}