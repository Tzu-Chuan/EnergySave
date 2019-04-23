using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using System.Data;

public partial class handler_getHistoryDetail_S : System.Web.UI.Page
{
    ReportSeasonV2_DB rs_db = new ReportSeasonV2_DB();
    PushItemDesc_DB pd_db = new PushItemDesc_DB();
    ExpandFinish_DB ef_db = new ExpandFinish_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        //功能
        //季報歷史資料查詢
        XmlDocument xDoc = new XmlDocument();

        if (LogInfo.mGuid == "")
        {
            xDoc = ExceptionUtil.GetErrorMassageDocument("timeout");
            return;
        }

        try {
            string xmlStr = string.Empty;
            string xmlStr1 = string.Empty;
            string xmlStr2 = string.Empty;
            string xmlStr3 = string.Empty;
            string xmlStr4 = string.Empty;
            DataTable dt = new DataTable();
            string RS_ID = (Request["RS_ID"] != null) ? Request["RS_ID"].ToString() : "";
            dt = rs_db.getHistorySeason(RS_ID);
            xmlStr1 = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
            if(dt.Rows.Count>0)
                xmlStr1 = handleXMLvalue(xmlStr1);
            xmlStr2 = "<person><comp>" + LogInfo.competence + "</comp></person>";
            /// 辦理情形&差異進度說明
            if (dt.Rows.Count > 0)
            {
                pd_db._PD_Year = dt.Rows[0]["RS_Year"].ToString();
                pd_db._PD_Season = dt.Rows[0]["RS_Season"].ToString();
                pd_db._PD_Stage = dt.Rows[0]["RS_Stage"].ToString();
            }
            DataTable pd_dt = pd_db.GetPiDesc(LogInfo.mGuid);
            xmlStr3 = PD_XML(pd_dt);
            DataTable ef_dt = ef_db.GetDataByRSID(RS_ID);
            xmlStr4 = DataTableToXml.ConvertDatatableToXML(ef_dt, "ExList", "ex_item");
            xmlStr = "<?xml version='1.0' encoding='UTF-8'?><root>" + xmlStr1 + xmlStr2 + xmlStr3 + xmlStr4 + "</root>";
            xDoc.LoadXml(xmlStr);

        }
        catch (Exception ex)
        {
            xDoc = ExceptionUtil.GetExceptionDocument(ex);
        }
        Response.ContentType = System.Net.Mime.MediaTypeNames.Text.Xml;
        xDoc.Save(Response.Output);
    }

    private string handleXMLvalue(string xmlstr)
    {
        string rVal = string.Empty;
        XmlDocument xDoc = new XmlDocument();
        xDoc.LoadXml(xmlstr);
        //if RC_CheckDate is null
        string checkdate = DateTime.Parse(xDoc.SelectSingleNode("/dataList/data_item/RC_CheckDate").InnerText).ToString("yyyy/MM/dd");
        if (checkdate == "1900/01/01")
            xDoc.SelectSingleNode("/dataList/data_item/RC_CheckDate").InnerText = " ";
        //RS_ID 加密
        xDoc.SelectSingleNode("/dataList/data_item/RS_ID").InnerText = Server.UrlEncode(Common.Encrypt(xDoc.SelectSingleNode("/dataList/data_item/RS_ID").InnerText));
        rVal = xDoc.OuterXml;
        return rVal;
    }

    private string PD_XML(DataTable dt)
    {
        string rVal = string.Empty;
        DataView dv = dt.DefaultView;
        if (dv.Count > 0)
        {
            XmlDocument doc = new XmlDocument();
            /// 根節點
            XmlElement pdList = doc.CreateElement("pdList");
            doc.AppendChild(pdList);
            XmlElement pdItem = doc.DocumentElement;
            for (int i = 0; i < dv.Count; i++)
            {
                pdItem = doc.CreateElement("pd_item");
                pdItem.SetAttribute("PD_ID", dv[i]["PD_ID"].ToString());
                pdItem.SetAttribute("PD_Guid", dv[i]["PD_Guid"].ToString());
                pdItem.SetAttribute("PD_PushitemGuid", dv[i]["PD_PushitemGuid"].ToString());
                pdItem.SetAttribute("PD_ProjectGuid", dv[i]["PD_ProjectGuid"].ToString());
                pdItem.SetAttribute("PD_Stage", dv[i]["PD_Stage"].ToString());
                pdItem.SetAttribute("PD_Year", dv[i]["PD_Year"].ToString());
                pdItem.SetAttribute("PD_Season", dv[i]["PD_Season"].ToString());
                pdItem.SetAttribute("PD_Summary", dv[i]["PD_Summary"].ToString());
                pdItem.SetAttribute("PD_BackwardDesc", dv[i]["PD_BackwardDesc"].ToString());
                pdItem.SetAttribute("PD_CreateDate", DateTime.Parse(dv[i]["PD_CreateDate"].ToString()).ToString("yyyy/MM/dd"));
                pdItem.SetAttribute("PD_ModDate", DateTime.Parse(dv[i]["PD_ModDate"].ToString()).ToString("yyyy/MM/dd"));
                pdItem.SetAttribute("PD_Status", dv[i]["PD_Status"].ToString());
                pdList.AppendChild(pdItem);
            }
            rVal = doc.OuterXml.ToString();
        }
        return rVal;
    }
}