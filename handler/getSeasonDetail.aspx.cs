using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Xml;

public partial class handler_getSeasonDetail : System.Web.UI.Page
{
    ReportSeasonV2_DB rs_db = new ReportSeasonV2_DB();
    PushItemDesc_DB pd_db = new PushItemDesc_DB();
    CodeTable_DB ct_db = new CodeTable_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        ///-----------------------------------------------------
        ///功    能: 季報詳細資料
        ///說明:
        /// * Request["Year"]: 年
        /// * Request["Season"]: 季
        /// * Request["Stage"]: 期
        ///-----------------------------------------------------
        XmlDocument xDoc = new XmlDocument();
        try
        {
            string Year = (Request["Year"] != null) ? Request["Year"].ToString().Trim() : "0";
            string Season = (Request["Season"] != null) ? Request["Season"].ToString().Trim() : "";
            string Stage = (Request["Stage"] != null) ? Request["Stage"].ToString().Trim() : "";

            /// 基本資料
            DataSet ds = rs_db.getSeasonDetail(LogInfo.mGuid, Year, Season, Stage);
            /// 累計預定進度&累計實際進度
            DataSet ds2 = rs_db.getSeasonProcess(LogInfo.mGuid, Year, Season, Stage);
            /// 季報資料
            DataTable dt = rs_db.getSeasonInfo(LogInfo.mGuid, Year, Season, Stage);
            /// 辦理情形&差異進度說明
            pd_db._PD_Year = Year;
            pd_db._PD_Season= Season;
            pd_db._PD_Stage = Stage;
            DataTable pd_dt = pd_db.GetPiDesc(LogInfo.mGuid);

            string xmlstr = string.Empty;
            string xmlstr2 = string.Empty;
            string xmlstr3 = string.Empty;
            string xmlstr4 = string.Empty;
            string xmlstr5 = string.Empty;
            string xmlstr6 = string.Empty;
            string xmlstr7 = string.Empty;

            xmlstr = DataTableToXml.ConvertDatatableToXML(dt, "seasonList", "season_item");
            xmlstr2 = DataTableToXml.ConvertDatatableToXML(ds.Tables[0], "dataList", "data_item");
            xmlstr3 = DataTableToXml.ConvertDatatableToXML(ds.Tables[1], "perList", "per_item");
            xmlstr4 = CheckPointXML(ds.Tables[2]);
            xmlstr5 = DataTableToXml.ConvertDatatableToXML(ds2.Tables[0], "processList", "process_item");
            xmlstr6 = DataTableToXml.ConvertDatatableToXML(ds2.Tables[1], "NotIncludeThisMonthProcessList", "process2_item");
            xmlstr7 = PD_XML(pd_dt);
            xmlstr = "<?xml version='1.0' encoding='UTF-8'?><root>" + xmlstr + xmlstr2 + xmlstr3 + xmlstr4 + xmlstr5 + xmlstr6 + xmlstr7 + "</root>";
            xDoc.LoadXml(xmlstr);
        }
        catch (Exception ex)
        {
            xDoc = ExceptionUtil.GetExceptionDocument(ex);
        }
        Response.ContentType = System.Net.Mime.MediaTypeNames.Text.Xml;
        xDoc.Save(Response.Output);
    }

    private string CheckPointXML(DataTable dt)
    {
        string rVal = string.Empty;
        DataView dv = dt.DefaultView;
        if (dv.Count > 0) {
            XmlDocument doc = new XmlDocument();
            /// 根節點
            XmlElement cpList = doc.CreateElement("cpList");
            doc.AppendChild(cpList);
            XmlElement PushItem = doc.DocumentElement;
            for (int i = 0; i < dv.Count; i++)
            {
                if ((dv[i]["P_Type"].ToString() == "03" || dv[i]["P_Type"].ToString() == "04") && dv[i]["P_ItemName"].ToString() == "99")
                    continue;
                if (i == 0)
                {
                    /// Node - PushItem 
                    PushItem = doc.CreateElement("PushItem");
                    PushItem.SetAttribute("P_Guid", dv[i]["P_Guid"].ToString());
                    PushItem.SetAttribute("P_Type", dv[i]["P_Type"].ToString());
                    PushItem.SetAttribute("P_Period", dv[i]["P_Period"].ToString());
                    string CodeStr = string.Empty;
                    switch (dv[i]["P_Type"].ToString())
                    {
                        default:
                            CodeStr = dv[i]["P_ItemName"].ToString();
                            break;
                        case "03":
                            CodeStr = getDeviceReplaceCode(dv[i]["P_ItemName"].ToString());
                            break;
                        case "04":
                            CodeStr = getExType_Cn(dv[i]["P_ItemName"].ToString());
                            break;
                    }
                    PushItem.SetAttribute("P_ItemName", CodeStr);
                    PushItem.SetAttribute("P_WorkRatio", dv[i]["P_WorkRatio"].ToString());
                    PushItem.SetAttribute("P_ExFinish", dv[i]["P_ExFinish"].ToString());
                    PushItem.SetAttribute("EF_Finish", dv[i]["EF_Finish"].ToString());
                    cpList.AppendChild(PushItem);
                }
                else if (dv[i - 1]["P_Guid"].ToString() != dv[i]["P_Guid"].ToString())
                {
                    /// Node - PushItem 
                    PushItem = doc.CreateElement("PushItem");
                    PushItem.SetAttribute("P_Guid", dv[i]["P_Guid"].ToString());
                    PushItem.SetAttribute("P_Type", dv[i]["P_Type"].ToString());
                    PushItem.SetAttribute("P_Period", dv[i]["P_Period"].ToString());
                    string CodeStr = string.Empty;
                    switch (dv[i]["P_Type"].ToString())
                    {
                        default:
                            CodeStr = dv[i]["P_ItemName"].ToString();
                            break;
                        case "03":
                            CodeStr = getDeviceReplaceCode(dv[i]["P_ItemName"].ToString());
                            break;
                        case "04":
                            CodeStr = getExType_Cn(dv[i]["P_ItemName"].ToString());
                            break;
                    }
                    PushItem.SetAttribute("P_ItemName", CodeStr);
                    PushItem.SetAttribute("P_WorkRatio", dv[i]["P_WorkRatio"].ToString());
                    PushItem.SetAttribute("P_ExFinish", dv[i]["P_ExFinish"].ToString());
                    PushItem.SetAttribute("EF_Finish", dv[i]["EF_Finish"].ToString());
                    cpList.AppendChild(PushItem);
                }

                /// Node - CheckPoint 
                XmlElement cp = doc.CreateElement("CheckPoint");
                /// CheckPoint各欄位
                XmlElement gid = doc.CreateElement("CP_Guid");
                gid.InnerText = VerificationString(dv[i]["CP_Guid"].ToString());
                cp.AppendChild(gid);
                XmlElement point = doc.CreateElement("CP_Point");
                point.InnerText = VerificationString(dv[i]["CP_Point"].ToString());
                cp.AppendChild(point);
                XmlElement year = doc.CreateElement("CP_ReserveYear");
                year.InnerText = VerificationString(dv[i]["CP_ReserveYear"].ToString());
                cp.AppendChild(year);
                XmlElement month = doc.CreateElement("CP_ReserveMonth");
                month.InnerText = VerificationString(dv[i]["CP_ReserveMonth"].ToString());
                cp.AppendChild(month);
                XmlElement desc = doc.CreateElement("CP_Desc");
                desc.InnerText = VerificationString(dv[i]["CP_Desc"].ToString());
                cp.AppendChild(desc);
                XmlElement process = doc.CreateElement("CP_Process");
                process.InnerText = VerificationString(dv[i]["CP_Process"].ToString());
                cp.AppendChild(process);
                XmlElement realprocess = doc.CreateElement("CP_RealProcess");
                realprocess.InnerText = VerificationString(dv[i]["CP_RealProcess"].ToString());
                cp.AppendChild(realprocess);

                PushItem.AppendChild(cp);
            }
            rVal = doc.OuterXml.ToString();
        }
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

    private string VerificationString(string str)
    {
        string rVal = string.Empty;
        if (string.IsNullOrEmpty(str))
            rVal = " ";
        else
            rVal = str;
        return rVal;
    }
    
    private string getDeviceReplaceCode(string item)
    {
        string rVal = string.Empty;
        CheckPoint_DB cp_db = new CheckPoint_DB();
        DataTable dt = cp_db.getPushItemName(item);
        if (dt.Rows.Count > 0)
            rVal = dt.Rows[0]["C_Item_cn"].ToString();
        return rVal;
    }

    private string getExType_Cn(string item)
    {
        string str = string.Empty;
        DataTable dt = ct_db.getCn("09", item);
        if (dt.Rows.Count > 0)
            str = dt.Rows[0]["C_Item_cn"].ToString();
        return str;
    }
}
