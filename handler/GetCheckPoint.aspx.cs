using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Data;
using System.Web.UI.WebControls;
using System.Xml;

public partial class handler_GetCheckPoint : System.Web.UI.Page
{
    ProjectInfo_DB p_db = new ProjectInfo_DB();
    CheckPoint_DB cp_db = new CheckPoint_DB();
    Member_DB M_Db = new Member_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        ///-----------------------------------------------------
        ///功    能: 查詢查核點列表
        ///說明:
        /// * Request["period"]: 期別
        /// * Request["type"]: 查核點類別
        /// * Request["person_id"]: 登入者ID
        ///-----------------------------------------------------
        XmlDocument xDoc = new XmlDocument();
        try
        {
            if (LogInfo.mGuid == "")
            {
                xDoc = ExceptionUtil.GetErrorMassageDocument("請重新登入");
                Response.ContentType = System.Net.Mime.MediaTypeNames.Text.Xml;
                xDoc.Save(Response.Output);
                return;
            }

            string period = (Request["period"] != null) ? Request["period"].ToString() : "";
            string type = (Request["type"] != null) ? Request["type"].ToString() : "";
            string person_id = (Request["person_id"] != null) ? Request["person_id"].ToString() : "";

            M_Db._M_ID = person_id;
            string project_id = M_Db.getProgectGuidByPersonId();

            string xmlStr = string.Empty;
            string xmlStr2 = string.Empty;

            cp_db._P_Period = period;
            cp_db._P_Type = type;
            cp_db._CP_ProjectId = project_id;
            DataTable dt = cp_db.SelectList();
            xmlStr = CheckPointXML(dt);

            //權限
            switch (LogInfo.competence)
            {
                case "SA":
                    p_db._I_GUID = project_id;
                    DataSet afdt = p_db.getAdminCityFlag();
                    if (afdt.Tables[1].Rows.Count > 0)
                    {
                        if (Int32.Parse(afdt.Tables[0].Rows[0]["num"].ToString()) == 0)
                            xmlStr2 += "<unVisiable>Y</unVisiable>";
                        else if (Int32.Parse(afdt.Tables[0].Rows[0]["num"].ToString()) > 0 && afdt.Tables[1].Rows[0]["I_Flag"].ToString() != "Y")
                            xmlStr2 += "<unVisiable>Y</unVisiable>";
                    }
                    break;
                case "01":
                    if (person_id != LogInfo.id)
                        xmlStr2 = "<unVisiable>Y</unVisiable>";
                    else
                    {
                        p_db._I_City = LogInfo.city;
                        DataTable fdt = p_db.CityFlagCount();
                        if (Int32.Parse(fdt.Rows[0]["num"].ToString()) > 0)
                            xmlStr2 = "<unVisiable>Y</unVisiable>";
                    }
                    break;
                case "02":
                    xmlStr2 = "<unVisiable>Y</unVisiable>";
                    break;
            }

            xmlStr = "<?xml version='1.0' encoding='UTF-8'?><root>" + xmlStr + xmlStr2 + "</root>";

            xDoc.LoadXml(xmlStr);
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
        if (dv.Count > 0)
        {
            XmlDocument doc = new XmlDocument();
            /// 根節點
            XmlElement cpList = doc.CreateElement("cpList");
            doc.AppendChild(cpList);
            XmlElement PushItem = doc.DocumentElement;
            for (int i = 0; i < dv.Count; i++)
            {
                if (i == 0)
                {
                    /// Node - PushItem 
                    PushItem = doc.CreateElement("PushItem");
                    PushItem.SetAttribute("P_Guid", dv[i]["P_Guid"].ToString());
                    PushItem.SetAttribute("P_Type", dv[i]["P_Type"].ToString());
                    PushItem.SetAttribute("P_Period", dv[i]["P_Period"].ToString());
                    string CodeStr = (dv[i]["P_Type"].ToString() == "03") ? getDeviceReplaceCode(dv[i]["P_ItemName"].ToString()) : dv[i]["P_ItemName"].ToString();
                    CodeStr = (dv[i]["P_Type"].ToString() == "04") ? getExPandCode(dv[i]["P_ItemName"].ToString()) : dv[i]["P_ItemName"].ToString();
                    PushItem.SetAttribute("P_ItemName", CodeStr);
                    PushItem.SetAttribute("P_ItemNameCode", dv[i]["P_ItemName"].ToString());
                    PushItem.SetAttribute("P_WorkRatio", dv[i]["P_WorkRatio"].ToString());
                    PushItem.SetAttribute("P_ExFinish", dv[i]["P_ExFinish"].ToString());
                    PushItem.SetAttribute("P_ExType", dv[i]["P_ExType"].ToString());
                    PushItem.SetAttribute("P_ExDeviceType", dv[i]["P_ExDeviceType"].ToString());
                    cpList.AppendChild(PushItem);
                }
                else if (dv[i - 1]["P_Guid"].ToString() != dv[i]["P_Guid"].ToString())
                {
                    /// Node - PushItem 
                    PushItem = doc.CreateElement("PushItem");
                    PushItem.SetAttribute("P_Guid", dv[i]["P_Guid"].ToString());
                    PushItem.SetAttribute("P_Type", dv[i]["P_Type"].ToString());
                    PushItem.SetAttribute("P_Period", dv[i]["P_Period"].ToString());
                    string CodeStr = (dv[i]["P_Type"].ToString() == "03") ? getDeviceReplaceCode(dv[i]["P_ItemName"].ToString()) : dv[i]["P_ItemName"].ToString();
                    CodeStr = (dv[i]["P_Type"].ToString() == "04") ? getExPandCode(dv[i]["P_ItemName"].ToString()) : dv[i]["P_ItemName"].ToString();
                    PushItem.SetAttribute("P_ItemName", CodeStr);
                    PushItem.SetAttribute("P_ItemNameCode", dv[i]["P_ItemName"].ToString());
                    PushItem.SetAttribute("P_WorkRatio", dv[i]["P_WorkRatio"].ToString());
                    PushItem.SetAttribute("P_ExFinish", dv[i]["P_ExFinish"].ToString());
                    PushItem.SetAttribute("P_ExType", dv[i]["P_ExType"].ToString());
                    PushItem.SetAttribute("P_ExDeviceType", dv[i]["P_ExDeviceType"].ToString());
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
                XmlElement sumary = doc.CreateElement("CP_Summary");
                sumary.InnerText = VerificationString(dv[i]["CP_Summary"].ToString());
                cp.AppendChild(sumary);
                XmlElement bdesc = doc.CreateElement("CP_BackwardDesc");
                bdesc.InnerText = VerificationString(dv[i]["CP_BackwardDesc"].ToString());
                cp.AppendChild(bdesc);
                PushItem.AppendChild(cp);

            }
            rVal = doc.OuterXml.ToString();
        }
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

    private string getExPandCode(string item)
    {
        string rVal = string.Empty;
        CheckPoint_DB cp_db = new CheckPoint_DB();
        DataTable dt = cp_db.getExPandTypeName(item);
        if (dt.Rows.Count > 0)
            rVal = dt.Rows[0]["C_Item_cn"].ToString();
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
}