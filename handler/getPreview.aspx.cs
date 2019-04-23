using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Xml;

public partial class handler_getPreview : System.Web.UI.Page
{
    ProjectInfo_DB p_db = new ProjectInfo_DB();
    CheckPoint_DB cb_db = new CheckPoint_DB();
    Member_DB m_db = new Member_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        ///-----------------------------------------------------
        ///功    能: 檢視專案明細
        ///說明:
        /// * Request["id"]: 登入者ID
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

            string id = (Request["id"] != null) ? Request["id"].ToString() : "";

            m_db._M_ID = id;
            string project_id = m_db.getProgectGuidByPersonId();

            string xmlStr = string.Empty;
            string xmlStr2 = string.Empty;
            string xmlStr3 = string.Empty;
            string xmlStr4 = string.Empty;
            string xmlStr5 = string.Empty;
            DataSet ds = p_db.getPreview(project_id);
            DataTable dt0 = ds.Tables[0];
            DataTable dt1 = ds.Tables[1];
            DataTable dt2 = ds.Tables[2];

            xmlStr = DataTableToXml.ConvertDatatableToXML(dt0, "dataList", "data_item");
            xmlStr2 = DataTableToXml.ConvertDatatableToXML(dt1, "dataList", "mb_item");
            xmlStr3 = DataTableToXml.ConvertDatatableToXML(dt2, "dataList", "mm_item");

            /// 擴大補助預計完成數
            DataTable ExDt = cb_db.getPushitemExFinish(project_id);
            //xmlStr5 = DataTableToXml.ConvertDatatableToXML(ExDt, "ExList", "ex_item");
            xmlStr5 = ExFinishXML(ExDt);

            /// check 是否為同縣市承辦人
            if (LogInfo.competence != "SA")
            {
                p_db._I_GUID = project_id;
                DataTable ccy = p_db.checkCity();
                if (ccy.Rows.Count > 0)
                {
                    if (ccy.Rows[0]["I_City"].ToString() != LogInfo.city)
                    {
                        Response.Write("DiffCity");
                        return;
                    }
                }
            }

            /// 權限
            switch (LogInfo.competence)
            {
                case "SA":
                    p_db._I_GUID = project_id;
                    DataSet afdt = p_db.getAdminCityFlag();
                    if (afdt.Tables[1].Rows.Count > 0)
                    {
                        if (Int32.Parse(afdt.Tables[0].Rows[0]["num"].ToString()) > 0)
                            xmlStr4 += "<unVisiable>Y</unVisiable>";
                    }
                    break;
                case "01":
                    xmlStr4 = "<unVisiable>Y</unVisiable>";
                    break;
                case "02":
                    xmlStr4 = "<unVisiable>Y</unVisiable>";
                    break;
            }

            xmlStr = "<root>" + xmlStr + xmlStr2 + xmlStr3 + xmlStr4 + xmlStr5 + "</root>";

            xDoc.LoadXml(xmlStr);
        }
        catch (Exception ex)
        {
            xDoc = ExceptionUtil.GetExceptionDocument(ex);
        }
        Response.ContentType = System.Net.Mime.MediaTypeNames.Text.Xml;
        xDoc.Save(Response.Output);
    }

    private string ExFinishXML(DataTable dt)
    {
        string rVal = string.Empty;
        DataView dv = dt.DefaultView;
        if (dv.Count > 0)
        {
            XmlDocument doc = new XmlDocument();
            /// 根節點
            XmlElement exList = doc.CreateElement("exList");
            doc.AppendChild(exList);
            XmlElement ExFinish = doc.DocumentElement;
            for (int i = 0; i < dv.Count; i++)
            {
                if (dv[i]["P_ItemName"].ToString() == "99")
                    continue;
                ExFinish = doc.CreateElement("ex_item");
                ExFinish.SetAttribute("P_Period", dv[i]["P_Period"].ToString());
                ExFinish.SetAttribute("P_ItemNameCode", dv[i]["P_ItemName"].ToString());
                ExFinish.SetAttribute("P_ItemName", dv[i]["C_Item_cn"].ToString());
                ExFinish.SetAttribute("P_ExFinish", dv[i]["P_ExFinish"].ToString());
                exList.AppendChild(ExFinish);
            }
            rVal = doc.OuterXml.ToString();
        }
        return rVal;
    }
}