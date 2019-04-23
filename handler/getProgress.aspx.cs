using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Xml;

public partial class handler_getProgress : System.Web.UI.Page
{
    ProjectInfo_DB p_db = new ProjectInfo_DB();
    CheckPoint_DB cp_db = new CheckPoint_DB();
    Member_DB mb = new Member_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        ///-----------------------------------------------------
        ///功    能: 查詢預定工作進度
        ///說明:
        /// * Request["mid"]: 登入者ID
        /// * Request["period"]: 期別
        ///-----------------------------------------------------
        XmlDocument xDoc = new XmlDocument();
        try
        {
            if (LogInfo.mGuid == "")
            {
                Response.Write("reLogin");
                return;
            }

            string mid = string.IsNullOrEmpty(Request["mid"]) ? "" : Request["mid"].ToString().Trim();
            string period = string.IsNullOrEmpty(Request["period"]) ? "" : Request["period"].ToString().Trim();

            mb._M_ID = mid;
            string project_id = mb.getProgectGuidByPersonId();

            //抓期程開始結束日
            string pdatestr = string.Empty;
            DataTable pdate = mb.getProjectDate();
            pdatestr = DataTableToXml.ConvertDatatableToXML(pdate, "dataList", "prodate");

            //抓各期時程
            string PeriodDate = string.Empty;
            DataTable priddt = cp_db.getPeriodDate(project_id);
            PeriodDate = DataTableToXml.ConvertDatatableToXML(priddt, "dataList", "priddate");

            string xmlStr = string.Empty;
            string xmlStr1 = string.Empty;
            string xmlStr2 = string.Empty;
            string xmlStr3 = string.Empty;
            string xmlStr4 = string.Empty;
            string xmlStr5 = string.Empty;
            DataTable dt = cp_db.getProgress(project_id, "", period);
            DataTable dt1 = cp_db.getProgress(project_id, "01", period);
            DataTable dt2 = cp_db.getProgress(project_id, "02", period);
            DataTable dt3 = cp_db.getProgress(project_id, "03", period);
            DataTable dt4 = cp_db.getProgress(project_id, "04", period);
            xmlStr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
            xmlStr1 = DataTableToXml.ConvertDatatableToXML(dt1, "dataList", "data01");
            xmlStr2 = DataTableToXml.ConvertDatatableToXML(dt2, "dataList", "data02");
            xmlStr3 = DataTableToXml.ConvertDatatableToXML(dt3, "dataList", "data03");
            xmlStr4 = DataTableToXml.ConvertDatatableToXML(dt4, "dataList", "data04");

            //權限
            switch (LogInfo.competence)
            {
                case "SA":
                    p_db._I_GUID = project_id;
                    DataSet afdt = p_db.getAdminCityFlag();
                    if (afdt.Tables[1].Rows.Count > 0)
                    {
                        if (Int32.Parse(afdt.Tables[0].Rows[0]["num"].ToString()) == 0)
                            xmlStr5 += "<unVisiable>Y</unVisiable>";
                        else if (Int32.Parse(afdt.Tables[0].Rows[0]["num"].ToString()) > 0 && afdt.Tables[1].Rows[0]["I_Flag"].ToString() != "Y")
                            xmlStr5 += "<unVisiable>Y</unVisiable>";
                    }
                    break;
                case "01":
                    if (mid != LogInfo.id)
                        xmlStr5 = "<unVisiable>Y</unVisiable>";
                    else
                    {
                        p_db._I_City = LogInfo.city;
                        DataTable fdt = p_db.CityFlagCount();
                        if (Int32.Parse(fdt.Rows[0]["num"].ToString()) > 0)
                            xmlStr5 = "<unVisiable>Y</unVisiable>";
                    }
                    break;
                case "02":
                    xmlStr5 = "<unVisiable>Y</unVisiable>";
                    break;
            }

            xmlStr5 += "<comp>" + LogInfo.competence + "</comp>";

            xmlStr = "<root>" + xmlStr + xmlStr1 + xmlStr2 + xmlStr3 + xmlStr4 + xmlStr5 + pdatestr + PeriodDate + "</root>";
            xDoc.LoadXml(xmlStr);
        }
        catch (Exception ex)
        {
            xDoc = ExceptionUtil.GetExceptionDocument(ex);
        }
        Response.ContentType = System.Net.Mime.MediaTypeNames.Text.Xml;
        xDoc.Save(Response.Output);
    }
}