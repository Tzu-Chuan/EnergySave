using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class handler_ChartExSave : System.Web.UI.Page
{
    Chart_DB c_db = new Chart_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        string City = (LogInfo.competence == "SA") ? Request["City"] : LogInfo.city;
        string Stage = Request["Stage"];

        c_db._strCity = City;
        c_db._strStage = Stage;
        DataSet ds = c_db.getExSave();
        string xmlStr = string.Empty;
        string xmlStr2 = string.Empty;
        xmlStr += getData(ds.Tables[0], "air");
        xmlStr += getData(ds.Tables[1], "old");
        xmlStr += getData(ds.Tables[2], "parking");
        xmlStr += getData(ds.Tables[3], "middle");
        xmlStr += getData(ds.Tables[4], "large");

        xmlStr2 += DataTableToXml.ConvertDatatableToXML(ds.Tables[0], "dataList", "dt_air");
        xmlStr2 += DataTableToXml.ConvertDatatableToXML(ds.Tables[1], "dataList", "dt_old");
        xmlStr2 += DataTableToXml.ConvertDatatableToXML(ds.Tables[2], "dataList", "dt_parking");
        xmlStr2 += DataTableToXml.ConvertDatatableToXML(ds.Tables[3], "dataList", "dt_middle");
        xmlStr2 += DataTableToXml.ConvertDatatableToXML(ds.Tables[4], "dataList", "dt_large");

        xmlStr += "<Comp>" + LogInfo.competence + "</Comp>";
        xmlStr = "<root>" + xmlStr + xmlStr2 + "</root>";

        Response.Write(xmlStr);
    }

    private string getData(DataTable dt, string tagname)
    {
        string xmlStr = string.Empty;
        string jsonStr = string.Empty;
        string jsonStr2 = string.Empty;

        if (dt.Rows.Count > 0)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                jsonStr = "{\"name\":\"當期規劃數節電量\",\"data\":[" + dt.Rows[i]["SUM_S"].ToString() + "]},";
                string yearROC = (Int32.Parse(dt.Rows[i]["RS_Year"].ToString()) - 1911).ToString();
                if (jsonStr2 != "") jsonStr2 += ",";
                jsonStr2 += "{\"name\":\""+ yearROC + "年第" + reROCString(dt.Rows[i]["RS_Season"].ToString()) + "季\",\"data\":[\"\"," + dt.Rows[i]["RM_SUMPre"].ToString() + ","+ dt.Rows[i]["RM_SUMFinish"].ToString() + "]}";
            }

            xmlStr += "<" + tagname + ">";
            xmlStr += "<series>[" + jsonStr + jsonStr2 + "]</series>";
            xmlStr += "</" + tagname + ">";
        }
        return xmlStr;
    }

    private string reROCString(string Str)
    {
        string tmp = string.Empty;
        switch (Str)
        {
            case "1":
                tmp = "一";
                break;
            case "2":
                tmp = "二";
                break;
            case "3":
                tmp = "三";
                break;
            case "4":
                tmp = "四";
                break;
        }
        return tmp;
    }
}