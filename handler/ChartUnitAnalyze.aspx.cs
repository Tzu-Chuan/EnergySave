using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class handler_ChartUnitAnalyze : System.Web.UI.Page
{
    Chart_DB c_db = new Chart_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        string City = (LogInfo.competence == "SA") ? Request["City"] : LogInfo.city;
        string Stage = Request["Stage"];

        c_db._strCity = City;
        c_db._strStage = Stage;
        DataSet ds = c_db.getUnitAnalyze();
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

        if (dt.Rows.Count > 0)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (tagname == "parking")
                {
                    jsonStr += "{\"name\":\"集合住宅\",\"y\":" + dt.Rows[i]["RM_SUM1"].ToString() + "},";
                    jsonStr += "{\"name\":\"辦公大樓\",\"y\":" + dt.Rows[i]["RM_SUM2"].ToString() + "},";
                    jsonStr += "{\"name\":\"服務業\",\"y\":" + dt.Rows[i]["RM_SUM3"].ToString() + "}";
                }
                else
                {
                    jsonStr += "{\"name\":\"機關\",\"y\":" + dt.Rows[i]["RM_SUM1"].ToString() + "},";
                    jsonStr += "{\"name\":\"學校\",\"y\":" + dt.Rows[i]["RM_SUM2"].ToString() + "},";
                    jsonStr += "{\"name\":\"服務業\",\"y\":" + dt.Rows[i]["RM_SUM3"].ToString() + "}";
                }
            }

            xmlStr += "<" + tagname + ">";
            xmlStr += "<series>[" + jsonStr + "]</series>";
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