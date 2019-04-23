using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class handler_ChartReDevice : System.Web.UI.Page
{
    Chart_DB c_db = new Chart_DB();
    protected void Page_Load(object sender, EventArgs e)
    {

        string City = (LogInfo.competence == "SA") ? Request["City"] : LogInfo.city;
        string Stage = Request["Stage"];

        c_db._strCity = City;
        c_db._strStage= Stage;
        DataSet ds = c_db.getReDevice();
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
        string sum1 = string.Empty; //當期規劃數量
        string sum2 = string.Empty; //申請數量
        string sum3 = string.Empty; //完成數量
        if (dt.Rows.Count > 0)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string yearROC = (Int32.Parse(dt.Rows[i]["RS_Year"].ToString()) - 1911).ToString();
                if (jsonStr != "") jsonStr += ",";
                jsonStr += "\"" + yearROC + "年第" + reROCString(dt.Rows[i]["RS_Season"].ToString()) + "季\"";

                if (sum1 != "") sum1 += ",";
                sum1 += dt.Rows[i]["SUM_S"].ToString();

                if (sum2 != "") sum2 += ",";
                sum2 += dt.Rows[i]["RM_SUM"].ToString();

                if (sum3 != "") sum3 += ",";
                sum3 += dt.Rows[i]["SUM_C"].ToString();
            }
            jsonStr2 = "{\"name\":\"當期規劃數量\",\"data\":[" + sum1 + "]},{\"name\":\"申請數量\",\"data\":[" + sum2 + "]},{\"name\":\"完成數量\",\"data\":[" + sum3 + "]}";
            xmlStr += "<" + tagname + ">";
            xmlStr += "<type>[" + jsonStr + "]</type>";
            xmlStr += "<series>[" + jsonStr2 + "]</series>";
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