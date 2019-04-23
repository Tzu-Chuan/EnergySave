using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Xml;

public partial class handler_getMoney_Execution : System.Web.UI.Page
{
    MoneyExecute_DB me_db = new MoneyExecute_DB();
    protected void Page_Load(object sender, EventArgs e)
    {

        XmlDocument xDoc = new XmlDocument();
        try
        {
            if (LogInfo.mGuid == "" || LogInfo.city == "")
            {
                Response.Write("reLogin");
                return;
            }

            string stage = (Request["stage"] != null) ? Request["stage"].ToString().Trim() : "";
            string cityid = LogInfo.city;
            string xmlStr = string.Empty;
            string xmlStr1 = string.Empty;
            string xmlStr2 = string.Empty;
            string xmlStr3 = string.Empty;
            string xmlStr4 = string.Empty;

            me_db._str_stage = stage;
            me_db._str_city = cityid;

            DataSet ds = me_db.getMoney();
            xmlStr1 = DataTableToXml.ConvertDatatableToXML(ds.Tables[0], "dataList", "data_item");
            if (ds.Tables[1].Rows.Count > 0)
                xmlStr2 = "<total>" + ds.Tables[1].Rows[0]["MoneyAll"].ToString() + "</total>";
            if (ds.Tables[2].Rows.Count > 0)
                xmlStr3 = "<cityName>" + ds.Tables[2].Rows[0]["C_Item_cn"].ToString() + "</cityName>";
            xmlStr4 = "<cityno>" + LogInfo.city + "</cityno>";

            xmlStr = "<root>" + xmlStr1 + xmlStr2 + xmlStr3 + xmlStr4 + "</root>";
            xDoc.LoadXml(xmlStr);
        }
        catch (Exception ex)
        {
            xDoc = ExceptionUtil.GetExceptionDocument(ex);
        }
        Response.ContentType = System.Net.Mime.MediaTypeNames.Text.Xml;
        xDoc.Save(Response.Output);

    }
    public bool IsReusable
    {
        get
        {
            return false;
        }
    }
}