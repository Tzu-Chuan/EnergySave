using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Xml;

public partial class handler_getMoneyExecute_Manager : System.Web.UI.Page
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
            string xmlStr = string.Empty;
            string xmlStr1 = string.Empty;
            string xmlStr2 = string.Empty;
            string stage = (Request["stage"] != null) ? Request["stage"].ToString().Trim() : "";
            string city = (Request["city"] != null) ? Request["city"].ToString().Trim() : "";

            me_db._PR_Stage = stage;
            me_db._PR_City = city;

            DataSet ds = me_db.getMoneyList();
            xmlStr1 = DataTableToXml.ConvertDatatableToXML(ds.Tables[0], "dataList", "data_item");
            xmlStr2 = DataTableToXml.ConvertDatatableToXML(ds.Tables[1], "dataList", "data_money");

            xmlStr = "<root>" + xmlStr1 + xmlStr2 + "</root>";
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