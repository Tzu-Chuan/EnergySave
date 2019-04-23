using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Xml;

public partial class handler_getMoney_ExecutionById : System.Web.UI.Page
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
            string aid = (Request["aid"] != null) ? Request["aid"].ToString().Trim() : "";

            me_db._PR_ID = aid;

            DataTable dt = me_db.getMoneyById();
            xmlStr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");

            xmlStr = "<root>" + xmlStr + "</root>";
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