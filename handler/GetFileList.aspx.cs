using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Xml;

public partial class handler_GetFileList : System.Web.UI.Page
{
    Member_DB m_db = new Member_DB();
    File_DB f_db = new File_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        XmlDocument xDoc = new XmlDocument();
        try
        {
            string type = (Request["type"] != null) ? Request["type"].ToString() : "";
            string ProjectGuid = (Request["pGuid"] != null) ? Request["pGuid"].ToString() : "";

            string xmlStr = string.Empty;

            f_db._file_parentid = ProjectGuid;
            f_db._file_type = type;
            DataTable dt = f_db.SelectFile();
            xmlStr = DataTableToXml.ConvertDatatableToXML(dt, "dataList", "data_item");
            xmlStr = handleXMLvalue(xmlStr);
            xmlStr = "<?xml version='1.0' encoding='UTF-8'?><root>" + xmlStr + "</root>";
            xDoc.LoadXml(xmlStr);
        }
        catch (Exception ex)
        {
            xDoc = ExceptionUtil.GetExceptionDocument(ex);
        }
        Response.ContentType = System.Net.Mime.MediaTypeNames.Text.Xml;
        xDoc.Save(Response.Output);
    }

    private string handleXMLvalue(string xmlstr)
    {
        string rVal = string.Empty;
        XmlDocument xDoc = new XmlDocument();
        xDoc.LoadXml(xmlstr);
        //file_id 加密
        XmlNodeList xItem = xDoc.SelectNodes("/dataList/data_item");
        if (xItem.Count > 0)
        {
            for (int i = 0; i < xItem.Count; i++)
            {
                XmlElement xNode = xDoc.CreateElement("download_id");
                xNode.InnerText = Server.UrlEncode(Common.Encrypt(xItem[i].ChildNodes[0].InnerText));
                xItem[i].AppendChild(xNode);
            }
        }
        rVal = xDoc.OuterXml;
        return rVal;
    }
}