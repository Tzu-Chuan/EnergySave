using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Xml;

public partial class handler_getNews_ById : System.Web.UI.Page
{
    News_DB n_db = new News_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        ///-----------------------------------------------------
        ///功    能: 撈公告列表
        ///說明:
        /// * Request["N_ID"]: 日期
        ///-----------------------------------------------------
        XmlDocument xDoc = new XmlDocument();
        try
        {
            string N_ID = (Request["N_ID"] != null) ? Request["N_ID"].ToString().Trim() : "";

            if (N_ID == "")
            {
                xDoc = ExceptionUtil.GetErrorMassageDocument("參數錯誤無ID");
            }
            else {
                n_db._N_ID = N_ID;
                DataTable ds = n_db.getNewsByID();

                string xmlstr = string.Empty;
                string xmlstr1 = string.Empty;
                xmlstr1 = DataTableToXml.ConvertDatatableToXML(ds, "dataList", "data_item");
                xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + xmlstr1 + "</root>";
                xDoc.LoadXml(xmlstr);
            }
            
        }
        catch (Exception ex)
        {
            xDoc = ExceptionUtil.GetExceptionDocument(ex);
        }
        Response.ContentType = System.Net.Mime.MediaTypeNames.Text.Xml;
        xDoc.Save(Response.Output);
    }
}