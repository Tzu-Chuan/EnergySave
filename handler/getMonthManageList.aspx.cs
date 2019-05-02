using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Xml;

public partial class handler_getMonthManageList : System.Web.UI.Page
{
    OtherManage_DB om_db = new OtherManage_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        ///-----------------------------------------------------
        ///功    能: 查詢月報抽單列表
        ///說明:
        /// * Request["PageNo"]:欲顯示的頁碼, 由零開始
        /// * Request["PageSize"]:每頁顯示的資料筆數, 未指定預設10
        /// * Request["city"]: 執行機關
        /// * Request["rcType"]: 月報類別
        ///-----------------------------------------------------
        XmlDocument xDoc = new XmlDocument();
        try
        {
            string PageNo = (Request["PageNo"] != null) ? Request["PageNo"].ToString().Trim() : "";
            int PageSize = (Request["PageSize"] != null) ? int.Parse(Request["PageSize"].ToString().Trim()) : 10;
            string city = (Request["city"] != null) ? Request["city"].ToString().Trim() : "";
            string rcType = (Request["rcType"] != null) ? Request["rcType"].ToString().Trim() : "";

            int pageEnd = (int.Parse(PageNo) + 1) * PageSize;
            int pageStart = pageEnd - PageSize + 1;

            DataSet dt = om_db.MonthList(pageStart.ToString(), pageEnd.ToString(), city, rcType);
            string xmlstr = string.Empty;
            string totalxml = "<total>" + dt.Tables[0].Rows[0]["total"].ToString() + "</total>";
            xmlstr = DataTableToXml.ConvertDatatableToXML(dt.Tables[1], "dataList", "data_item");
            xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + totalxml + xmlstr + "</root>";
            xDoc.LoadXml(xmlstr);
        }
        catch (Exception ex)
        {
            xDoc = ExceptionUtil.GetExceptionDocument(ex);
        }
        Response.ContentType = System.Net.Mime.MediaTypeNames.Text.Xml;
        xDoc.Save(Response.Output);
    }
}