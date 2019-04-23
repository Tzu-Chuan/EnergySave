using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Xml;

public partial class handler_getNews : System.Web.UI.Page
{
    News_DB n_db = new News_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        ///-----------------------------------------------------
        ///功    能: 撈公告列表
        ///說明:
        /// * Request["SearchDate"]: 日期
        /// * Request["SearchStr"]: 關鍵字
        /// * Request["PageNo"]:欲顯示的頁碼, 由零開始
        /// * Request["PageSize"]:每頁顯示的資料筆數, 未指定預設10
        ///-----------------------------------------------------
        XmlDocument xDoc = new XmlDocument();
        try
        {
            string PageNo = (Request["CurrentPage"] != null) ? Request["CurrentPage"].ToString().Trim() : "";
            int PageSize = (Request["PageSize"] != null) ? int.Parse(Request["PageSize"].ToString().Trim()) : 10;
            int pageEnd = (int.Parse(PageNo) + 1) * PageSize;
            int pageStart = pageEnd - PageSize + 1;

            string SearchDate = (Request["SearchDate"] != null) ? Request["SearchDate"].ToString().Trim() : "";
            string SearchStr = (Request["SearchStr"] != null) ? Request["SearchStr"].ToString().Trim() : "";

            n_db._N_Date = SearchDate;
            n_db._strKeyWord = SearchStr;
            DataSet ds = n_db.getNewsList(pageStart.ToString(), pageEnd.ToString());

            string xmlstr = string.Empty;
            string xmlstr1 = string.Empty;
            string xmlstr2 = string.Empty;
            xmlstr1 = "<total>" + ds.Tables[0].Rows[0]["total"].ToString() + "</total>";
            xmlstr2 = DataTableToXml.ConvertDatatableToXML(ds.Tables[1], "dataList", "data_item");
            xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + xmlstr1 + xmlstr2 + "</root>";
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