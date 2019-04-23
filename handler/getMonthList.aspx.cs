using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Xml;

public partial class handler_getMonthList : System.Web.UI.Page
{
    ReportMonth_DB rm_db = new ReportMonth_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        ///-----------------------------------------------------
        ///功    能: 查詢月報列表
        ///說明:
        /// * Request["PageNo"]:欲顯示的頁碼, 由零開始
        /// * Request["PageSize"]:每頁顯示的資料筆數, 未指定預設10
        ///-----------------------------------------------------
        XmlDocument xDoc = new XmlDocument();
        try
        {
            string PageNo = (Request["PageNo"] != null) ? Request["PageNo"].ToString().Trim() : "";
            int PageSize = (Request["PageSize"] != null) ? int.Parse(Request["PageSize"].ToString().Trim()) : 10;
            string mtype = (Request["mtype"] != null) ? Request["mtype"].ToString().Trim() : "";
            string year = (Request["year"] != null) ? Request["year"].ToString().Trim() : "";
            string month = (Request["month"] != null) ? Request["month"].ToString().Trim() : "";
            string stage = (Request["stage"] != null) ? Request["stage"].ToString().Trim() : "";
            if (year!="") {
                year = (Convert.ToInt32(year) + 1911).ToString().Trim();
            }
            int pageEnd = (int.Parse(PageNo) + 1) * PageSize;
            int pageStart = pageEnd - PageSize + 1;

            DataSet ds = new DataSet();
            rm_db._RM_Stage = stage;
            rm_db._RM_Year = year;
            rm_db._RM_Month = month;
            rm_db._RM_ReportType = mtype;
            //設備汰換 因地制宜 共用
            ds = rm_db.getMonthList(LogInfo.mGuid, pageStart.ToString(), pageEnd.ToString());
            
            //rs_db._RS_Year = year;
            //rs_db._RS_Season = season;
            //rs_db._RS_Stage = stage;
            //DataSet dt = rs_db.getSeasonList(LogInfo.mGuid, pageStart.ToString(), pageEnd.ToString());
            string xmlstr = string.Empty;
            string xmlstr1 = string.Empty;
            if (ds.Tables.Count>0 && ds.Tables[1].Rows.Count > 0)
            {
                xmlstr1 += "<dataList>";
                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                {
                    xmlstr1 += "<data_item>";
                    xmlstr1 += "<itemNo>" + ds.Tables[1].Rows[i]["itemNo"].ToString() + "</itemNo>";
                    xmlstr1 += "<RM_ReportGuid>" + ds.Tables[1].Rows[i]["RM_ReportGuid"].ToString() + "</RM_ReportGuid>";
                    xmlstr1 += "<enGuid>" + Server.UrlEncode(Common.Encrypt(ds.Tables[1].Rows[i]["RM_ReportGuid"].ToString())) + "</enGuid>";
                    xmlstr1 += "<RM_Stage>" + ds.Tables[1].Rows[i]["RM_Stage"].ToString() + "</RM_Stage>";
                    xmlstr1 += "<RM_Year>" + ds.Tables[1].Rows[i]["RM_Year"].ToString() + "</RM_Year>";
                    xmlstr1 += "<RM_Month>" + ds.Tables[1].Rows[i]["RM_Month"].ToString() + "</RM_Month>";
                    xmlstr1 += "<RM_CreateDate>" + Convert.ToDateTime(ds.Tables[1].Rows[i]["RM_CreateDate"].ToString()).ToString("o") + "</RM_CreateDate>";
                    xmlstr1 += "<RC_CheckType>" + ds.Tables[1].Rows[i]["RC_CheckType"].ToString() + "</RC_CheckType>";
                    xmlstr1 += "<RC_Status>" + ds.Tables[1].Rows[i]["RC_Status"].ToString() + "</RC_Status>";
                    xmlstr1 += "<MTypeNum>" + ds.Tables[1].Rows[i]["MTypeNum"].ToString() + "</MTypeNum>";
                    xmlstr1 += "<MTypeName>" + ds.Tables[1].Rows[i]["MTypeName"].ToString() + "</MTypeName>";
                    xmlstr1 += "</data_item>";
                }
                xmlstr1 += "</dataList>";

                string totalxml = "<total>" + ds.Tables[0].Rows[0]["total"].ToString() + "</total>";
                xmlstr = "<?xml version='1.0' encoding='utf-8'?><root>" + totalxml + xmlstr1 + "</root>";
            }
            else {
                xmlstr = "<?xml version='1.0' encoding='utf-8'?><root></root>";
            }
            
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