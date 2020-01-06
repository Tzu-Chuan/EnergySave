using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Configuration;
using System.Drawing;

public partial class handler_ExportReportMonthEx : System.Web.UI.Page
{
    ReportMonth_DB rm = new ReportMonth_DB();
    Member_DB mb = new Member_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        string period1 = string.Empty;
        string period2 = string.Empty;
        string period3 = string.Empty;

        //string newName = Guid.NewGuid().ToString("N").Substring(0, 10) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
        //string FileName = string.Empty;
        string newName = Guid.NewGuid().ToString("N").Substring(0, 10) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

        string FileName = newName;
        Aspose.Words.Document doc = new Aspose.Words.Document(Server.MapPath("~/Template/ReportMonthEx.docx"));
        DocumentBuilder builder = new DocumentBuilder(doc);
        if (Request.QueryString["v"] != null)
        {
            //副檔名
            string ename = (Request.QueryString["tp"].ToString() == "WORD") ? ".docx" : ".pdf";

            string fileCity = "", fileYear = "", fileMonth = "";
            string strHtml = "";
            DataTable dt = new DataTable();
            string rmguid = string.IsNullOrEmpty(Request.QueryString["v"]) ? "" : Common.Decrypt(Request.QueryString["v"].ToString());
            //mrguid = "ba8aec0a63d2462ab49759731ee0bfd1";
            rm._RM_ReportGuid = rmguid;
            dt = rm.selectMonthReportDetailEx();

            if (dt.Rows.Count > 0)
            {
                
                strHtml = getHtml(dt);
                string strYear = "";
                strYear = (dt.Rows[0]["RM_Year"] == null || dt.Rows[0]["RM_Year"].ToString().Trim() == "") ? "0" : dt.Rows[0]["RM_Year"].ToString().Trim();
                
                builder.MoveToDocumentEnd();//從最下面開始放
                builder.InsertHtml(strHtml);
                builder.Write("\n");
            }

            fileCity = getfileCity(dt);
            fileYear = getfileYear(dt);
            fileMonth = getfileMonth(dt);
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
            {
                run.Font.Name = "新細明體";
                run.Font.Size = 12;
            }
            doc.Save(Server.MapPath("~/Template/" + newName + ename));
            Response.Clear();
            Response.ClearHeaders();

            FileName = fileCity + "_" + fileYear + "年" + fileMonth + "月擴大補助月報";
            string BrowserName = Request.Browser.Browser.ToLower();
            FileName = (BrowserName != "firefox") ? Server.UrlEncode(FileName + ename) : FileName + ename; // firefox 就愛跟別人不一樣
            Response.AddHeader("content-disposition", "attachment;filename=" + FileName);
            Response.ContentType = "application/octet-stream";
            Response.WriteFile(Server.MapPath("~/Template/" + newName + ename));
            Response.Flush();
            File.Delete(Server.MapPath("~/Template/" + newName + ename));
            Response.End();

        }
    }
    private string monthDiff(string startM, string endM)
    {
        string tmpstr = string.Empty;
        DateTime dt1 = Convert.ToDateTime(startM);
        DateTime dt2 = Convert.ToDateTime(endM);

        int Year = dt2.Year - dt1.Year;
        int Month = (dt2.Year - dt1.Year) * 12 + (dt2.Month - dt1.Month) + 1;

        return Month.ToString();
    }

    
    private string getHtml(DataTable dt)
    {
        string strHtml = "";
        string strItem, type1, type2, ItemCname, ptno, strUnit="";
        string[] P_ExFinish, splitRM_Finish, splitRM_Finish01;
        string strP_ExFinish, strRM_Finish, strRM_Finish01;
        string strcountApplyKW, strcountApply01;
        string strYear = "", strSY = "", strSM = "", strSD = "", strEY = "", strEM = "", strED = "", strS = "", strE = "";
        string strPeopleName = "", stePeoplePhone = "", strManager = "", strChkDate = "", strCreateDate = "";
        int rowSpan = 0;

        strYear = dt.Rows[0]["RM_Year"].ToString().Trim();
        rowSpan = (dt.Rows.Count) * 8;

        if (dt.Rows[0]["RM_Stage"].ToString().Trim() == "1")
        {
            strS = dt.Rows[0]["I_1_Sdate"].ToString().Trim();
            strE = dt.Rows[0]["I_1_Edate"].ToString().Trim();
            strSY = (Convert.ToDateTime(strS).Year - 1911).ToString();
            strSM = Convert.ToDateTime(strS).Month.ToString();
            strSD = Convert.ToDateTime(strS).Day.ToString();
            strEY = Convert.ToDateTime(strE).Year.ToString();
            strEM = Convert.ToDateTime(strE).Month.ToString();
            strED = Convert.ToDateTime(strE).Day.ToString();
        }
        if (dt.Rows[0]["RM_Stage"].ToString().Trim() == "2")
        {
            strS = dt.Rows[0]["I_2_Sdate"].ToString().Trim();
            strE = dt.Rows[0]["I_2_Edate"].ToString().Trim();
            strSY = (Convert.ToDateTime(strS).Year - 1911).ToString();
            strSM = Convert.ToDateTime(strS).Month.ToString();
            strSD = Convert.ToDateTime(strS).Day.ToString();
            strEY = Convert.ToDateTime(strE).Year.ToString();
            strEM = Convert.ToDateTime(strE).Month.ToString();
            strED = Convert.ToDateTime(strE).Day.ToString();
        }
        if (dt.Rows[0]["RM_Stage"].ToString().Trim() == "3")
        {
            strS = dt.Rows[0]["I_3_Sdate"].ToString().Trim();
            strE = dt.Rows[0]["I_3_Edate"].ToString().Trim();
            strSY = (Convert.ToDateTime(strS).Year - 1911).ToString();
            strSM = Convert.ToDateTime(strS).Month.ToString();
            strSD = Convert.ToDateTime(strS).Day.ToString();
            strEY = Convert.ToDateTime(strE).Year.ToString();
            strEM = Convert.ToDateTime(strE).Month.ToString();
            strED = Convert.ToDateTime(strE).Day.ToString();
        }

        
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            
            if (i == 0)
            {
                strHtml += "<table border='1' cellpadding='0' cellspacing ='0' width='100%' >";
                strHtml += "<tr>";
                strHtml += "<td colspan='2' align='center'>執行機關</td><td colspan='3' align='center'>" + dt.Rows[0]["cityname"].ToString().Trim() + "</td>";
                strHtml += "<td colspan='3' align='center'>承辦局處</td><td colspan='6' align='center'>" + dt.Rows[0]["M_Office"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td colspan='2' align='center' valign='center'>期數</td><td colspan ='12'>第" + dt.Rows[0]["RM_Stage"].ToString().Trim() + "期 自" + strSY + "年" + strSM + "月" + strSD + "日起至" + strEY + "年" + strEM + "月" + strED + "日止，共" + monthDiff(strS, strE) + "個月</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td colspan='2' align='center'>提報月份</td><td colspan='12'>" + (Convert.ToInt32(strYear) - 1911).ToString() + "年" + dt.Rows[0]["RM_Month"].ToString().Trim() + "月</td>";
                strHtml += "</tr>";

                strPeopleName = dt.Rows[0]["M_Name"].ToString().Trim();
                stePeoplePhone = dt.Rows[0]["M_Tel"].ToString().Trim();
                strManager = dt.Rows[0]["bossname"].ToString().Trim();
                if (dt.Rows[0]["RC_CheckDate"].ToString() == "")
                {
                    strChkDate = DateTime.Now.ToString("yyyy/MM/dd");
                }
                else {
                    strChkDate = Convert.ToDateTime(dt.Rows[0]["RC_CheckDate"].ToString()).ToString("yyyy/MM/dd");
                }
                
                strCreateDate = Convert.ToDateTime(dt.Rows[0]["RM_ModDate"].ToString()).ToString("yyyy/MM/dd");
            }


            strItem = dt.Rows[i]["P_ItemName"].ToString().Trim();
            type1 = dt.Rows[i]["P_ExType"].ToString().Trim();
            type2 = dt.Rows[i]["P_ExDeviceType"].ToString().Trim();
            ItemCname = dt.Rows[i]["P_ItemName_c"].ToString().Trim();
            ptno = dt.Rows[i]["P_ItemName"].ToString().Trim();

            P_ExFinish = dt.Rows[i]["P_ExFinish"].ToString().Trim().Split('.');
            splitRM_Finish = dt.Rows[i]["RM_Finish"].ToString().Trim().Split('.');
            splitRM_Finish01 = dt.Rows[i]["RM_Finish01"].ToString().Trim().Split('.');
            strP_ExFinish = P_ExFinish[0];
            strRM_Finish = splitRM_Finish[0];
            strRM_Finish01 = splitRM_Finish01[0];

            strcountApplyKW = dt.Rows[i]["countApplyKW"].ToString().Trim();//累計申請數KW
            strcountApply01 = dt.Rows[i]["countApply01"].ToString().Trim();//累計申請數

            if (strItem == "99")
            {
                ptno = ptno + dt.Rows[i]["P_ExType"].ToString().Trim() + dt.Rows[i]["P_ExDeviceType"].ToString().Trim();
                ItemCname += "<br />" + dt.Rows[i]["P_ExType_c"].ToString().Trim() + "-" + dt.Rows[i]["P_ExDeviceType_c"].ToString().Trim();
            }


            if (strItem == "01" || strItem == "23" || strItem == "33")
            {
                strUnit = "KW";
            }
            if (strItem == "02" || strItem == "21")
            {
                strUnit = "具";
            }
            if (strItem == "03" || strItem == "22" || strItem == "29")
            {
                strUnit = "盞";
            }
            if (strItem == "04" || strItem == "26" || strItem == "30" || strItem == "31")
            {
                strUnit = "套";
            }
            if (strItem == "05" || strItem == "06" || strItem == "07" || strItem == "08" || strItem == "09" || strItem == "10" || strItem == "11" || strItem == "12" || strItem == "13" || strItem == "14" || strItem == "15" || strItem == "19" || strItem == "20")
            {
                strUnit = "台";
            }
            if (strItem == "17" || strItem == "18" || strItem == "24" || strItem == "28" || strItem == "32")
            {
                strUnit = "顆";
            }
            if (strItem == "27")
            {
                strUnit = "噸";
            }
            if (strItem == "25")
            {
                strUnit = "個";
            }
            if (strItem == "16")
            {
                strUnit = "10顆一單位";
            }


            if (strItem == "01" || strItem == "33")
            {//type1 == "02" && type2 == "01" =>空調

                if (i == 0)
                {
                    strHtml += "<tr><td rowspan='" + rowSpan + "' align='center' width = '5%'>本<br>月<br>執<br>行<br>進<br>度</td>";
                    strHtml += "<td rowspan='8' width='11%'>" + ItemCname + "</td>";
                    strHtml += "<td colspan='6'>本期累計核定數：" + strRM_Finish01 + "&nbsp;(台)</td>";
                    strHtml += "<td colspan='6'>本期規劃數：" + strP_ExFinish + "&nbsp;(KW)</td>";
                    strHtml += "</tr>";
                    //strHtml += "<td colspan='6'>本期累計申請數：" + strcountApplyKW + "</td>";
                    //strHtml += "<td>本期累計完成數："+ strRM_Finish + "&nbsp;(kW)</td>";
                }
                else {
                    strHtml += "<tr>";
                    strHtml += "<td rowspan='8'>" + ItemCname + "</td>";
                    strHtml += "<td colspan='6'>本期累計核定數：" + strRM_Finish01 + "&nbsp;(台)</td>";
                    strHtml += "<td colspan='6'>本期規劃數：" + strP_ExFinish + "&nbsp;(台)</td>";
                    strHtml += "</tr>";
                    //strHtml += "<td colspan='6'>本期累計申請數：" + strcountApplyKW + "</td>";
                    //strHtml += "<td>本期累計完成數："+ strRM_Finish + "&nbsp;(kW)</td>";
                }
                
                strHtml += "<tr>";
                strHtml += "<th colspan='3' width='21%'>本月申請數量(台)</th>";
                strHtml += "<th colspan='3' width='21%'>本月核定數量(台)</th>";
                strHtml += "<th colspan='3' width='21%'>本月申請總冷氣能力(kW)</th>";
                strHtml += "<th colspan='3' width='21%'>本月完成總冷氣能力(kW)</th>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='center'>服務業</td>";
                strHtml += "<td align='center'>機關學校</td>";
                strHtml += "<td align='center'>住宅</td>";
                strHtml += "<td align='center'>服務業</td>";
                strHtml += "<td align='center'>機關學校</td>";
                strHtml += "<td align='center'>住宅</td>";
                strHtml += "<td align='center'>服務業</td>";
                strHtml += "<td align='center'>機關學校</td>";
                strHtml += "<td align='center'>住宅</td>";
                strHtml += "<td align='center'>服務業</td>";
                strHtml += "<td align='center'>機關學校</td>";
                strHtml += "<td align='center'>住宅</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type1Value1"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type1Value2"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type1Value3"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type2Value1"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type2Value2"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type2Value3"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type3Value1"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type3Value2"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type3Value3"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type4Value1"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type4Value2"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type4Value3"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='right'>合計</td>";
                strHtml += "<td colspan='2' align='right'>"+ dt.Rows[i]["RM_Type1ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type2ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type3ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type4ValueSum"].ToString().Trim() + "</td>";
                strHtml += "</tr>";

                strHtml += "<tr>";
                strHtml += "<th colspan='4' align='center'>本月申請數預期年節電量(度)</th>";
                strHtml += "<th colspan='4' align='center'>本月核定數預期年節電量(度)</th>";
                strHtml += "<th colspan='4' align='center'>本月未核定數之年節電量(度)</th>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td colspan='4' align='right'>" + dt.Rows[i]["RM_PreVal"].ToString().Trim() + "</td>";
                strHtml += "<td colspan='4' align='right'>" + dt.Rows[i]["RM_ChkVal"].ToString().Trim() + "</td>";
                strHtml += "<td colspan='4' align='right'>" + dt.Rows[i]["RM_NotChkVal"].ToString().Trim() + "</td>";
                strHtml+= "</tr>";
                strHtml+= "<tr>";
                strHtml += "<td colspan='12'>補充說明：" + dt.Rows[i]["RM_Remark"].ToString().Trim() + "</td>";
                strHtml+= "</tr>";
            }
            else if (strItem == "02" || strItem == "03")
            {
                if (i == 0)
                {
                    strHtml += "<tr><td rowspan='" + rowSpan + "' align='center' width = '5%'>本<br>月<br>執<br>行<br>進<br>度</td>";
                    strHtml += "<td rowspan='8' width='11%'>" + ItemCname + "</td>";
                    strHtml += "<td colspan='6'>本期累計核定數：" + dt.Rows[i]["countFinish02"].ToString().Trim() + "&nbsp;(" + strUnit + "))</td>";
                    strHtml += "<td colspan='6'>本期規劃數：" + strP_ExFinish + "&nbsp;(" + strUnit + "))</td>";
                    strHtml += "</tr>";
                }
                else
                {
                    strHtml += "<tr>";
                    strHtml += "<td rowspan='8'>" + ItemCname + "</td>";
                    strHtml += "<td colspan='6'>本期累計核定數：" + dt.Rows[i]["countFinish02"].ToString().Trim() + "&nbsp;(" + strUnit + "))</td>";
                    strHtml += "<td colspan='6'>本期規劃數：" + strP_ExFinish + "&nbsp;(" + strUnit + "))</td>";
                    strHtml += "</tr>";
                }
                
                strHtml += "<tr>";
                if (strItem == "03") {
                    strHtml += "<th colspan='3' width='21%'>本月申請數量(" + strUnit + ")</th>";
                    strHtml += "<th colspan='3' width='21%'>本月核定數量(" + strUnit + ")</th>";
                }
                else {
                    //02
                    strHtml += "<th colspan='3' width='21%'>本月申請數量(" + strUnit + ")</th>";
                    strHtml += "<th colspan='3' width='21%'>本月核定數量(" + strUnit + ")</th>";
                }
                strHtml += "<th colspan='3' width='21%'>本月申請更換照明瓦數(W)</th>";
                strHtml += "<th colspan='3' width='21%'>本月完成更換照明瓦數(W)</th>";

                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='center'>服務業</td>";
                strHtml += "<td align='center'>機關學校</td>";
                strHtml += "<td align='center'>住宅</td>";
                strHtml += "<td align='center'>服務業</td>";
                strHtml += "<td align='center'>機關學校</td>";
                strHtml += "<td align='center'>住宅</td>";
                strHtml += "<td align='center'>服務業</td>";
                strHtml += "<td align='center'>機關學校</td>";
                strHtml += "<td align='center'>住宅</td>";
                strHtml += "<td align='center'>服務業</td>";
                strHtml += "<td align='center'>機關學校</td>";
                strHtml += "<td align='center'>住宅</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type1Value1"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type1Value2"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type1Value3"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type2Value1"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type2Value2"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type2Value3"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type3Value1"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type3Value2"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type3Value3"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type4Value1"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type4Value2"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type4Value3"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='right'>合計</td>";
                strHtml += "<td colspan='2' align='right'>"+ dt.Rows[i]["RM_Type1ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type2ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type3ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type4ValueSum"].ToString().Trim() + "</td>";
                strHtml += "</tr>";

                strHtml += "<tr>";
                strHtml += "<th colspan='4' align='center'>本月申請數預期年節電量(度)</th>";
                strHtml += "<th colspan='4' align='center'>本月核定數預期年節電量(度)</th>";
                strHtml += "<th colspan='4' align='center'>本月未核定數之年節電量(度)</th>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td colspan='4' align='right'>" + dt.Rows[i]["RM_PreVal"].ToString().Trim() + "</td>";
                strHtml += "<td colspan='4' align='right'>" + dt.Rows[i]["RM_ChkVal"].ToString().Trim() + "</td>";
                strHtml += "<td colspan='4' align='right'>" + dt.Rows[i]["RM_NotChkVal"].ToString().Trim() + "</td>";
                strHtml+= "</tr>";
                strHtml+= "<tr>";
                strHtml += "<td colspan='12'>補充說明：" + dt.Rows[i]["RM_Remark"].ToString().Trim() + "</td>";
                strHtml+= "</tr>";
            }
            else
            {
                //if (strItem == "05")
                //{
                //    strUnit = "KW";
                //}
                //else if (strItem == "14")
                //{
                //    strUnit = "組";
                //}
                //else
                //{
                //    strUnit = "台";
                //}
                if (i == 0)
                {
                    strHtml += "<tr><td rowspan='" + rowSpan + "' align='center' width = '5%'>本<br>月<br>執<br>行<br>進<br>度</td>";
                    strHtml += "<td rowspan='8' width='11%'>" + ItemCname + "</td>";
                    strHtml += "<td colspan='6'>本期累計完成數：" + dt.Rows[i]["countFinish03"].ToString().Trim() + "&nbsp;(" + strUnit + ")</td>";
                    strHtml += "<td colspan='6'>本期規劃數：" + strP_ExFinish + "&nbsp;(" + strUnit + ")</td>";
                    strHtml += "</tr>";
                }
                else
                {
                    strHtml += "<tr>";
                    strHtml += "<td rowspan='8'>" + ItemCname + "</td>";
                    strHtml += "<td colspan='6'>本期累計完成數：" + dt.Rows[i]["countFinish03"].ToString().Trim() + "&nbsp;(" + strUnit + ")</td>";
                    strHtml += "<td colspan='6'>本期規劃數：" + strP_ExFinish + "&nbsp;(" + strUnit + ")</td>";
                    strHtml += "</tr>";
                    //strHtml += "<td colspan='6'>本期累計申請數：" + strcountApply01 + "&nbsp;(" + strUnit + ")</td>";
                    //strHtml += "<td colspan='6'>本期累計完成數：" + strRM_Finish + "&nbsp;(" + strUnit + ")</td>";
                }
                
                strHtml += "<tr>";
                strHtml += "<th colspan='3' valign='top'>本月申請數量(" + strUnit + ")</th>";
                strHtml += "<th colspan='3' valign='top'>本月核定數量(" + strUnit + ")</th>";
                strHtml += "<th colspan='3' valign='top'>本月完成數量(" + strUnit + ")</th>";
                strHtml += "<th colspan='3' valign='top'></th>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='center'>服務業</td>";
                strHtml += "<td align='center'>機關學校</td>";
                strHtml += "<td align='center'>住宅</td>";
                strHtml += "<td align='center'>服務業</td>";
                strHtml += "<td align='center'>機關學校</td>";
                strHtml += "<td align='center'>住宅</td>";
                strHtml += "<td align='center'>服務業</td>";
                strHtml += "<td align='center'>機關學校</td>";
                strHtml += "<td align='center'>住宅</td>";
                strHtml += "<td colspan='3'></td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type1Value1"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type1Value2"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type1Value3"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type2Value1"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type2Value2"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type2Value3"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type3Value1"].ToString().Trim().Replace(".0", "") + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type3Value2"].ToString().Trim().Replace(".0", "") + "</td>";
                strHtml += "<td align='right'>" + dt.Rows[i]["RM_Type3Value3"].ToString().Trim().Replace(".0", "") + "</td>";
                strHtml += "<td colspan='3'></td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='right'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type1ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type2ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td align='right'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type3ValueSum"].ToString().Trim().Replace(".0", "") + "</td>";
                strHtml += "<td colspan='3'></td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<th colspan='3'>本月申請數預期年節電量(度)</th>";
                strHtml += "<th colspan='3'>本月核定數預期年節電量(度)</th>";
                strHtml += "<th colspan='3'>本月未核定數之年節電量(度)</th>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td colspan='3' align='right'>" + dt.Rows[i]["RM_PreVal"].ToString().Trim() + "</td>";
                strHtml += "<td colspan='3' align='right'>" + dt.Rows[i]["RM_ChkVal"].ToString().Trim() + "</td>";
                strHtml += "<td colspan='3' align='right'>" + dt.Rows[i]["RM_NotChkVal"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td colspan='9'>補充說明：" + dt.Rows[i]["RM_Remark"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
            }
        }
        strHtml += "<tr>";
        strHtml += "<td colspan='7'>填表人：" + strPeopleName + "</td>";
        strHtml += "<td colspan='7' rowspan='2'>主管：" + strManager + "</td>";
        strHtml += "</tr>";
        strHtml += "<tr>";
        strHtml += "<td colspan='7'>電話：" + stePeoplePhone + "</td>";
        strHtml += "</tr>";
        strHtml += "<tr>";
        strHtml += "<td colspan='7'>填表日期：" + strCreateDate + "</td>";
        strHtml += "<td colspan='7'>簽核日期：" + strChkDate + "</td>";
        strHtml += "</tr>";
        strHtml += "</table>";
        
        return strHtml;
    }

    //取得 檔機關
    private string getfileCity(DataTable dt)
    {
        string fileCity = "";
        fileCity = dt.Rows[0]["cityname"].ToString().Trim();
        return fileCity;
    }
    //取得 檔案年
    private string getfileYear(DataTable dt)
    {
        string strYear = "", fileYear = "";
        strYear = (dt.Rows[0]["RM_Year"] == null || dt.Rows[0]["RM_Year"].ToString().Trim() == "") ? "0" : dt.Rows[0]["RM_Year"].ToString().Trim();
        fileYear = (Convert.ToInt32(strYear) - 1911).ToString();
        return fileYear;
    }
    //取得 檔案月
    private string getfileMonth(DataTable dt)
    {
        string fileMonth = "";
        fileMonth = dt.Rows[0]["RM_Month"].ToString().Trim();
        return fileMonth;
    }

    //項目名稱斷行
    private string ItemLineBreak(string str)
    {
        string recVal = string.Empty;
        for (int i = 0; i < str.Length; i++)
        {
            if (i % 5 == 0)
                recVal += "<br>";
            recVal += str.Substring(i, 1);
        }
        return recVal;
    }


    //補充說明斷行
    private string PsLineBreak(string str)
    {
        string recVal = string.Empty;
        for (int i = 0; i < str.Length; i++)
        {
            if (i % 30 == 0)
                recVal += "<br>";
            recVal += str.Substring(i, 1);
        }
        return recVal;
    }
}