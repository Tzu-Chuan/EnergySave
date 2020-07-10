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

public partial class handler_ExportReportMonth : System.Web.UI.Page
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
        Aspose.Words.Document doc = new Aspose.Words.Document(Server.MapPath("~/Template/ReportMonth.docx"));
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
            dt = rm.selectMonthReportDetail();
            string chkdatatype123 = "";
            string chkdatatype45 = "";
            if (dt.Rows.Count > 0)
            {
                for (int k = 0; k < dt.Rows.Count; k++)
                {
                    if (dt.Rows[k]["RM_CPType"].ToString().Trim() == "01" || dt.Rows[k]["RM_CPType"].ToString().Trim() == "02" || dt.Rows[k]["RM_CPType"].ToString().Trim() == "03")
                    {
                        chkdatatype123 = "Y";
                    }
                    if (dt.Rows[k]["RM_CPType"].ToString().Trim() == "04" || dt.Rows[k]["RM_CPType"].ToString().Trim() == "05")
                    {
                        chkdatatype45 = "Y";
                    }
                }
                string strYear = "";
                strYear = (dt.Rows[0]["RM_Year"] == null || dt.Rows[0]["RM_Year"].ToString().Trim() == "") ? "0" : dt.Rows[0]["RM_Year"].ToString().Trim();
                //同時有01~03 || 04~05的資料
                if (chkdatatype123 == "Y" && chkdatatype45 == "Y")
                {
                    strHtml = html1_5(strYear, strHtml, dt);
                }
                //只有01~03資料
                if (chkdatatype123 == "Y" && chkdatatype45 == "")
                {
                    strHtml = html123(strYear, strHtml, dt);
                }
                //只有04~05資料
                if (chkdatatype123 == "" && chkdatatype45 == "Y")
                {
                    strHtml = html45(strYear, strHtml, dt);
                }

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

            FileName = fileCity + "_" + fileYear + "年" + fileMonth + "月月報";
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

    //包含01~03其中一種  && 04 05 (中型 大型)其中一種
    private string html1_5(string strYear,string strHtml,DataTable dt)
    {
        string strPeopleName = "", stePeoplePhone = "", strManager = "", strChkDate = "", strCreateDate = "";
        string thName1 = "", thName2 = "", thName3 = "";
        string strU = "", strAppTitle = "", strFinishTitle = "";
        string strS = "", strE = "", strSY = "", strSM = "", strSD = "", strEY = "", strEM = "", strED = "";

        int CPTypeCount = 0;
        int rowSpan = 0;
        DataTable dtGroup = dt.DefaultView.ToTable(true, "RM_CPType");
        CPTypeCount = dtGroup.Rows.Count;
        rowSpan = CPTypeCount * 8;
        strHtml += "<table border='1' cellpadding='0' cellspacing='0' width='100%'>";
        strHtml += "<tr>";
        strHtml += "<td colspan='2' align='center'>執行機關</td><td colspan='3' align='center'>" + dt.Rows[0]["cityname"].ToString().Trim() + "</td>";
        strHtml += "<td colspan='3' align='center'>承辦局處</td><td colspan='6'>" + dt.Rows[0]["M_Office"].ToString().Trim() + "</td>";
        strHtml += "</tr>";
        strHtml += "<tr>";
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
        strHtml += "<td colspan='2' align='center' valign='center'>期數</td><td colspan ='12'>第" + dt.Rows[0]["RM_Stage"].ToString().Trim() + "期 自" + strSY + "年" + strSM + "月" + strSD + "日起至" + strEY + "年" + strEM + "月" + strED + "日止，共" + monthDiff(strS, strE) + "個月</td>";
        strHtml += "</tr>";
        strHtml += "<tr>";
        strHtml += "<td colspan='2' align='center'>提報月份</td><td colspan='12'>" + (Convert.ToInt32(strYear) - 1911).ToString() + "年" + dt.Rows[0]["RM_Month"].ToString().Trim() + "月</td>";
        strHtml += "</tr>";

        for (int i = 0; i < dt.Rows.Count; i++)
        //for (int i = 0; i < 1; i++)
        {

            if (dt.Rows[i]["RM_CPType"].ToString().Trim() != "03")
            {
                thName1 = "機關";
                thName2 = "學校";
                thName3 = "服務<br>業";
            }
            else
            {//室內停車場智慧照明
                thName1 = "集合<br>住宅 ";
                thName2 = "辦公<br>大樓";
                thName3 = "服務<br>業";
            }
            strHtml += "<tr>";
            if (i == 0)
            {
                strPeopleName = dt.Rows[0]["M_Name"].ToString().Trim();
                stePeoplePhone = dt.Rows[0]["M_Tel"].ToString().Trim();
                strManager = dt.Rows[0]["bossname"].ToString().Trim();
                if (dt.Rows[0]["RC_CheckDate"].ToString() != "")
                    strChkDate = Convert.ToDateTime(dt.Rows[0]["RC_CheckDate"].ToString()).ToString("yyyy/MM/dd");
                strCreateDate = Convert.ToDateTime(dt.Rows[0]["RM_ModDate"].ToString()).ToString("yyyy/MM/dd");
                strHtml += "<td rowspan='" + rowSpan + "' align='center' width = '5%'>本<br>月<br>執<br>行<br>進<br>度</td>";
            }
            if (dt.Rows[i]["RM_CPType"].ToString().Trim() != "04" && dt.Rows[i]["RM_CPType"].ToString().Trim() != "05")
            {
                if (dt.Rows[i]["RM_CPType"].ToString().Trim() == "01")
                {
                    strU = "(KW)";
                    strAppTitle = "本月申請總冷<br>氣能力(KW)";
                    strFinishTitle = "本月完成總冷<br>氣能力(kW)";
                }
                if (dt.Rows[i]["RM_CPType"].ToString().Trim() == "02")
                {
                    strU = "(具)";
                    strAppTitle = "本月申請更換<br>照明瓦數(W)";
                    strFinishTitle = "本月完成更換<br>照明瓦數(W)";
                }
                if (dt.Rows[i]["RM_CPType"].ToString().Trim() == "03")
                {
                    strU = "(盞)";
                    strAppTitle = "本月申請更換<br>照明瓦數(W)";
                    strFinishTitle = "本月完成更換<br>照明瓦數(W)";
                }

                strHtml += "<td rowspan='8' width='15%'>" + ItemLineBreak(dt.Rows[i]["C_Item_cn"].ToString().Trim()) + "</td>";
                strHtml += "<td colspan='6'>本期累計核定數：" + dt.Rows[i]["RM_Finish"].ToString().Trim() + strU + "</td>";
                strHtml += "<td colspan='6'>本期規劃數：" + dt.Rows[i]["RM_Planning"].ToString().Trim() + strU + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                if (dt.Rows[i]["RM_CPType"].ToString().Trim() == "01")
                {
                    strHtml += "<td colspan='3' align='center' width='20%'>本月申請數量(台)</td>";
                    strHtml += "<td colspan='3' align='center' width='20%'>本月核定數量(台)</td>";
                }
                else {
                    strHtml += "<td colspan='3' align='center' width='20%'>本月申請數量" + strU + "</td>";
                    strHtml += "<td colspan='3' align='center' width='20%'>本月核定數量" + strU + "</td>";
                }
                
                strHtml += "<td colspan='3' align='center' width='20%'>" + strAppTitle + "</td>";
                strHtml += "<td colspan='3' align='center' width='20%'>" + strFinishTitle + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='center'>" + thName1 + "</td>";
                strHtml += "<td align='center'>" + thName2 + "</td>";
                strHtml += "<td align='center'>" + thName3 + "</td>";
                strHtml += "<td align='center'>" + thName1 + "</td>";
                strHtml += "<td align='center'>" + thName2 + "</td>";
                strHtml += "<td align='center'>" + thName3 + "</td>";
                strHtml += "<td align='center'>" + thName1 + "</td>";
                strHtml += "<td align='center'>" + thName2 + "</td>";
                strHtml += "<td align='center'>" + thName3 + "</td>";
                strHtml += "<td align='center'>" + thName1 + "</td>";
                strHtml += "<td align='center'>" + thName2 + "</td>";
                strHtml += "<td align='center'>" + thName3 + "</td>";
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
                strHtml += "<td align='center'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type1ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td align='center'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type2ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td align='center'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type3ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td align='center'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type4ValueSum"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td colspan='3' align='center'>申請數預期年節電量(度)</td>";
                strHtml += "<td colspan='3' align='center'>核定數預期年節電量(度)</td>";
                strHtml += "<td colspan='6' align='center'>未核定數之年節電量(度) </td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td colspan='3' align='right'>" + dt.Rows[i]["RM_PreVal"].ToString().Trim() + "</td>";
                strHtml += "<td colspan='3' align='right'>" + dt.Rows[i]["RM_ChkVal"].ToString().Trim() + "</td>";
                strHtml += "<td colspan='6' align='right'>" + dt.Rows[i]["RM_NotChkVal"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                if (dt.Rows[i]["RM_CPType"].ToString().Trim() == "02")
                {
                    strHtml += "<td colspan='12'> 補充說明：(T8/T9)" + PsLineBreak(dt.Rows[i]["RM_Remark"].ToString().Trim()) + "</td>";
                }
                else
                {
                    strHtml += "<td colspan='12'> 補充說明：" + PsLineBreak(dt.Rows[i]["RM_Remark"].ToString().Trim()) + "</td>";
                }
                strHtml += "</tr>";
            }
            else
            {
                strHtml += "<td rowspan='8'>" + ItemLineBreak(dt.Rows[i]["C_Item_cn"].ToString().Trim()) + "</td>";
                strHtml += "<td colspan='6'>本期累計完成數：" + dt.Rows[i]["RM_Finish"].ToString().Trim() + "(套)</td>";
                strHtml += "<td colspan='6'>本期規劃數：" + dt.Rows[i]["RM_Planning"].ToString().Trim() + "(套)</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td colspan='3' align='center' width='20%'>申請數量(套)</td>";
                strHtml += "<td colspan='3' align='center' width='20%'>核定數量(套)</td>";
                strHtml += "<td colspan='3' align='center' width='20%'>完成數量(套)</td>";
                strHtml += "<td colspan='3' rowspan='4' width='20%'></td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td algn='center'>" + thName1 + "</td>";
                strHtml += "<td algn='center'>" + thName2 + "</td>";
                strHtml += "<td algn='center'>" + thName3 + "</td>";
                strHtml += "<td algn='center'>" + thName1 + "</td>";
                strHtml += "<td algn='center'>" + thName2 + "</td>";
                strHtml += "<td algn='center'>" + thName3 + "</td>";
                strHtml += "<td algn='center'>" + thName1 + "</td>";
                strHtml += "<td algn='center'>" + thName2 + "</td>";
                strHtml += "<td algn='center'>" + thName3 + "</td>";
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
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='center'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type1ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td align='center'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type2ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td align='center'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type3ValueSum"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='center' colspan='3'>申請數預期年節電量(度)</td>";
                strHtml += "<td align='center' colspan='3'>核定數預期年節電量(度)</td>";
                strHtml += "<td align='center' colspan='6'>未核定數之年節電量(度) </td>";
                strHtml += "</tr >";
                strHtml += "<tr>";
                strHtml += "<td align='right' colspan='3'>" + dt.Rows[i]["RM_PreVal"].ToString().Trim() + "</td>";
                strHtml += "<td align='right' colspan='3'>" + dt.Rows[i]["RM_ChkVal"].ToString().Trim() + "</td>";
                strHtml += "<td align='right' colspan='6'>" + dt.Rows[i]["RM_NotChkVal"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td colspan='12'> 補充說明：" + PsLineBreak(dt.Rows[i]["RM_Remark"].ToString().Trim()) + "</td>";
                strHtml += "</tr>";
            }

        }
        strHtml += "</table>";
        strHtml += "<table border='1' cellpadding='0' cellspacing ='0' width='100%'>";
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
        strHtml += "</tr></table>";
        //strPeopleName = dt.Rows[0]["M_Name"].ToString().Trim();
        //stePeoplePhone = dt.Rows[0]["M_Phone"].ToString().Trim();
        //strManager = dt.Rows[0]["bossname"].ToString().Trim();
        //strChkDate = Convert.ToDateTime(dt.Rows[0]["RC_CheckDate"].ToString()).ToString("yyyy/MM/dd");
        //strCreateDate = Convert.ToDateTime(dt.Rows[0]["RC_CreateDate"].ToString()).ToString("yyyy/MM/dd");

        return strHtml;
    }

    //只有01~03 沒有04~05
    private string html123(string strYear, string strHtml, DataTable dt)
    {
        string strPeopleName = "", stePeoplePhone = "", strManager = "", strChkDate = "", strCreateDate = "";
        string thName1 = "", thName2 = "", thName3 = "";
        string strU = "", strUSumAppl = "", strUSumFinish = "";
        string strS = "", strE = "", strSY = "", strSM = "", strSD = "", strEY = "", strEM = "", strED = "";

        int CPTypeCount = 0;
        int rowSpan = 0;
        DataTable dtGroup = dt.DefaultView.ToTable(true, "RM_CPType");
        CPTypeCount = dtGroup.Rows.Count;
        rowSpan = CPTypeCount * 8;

        strHtml += "<table border='1' cellpadding='0' cellspacing ='0' width='100%' >";
        strHtml += "<tr>";
        strHtml += "<td colspan='2' align='center'>執行機關</td><td colspan='3' align='center'>" + dt.Rows[0]["cityname"].ToString().Trim() + "</td>";
        strHtml += "<td colspan='3' align='center'>承辦局處</td><td colspan='6' align='center'>" + dt.Rows[0]["M_Office"].ToString().Trim() + "</td>";
        strHtml += "</tr>";
        strHtml += "<tr>";
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
        strHtml += "<td colspan='2' align='center' valign='center'>期數</td><td colspan ='12'>第" + dt.Rows[0]["RM_Stage"].ToString().Trim() + "期 自" + strSY + "年" + strSM + "月" + strSD + "日起至" + strEY + "年" + strEM + "月" + strED + "日止，共" + monthDiff(strS, strE) + "個月</td>";
        strHtml += "</tr>";
        strHtml += "<tr>";
        strHtml += "<td colspan='2' align='center'>提報月份</td><td colspan='12'>" + (Convert.ToInt32(strYear) - 1911).ToString() + "年" + dt.Rows[0]["RM_Month"].ToString().Trim() + "月</td>";
        strHtml += "</tr>";

        //for (int i = 0; i < 1; i++)
        for (int i = 0; i < dt.Rows.Count; i++)
        {

            if (dt.Rows[i]["RM_CPType"].ToString().Trim() != "03")
            {
                thName1 = "機關";
                thName2 = "學校";
                thName3 = "服務業";
            }
            else
            {//室內停車場智慧照明
                thName1 = "集合<br>住宅 ";
                thName2 = "辦公<br>大樓";
                thName3 = "服務業";
            }
            strHtml += "<tr>";
            if (i == 0)
            {
                strPeopleName = dt.Rows[0]["M_Name"].ToString().Trim();
                stePeoplePhone = dt.Rows[0]["M_Tel"].ToString().Trim();
                strManager = dt.Rows[0]["bossname"].ToString().Trim();
                strChkDate = Convert.ToDateTime(dt.Rows[0]["RC_CheckDate"].ToString()).ToString("yyyy/MM/dd");
                strCreateDate = Convert.ToDateTime(dt.Rows[0]["RM_ModDate"].ToString()).ToString("yyyy/MM/dd");
                strHtml += "<td rowspan='" + rowSpan + "' align='center' width = '5%'>本<br>月<br>執<br>行<br>進<br>度</td>";
            }
            if (dt.Rows[i]["RM_CPType"].ToString().Trim() != "04" && dt.Rows[i]["RM_CPType"].ToString().Trim() != "05")
            {
                if (dt.Rows[i]["RM_CPType"].ToString().Trim() == "01")
                {
                    strU = "(KW)";
                    strUSumAppl = "本月申請總冷氣能力(kW)";
                    strUSumFinish = "本月完成總冷氣能力(kW)";
                }
                if (dt.Rows[i]["RM_CPType"].ToString().Trim() == "02")
                {
                    strU = "(具)";
                    strUSumAppl = "本月申請更換照明瓦數(W)";
                    strUSumFinish = "本月完成更換照明瓦數(W)";
                }
                if (dt.Rows[i]["RM_CPType"].ToString().Trim() == "03")
                {
                    strU = "(盞)";
                    strUSumAppl = "本月申請更換照明瓦數(W)";
                    strUSumFinish = "本月完成更換照明瓦數(W)";
                }
                strHtml += "<td rowspan='8'>" + dt.Rows[i]["C_Item_cn"].ToString().Trim() + "</td>";
                strHtml += "<td colspan='6'>本期累計核定數：" + dt.Rows[i]["RM_Finish"].ToString().Trim() + strU + "</td>";
                strHtml += "<td colspan='6'>本期規劃數：" + dt.Rows[i]["RM_Planning"].ToString().Trim() + strU + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td colspan='3' align='center'>申請數量" + strU + "</td>";
                strHtml += "<td colspan='3' align='center'>核定數量" + strU + "</td>";
                //strHtml += "<td colspan='3' align='center'>申請總冷氣<br>能力(kW)</td>";
                //strHtml += "<td colspan='3' align='center'>完成總冷氣<br>能力(kW)</td>";
                strHtml += "<td colspan='3' align='center'>" + strUSumAppl + "</td>";
                strHtml += "<td colspan='3' align='center'>" + strUSumFinish + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='center'>" + thName1 + "</td>";
                strHtml += "<td align='center'>" + thName2 + "</td>";
                strHtml += "<td align='center'>" + thName3 + "</td>";
                strHtml += "<td align='center'>" + thName1 + "</td>";
                strHtml += "<td align='center'>" + thName2 + "</td>";
                strHtml += "<td align='center'>" + thName3 + "</td>";
                strHtml += "<td align='center'>" + thName1 + "</td>";
                strHtml += "<td align='center'>" + thName2 + "</td>";
                strHtml += "<td align='center'>" + thName3 + "</td>";
                strHtml += "<td align='center'>" + thName1 + "</td>";
                strHtml += "<td align='center'>" + thName2 + "</td>";
                strHtml += "<td align='center'>" + thName3 + "</td>";
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
                strHtml += "<td>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type1ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type2ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type3ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type4ValueSum"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td colspan='4' align='center'>申請數預期年節電量(度)</td>";
                strHtml += "<td colspan='4' align='center'>申請數預期年節電量(度)</td>";
                strHtml += "<td colspan='4' align='center'>申請數預期年節電量(度)</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td colspan='4' align='right' >" + dt.Rows[i]["RM_PreVal"].ToString().Trim() + "</td>";
                strHtml += "<td colspan='4' align='right' >" + dt.Rows[i]["RM_ChkVal"].ToString().Trim() + "</td>";
                strHtml += "<td colspan='4' align='right' >" + dt.Rows[i]["RM_NotChkVal"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                if (dt.Rows[i]["RM_CPType"].ToString().Trim() == "02")
                {
                    strHtml += "<td colspan='12'> 補充說明：(T8/T9)" + dt.Rows[i]["RM_Remark"].ToString().Trim() + "</td>";
                }
                else
                {
                    strHtml += "<td colspan='12'> 補充說明：" + dt.Rows[i]["RM_Remark"].ToString().Trim() + "</td>";
                }
                strHtml += "</tr>";
            }

        }
        strHtml += "</table>";
        strHtml += "<table border='1' cellpadding='0' cellspacing ='0' width='100%'>";
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

        //strPeopleName = dt.Rows[0]["M_Name"].ToString().Trim();
        //stePeoplePhone = dt.Rows[0]["M_Phone"].ToString().Trim();
        //strManager = dt.Rows[0]["bossname"].ToString().Trim();
        //strChkDate = Convert.ToDateTime(dt.Rows[0]["RC_CheckDate"].ToString()).ToString("yyyy/MM/dd");
        //strCreateDate = Convert.ToDateTime(dt.Rows[0]["RC_CreateDate"].ToString()).ToString("yyyy/MM/dd");


        return strHtml;
    }

    //只有04~05 沒有01~03
    private string html45(string strYear, string strHtml, DataTable dt)
    {
        string strPeopleName = "", stePeoplePhone = "", strManager = "", strChkDate = "", strCreateDate = "";
        string thName1 = "", thName2 = "", thName3 = "";
        string strS = "", strE = "", strSY = "", strSM = "", strSD = "", strEY = "", strEM = "", strED = "";

        int CPTypeCount = 0;
        int rowSpan = 0;
        DataTable dtGroup = dt.DefaultView.ToTable(true, "RM_CPType");
        CPTypeCount = dtGroup.Rows.Count;
        rowSpan = CPTypeCount * 8;
        
        strHtml += "<table border='1' cellpadding='0' cellspacing ='0' width='100%' >";
        strHtml += "<tr>";
        strHtml += "<td colspan='3' align='center'>執行機關</td><td colspan='3' align='center'>" + dt.Rows[0]["cityname"].ToString().Trim() + "</td>";
        strHtml += "<td colspan='3' align='center'>承辦局處</td><td colspan='3' align='center'>" + dt.Rows[0]["M_Office"].ToString().Trim() + "</td>";
        strHtml += "</tr>";
        strHtml += "<tr>";
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
        strHtml += "<td colspan='3' align='center' valign='center'>期數</td><td colspan ='9'>第" + dt.Rows[0]["RM_Stage"].ToString().Trim() + "期 自" + strSY + "年" + strSM + "月" + strSD + "日起至" + strEY + "年" + strEM + "月" + strED + "日止，共" + monthDiff(strS, strE) + "個月</td>";
        strHtml += "</tr>";
        strHtml += "<tr>";
        strHtml += "<td colspan='3' align='center'>提報月份</td><td colspan='9'>" + (Convert.ToInt32(strYear) - 1911).ToString() + "年" + dt.Rows[0]["RM_Month"].ToString().Trim() + "月</td>";
        strHtml += "</tr>";
        
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            thName1 = "機關";
            thName2 = "學校";
            thName3 = "服務業";
            strHtml += "<tr>";
            if (i == 0)
            {
                strPeopleName = dt.Rows[0]["M_Name"].ToString().Trim();
                stePeoplePhone = dt.Rows[0]["M_Phone"].ToString().Trim();
                strManager = dt.Rows[0]["bossname"].ToString().Trim();
                if (dt.Rows[0]["RC_CheckDate"] == null || dt.Rows[0]["RC_CheckDate"].ToString().Trim() == "")
                {
                    strChkDate = "";
                }
                else {
                    strChkDate = Convert.ToDateTime(dt.Rows[0]["RC_CheckDate"].ToString()).ToString("yyyy/MM/dd");
                }
                if (dt.Rows[0]["RC_CheckDate"] == null || dt.Rows[0]["RC_CheckDate"].ToString().Trim() == "")
                {
                    strCreateDate = "";
                }
                else {
                    strCreateDate = Convert.ToDateTime(dt.Rows[0]["RM_ModDate"].ToString()).ToString("yyyy/MM/dd");
                }
                
                
                strHtml += "<td rowspan='" + rowSpan + "' align='center' width = '5%'>本<br>月<br>執<br>行<br>進<br>度</td>";
            }
            if (dt.Rows[i]["RM_CPType"].ToString().Trim() == "04" || dt.Rows[i]["RM_CPType"].ToString().Trim() == "05")
            {
                strHtml += "<td rowspan='8' colspan = '2'>" + dt.Rows[i]["C_Item_cn"].ToString().Trim() + "</td>";
                strHtml += "<td colspan='4'>本期累計完成數：" + dt.Rows[i]["RM_Finish"].ToString().Trim() + "(套)</td>";
                strHtml += "<td colspan='5'>本期規劃數：" + dt.Rows[i]["RM_Planning"].ToString().Trim() + "(套)</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td colspan='3' align='center' >申請數量(套)</td>";
                strHtml += "<td colspan='3' align='center' >核定數量(套)</td>";
                strHtml += "<td colspan='3' align='center' >完成數量(套)</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td algn='center'>" + thName1 + "</td>";
                strHtml += "<td algn='center'>" + thName2 + "</td>";
                strHtml += "<td algn='center'>" + thName3 + "</td>";
                strHtml += "<td algn='center'>" + thName1 + "</td>";
                strHtml += "<td algn='center'>" + thName2 + "</td>";
                strHtml += "<td algn='center'>" + thName3 + "</td>";
                strHtml += "<td algn='center'>" + thName1 + "</td>";
                strHtml += "<td algn='center'>" + thName2 + "</td>";
                strHtml += "<td algn='center'>" + thName3 + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='right' >" + dt.Rows[i]["RM_Type1Value1"].ToString().Trim() + "</td>";
                strHtml += "<td align='right' >" + dt.Rows[i]["RM_Type1Value2"].ToString().Trim() + "</td>";
                strHtml += "<td align='right' >" + dt.Rows[i]["RM_Type1Value3"].ToString().Trim() + "</td>";
                strHtml += "<td align='right' >" + dt.Rows[i]["RM_Type2Value1"].ToString().Trim() + "</td>";
                strHtml += "<td align='right' >" + dt.Rows[i]["RM_Type2Value2"].ToString().Trim() + "</td>";
                strHtml += "<td align='right' >" + dt.Rows[i]["RM_Type2Value3"].ToString().Trim() + "</td>";
                strHtml += "<td align='right' >" + dt.Rows[i]["RM_Type3Value1"].ToString().Trim() + "</td>";
                strHtml += "<td align='right' >" + dt.Rows[i]["RM_Type3Value2"].ToString().Trim() + "</td>";
                strHtml += "<td align='right' >" + dt.Rows[i]["RM_Type3Value3"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='center' >合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type1ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td align='center'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type2ValueSum"].ToString().Trim() + "</td>";
                strHtml += "<td align='center'>合計</td>";
                strHtml += "<td colspan='2' align='right'>" + dt.Rows[i]["RM_Type3ValueSum"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td align='center' colspan='3'>申請數預期年節電量(度)</td>";
                strHtml += "<td align='center' colspan='3'>申請數預期年節電量(度)</td>";
                strHtml += "<td align='center' colspan='3'>申請數預期年節電量(度)</td>";
                strHtml += "</tr >";
                strHtml += "<tr>";
                strHtml += "<td align='right' colspan='3'>" + dt.Rows[i]["RM_PreVal"].ToString().Trim() + "</td>";
                strHtml += "<td align='right' colspan='3'>" + dt.Rows[i]["RM_ChkVal"].ToString().Trim() + "</td>";
                strHtml += "<td align='right' colspan='3'>" + dt.Rows[i]["RM_NotChkVal"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td colspan='9'> 補充說明：" + dt.Rows[i]["RM_Remark"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
            }
        }
        strHtml += "</table>";
        strHtml += "<table border = '1' cellpadding = '0' cellspacing = '0' width = '100%'>";
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
        //strPeopleName = dt.Rows[0]["M_Name"].ToString().Trim();
        //stePeoplePhone = dt.Rows[0]["M_Phone"].ToString().Trim();
        //strManager = dt.Rows[0]["bossname"].ToString().Trim();
        //strChkDate = Convert.ToDateTime(dt.Rows[0]["RC_CheckDate"].ToString()).ToString("yyyy/MM/dd");
        //strCreateDate = Convert.ToDateTime(dt.Rows[0]["RC_CreateDate"].ToString()).ToString("yyyy/MM/dd");
        

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
        string strYear = "", fileYear="";
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