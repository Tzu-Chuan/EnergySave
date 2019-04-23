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

public partial class handler_ExportInfo : System.Web.UI.Page
{
    Member_DB m_db = new Member_DB();
    CheckPoint_DB cp_db = new CheckPoint_DB();
    CodeTable_DB ct_db = new CodeTable_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        string period1 = string.Empty;
        string period2 = string.Empty;
        string period3 = string.Empty;

        //新檔名
        string newName = Guid.NewGuid().ToString("N").Substring(0, 10) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
        string FileName = string.Empty;
        Aspose.Words.Document doc = new Aspose.Words.Document(Server.MapPath("~/Template/ProInfo.docx"));
        DocumentBuilder builder = new DocumentBuilder(doc);
        if (Request.QueryString["v"] != null && Request.QueryString["tp"] != null)
        {
            //副檔名
            string ext = (Request.QueryString["tp"].ToString() == "W") ? ".docx" : ".pdf";
            string id  = Request.QueryString["v"].ToString();
            m_db._M_ID = id;
            string projectid = m_db.getProgectGuidByPersonId();

            DataSet mds = m_db.getPersonInfoById();
            DataTable mmdt = mds.Tables[0];
            DataTable addt = mds.Tables[1];
            DataTable exDt = cp_db.getPushitemExFinish(projectid);
            if (projectid != "")
            {
                DataSet ds = cp_db.getExportInfo(projectid);
                DataTable info_dt = ds.Tables[0];
                DataTable cp_dt = ds.Tables[1];
                if (info_dt.Rows.Count > 0)
                {
                    doc.Range.Replace("@PD_StartDate", ROC_Date(info_dt.Rows[0]["PD_StartDate"].ToString()), false, false);
                    doc.Range.Replace("@PD_EndDate", ROC_Date(info_dt.Rows[0]["PD_EndDate"].ToString()), false, false);
                    doc.Range.Replace("@mDiff_t", monthDiff(info_dt.Rows[0]["PD_StartDate"].ToString(), info_dt.Rows[0]["PD_EndDate"].ToString()), false, false);
                    doc.Range.Replace("@I_1_Sdate", ROC_Date(info_dt.Rows[0]["I_1_Sdate"].ToString()), false, false);
                    doc.Range.Replace("@I_1_Edate", ROC_Date(info_dt.Rows[0]["I_1_Edate"].ToString()), false, false);
                    doc.Range.Replace("@mDiff_1", monthDiff(info_dt.Rows[0]["I_1_Sdate"].ToString(), info_dt.Rows[0]["I_1_Edate"].ToString()), false, false);
                    doc.Range.Replace("@I_2_Sdate", ROC_Date(info_dt.Rows[0]["I_2_Sdate"].ToString()), false, false);
                    doc.Range.Replace("@I_2_Edate", ROC_Date(info_dt.Rows[0]["I_2_Edate"].ToString()), false, false);
                    doc.Range.Replace("@mDiff_2", monthDiff(info_dt.Rows[0]["I_2_Sdate"].ToString(), info_dt.Rows[0]["I_2_Edate"].ToString()), false, false);
                    doc.Range.Replace("@I_3_Sdate", ROC_Date(info_dt.Rows[0]["I_3_Sdate"].ToString()), false, false);
                    doc.Range.Replace("@I_3_Edate", ROC_Date(info_dt.Rows[0]["I_3_Edate"].ToString()), false, false);
                    doc.Range.Replace("@mDiff_3", monthDiff(info_dt.Rows[0]["I_3_Sdate"].ToString(), info_dt.Rows[0]["I_3_Edate"].ToString()), false, false);
                    doc.Range.Replace("@I_Money_item1_1", info_dt.Rows[0]["I_Money_item1_1"].ToString(), false, false);
                    doc.Range.Replace("@I_Money_item1_2", info_dt.Rows[0]["I_Money_item1_2"].ToString(), false, false);
                    doc.Range.Replace("@I_Money_item1_3", info_dt.Rows[0]["I_Money_item1_3"].ToString(), false, false);
                    doc.Range.Replace("@I_Money_item1_all", info_dt.Rows[0]["I_Money_item1_all"].ToString(), false, false);
                    doc.Range.Replace("@I_Money_item2_1", info_dt.Rows[0]["I_Money_item2_1"].ToString(), false, false);
                    doc.Range.Replace("@I_Money_item2_2", info_dt.Rows[0]["I_Money_item2_2"].ToString(), false, false);
                    doc.Range.Replace("@I_Money_item2_3", info_dt.Rows[0]["I_Money_item2_3"].ToString(), false, false);
                    doc.Range.Replace("@I_Money_item2_all", info_dt.Rows[0]["I_Money_item2_all"].ToString(), false, false);
                    doc.Range.Replace("@I_Money_item3_1", info_dt.Rows[0]["I_Money_item3_1"].ToString(), false, false);
                    doc.Range.Replace("@I_Money_item3_2", info_dt.Rows[0]["I_Money_item3_2"].ToString(), false, false);
                    doc.Range.Replace("@I_Money_item3_3", info_dt.Rows[0]["I_Money_item3_3"].ToString(), false, false);
                    doc.Range.Replace("@I_Money_item3_all", info_dt.Rows[0]["I_Money_item3_all"].ToString(), false, false);
                    doc.Range.Replace("@I_Money_item4_1", info_dt.Rows[0]["I_Money_item4_1"].ToString(), false, false);
                    doc.Range.Replace("@I_Money_item4_2", info_dt.Rows[0]["I_Money_item4_2"].ToString(), false, false);
                    doc.Range.Replace("@I_Money_item4_3", info_dt.Rows[0]["I_Money_item4_3"].ToString(), false, false);
                    doc.Range.Replace("@I_Money_item4_all", info_dt.Rows[0]["I_Money_item4_all"].ToString(), false, false);
                    double tmpall = double.Parse(info_dt.Rows[0]["I_Money_item1_1"].ToString()) + double.Parse(info_dt.Rows[0]["I_Money_item2_1"].ToString()) + double.Parse(info_dt.Rows[0]["I_Money_item3_1"].ToString());
                    doc.Range.Replace("@I_Money_t1", tmpall.ToString("#0.000"), false, false);
                    tmpall = double.Parse(info_dt.Rows[0]["I_Money_item1_2"].ToString()) + double.Parse(info_dt.Rows[0]["I_Money_item2_2"].ToString()) + double.Parse(info_dt.Rows[0]["I_Money_item3_2"].ToString());
                    doc.Range.Replace("@I_Money_t2", tmpall.ToString("#0.000"), false, false);
                    tmpall = double.Parse(info_dt.Rows[0]["I_Money_item1_3"].ToString()) + double.Parse(info_dt.Rows[0]["I_Money_item2_3"].ToString()) + double.Parse(info_dt.Rows[0]["I_Money_item3_3"].ToString());
                    doc.Range.Replace("@I_Money_t3", tmpall.ToString("#0.000"), false, false);
                    tmpall = double.Parse(info_dt.Rows[0]["I_Money_item1_all"].ToString()) + double.Parse(info_dt.Rows[0]["I_Money_item2_all"].ToString()) + double.Parse(info_dt.Rows[0]["I_Money_item3_all"].ToString());
                    doc.Range.Replace("@I_Money_ta", tmpall.ToString("#0.000"), false, false);
                    if (info_dt.Rows[0]["I_Other_Oneself"].ToString() == "Y")
                    {
                        string OtherStr = "自籌款，金額：" + info_dt.Rows[0]["I_Other_Oneself_Money"].ToString() + " 仟元";
                        doc.Range.Replace("@Other_Money1", OtherStr, false, false);
                    }
                    if (info_dt.Rows[0]["I_Other_Oneself"].ToString() == "Y")
                    {
                        string OtherStr2 = "其他機關補助，機關名稱：" + info_dt.Rows[0]["I_Other_Other_name"].ToString() + "，金額：" + info_dt.Rows[0]["I_Other_Other_Money"].ToString() + " 仟元";
                        doc.Range.Replace("@Other_Money2", OtherStr2, false, false);
                    }
                    //換行(Enter) = ControlChar.LineBreak
                    doc.Range.Replace("@I_Target", info_dt.Rows[0]["I_Target"].ToString().Replace("\n", ControlChar.LineBreak), false, false);
                    doc.Range.Replace("@I_Summary", info_dt.Rows[0]["I_Summary"].ToString().Replace("\n", ControlChar.LineBreak), false, false);
                    //設備汰換與智慧用電預計完成數
                    doc.Range.Replace("@I_Finish_item1_1", info_dt.Rows[0]["I_Finish_item1_1"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item1_2", info_dt.Rows[0]["I_Finish_item1_2"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item1_3", info_dt.Rows[0]["I_Finish_item1_3"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item1_all", info_dt.Rows[0]["I_Finish_item1_all"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item2_1", info_dt.Rows[0]["I_Finish_item2_1"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item2_2", info_dt.Rows[0]["I_Finish_item2_2"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item2_3", info_dt.Rows[0]["I_Finish_item2_3"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item2_all", info_dt.Rows[0]["I_Finish_item2_all"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item3_1", info_dt.Rows[0]["I_Finish_item3_1"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item3_2", info_dt.Rows[0]["I_Finish_item3_2"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item3_3", info_dt.Rows[0]["I_Finish_item3_3"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item3_all", info_dt.Rows[0]["I_Finish_item3_all"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item4_1", info_dt.Rows[0]["I_Finish_item4_1"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item4_2", info_dt.Rows[0]["I_Finish_item4_2"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item4_3", info_dt.Rows[0]["I_Finish_item4_3"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item4_all", info_dt.Rows[0]["I_Finish_item4_all"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item5_1", info_dt.Rows[0]["I_Finish_item5_1"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item5_2", info_dt.Rows[0]["I_Finish_item5_2"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item5_3", info_dt.Rows[0]["I_Finish_item5_3"].ToString(), false, false);
                    doc.Range.Replace("@I_Finish_item5_all", info_dt.Rows[0]["I_Finish_item5_all"].ToString(), false, false);


                    builder.MoveToBookmark("ExFinish", false, true);
                    builder.Font.Bold = true;
                    builder.Font.Size = 16;
                    builder.Write("\n擴大補助預計完成數");
                    builder.Font.Bold = false;
                    builder.Font.Size = 14;

                    DataView exDv = exDt.DefaultView;
                    exDv.Sort = "P_Period";
                    string tmpstr = string.Empty;
                    if (exDv.Count > 0)
                    {
                        string tmpStage = exDv[0]["P_Period"].ToString();
                        tmpstr = "<table border='1' cellpadding='0' cellspacing='0' width='100%'>";
                        tmpstr += "<tr><th align='center'>項目</th><th align='center'>第1期</th></tr>";
                        for (int i = 0; i < exDv.Count; i++)
                        {
                            if(tmpStage != exDv[i]["P_Period"].ToString())
                            {
                                tmpstr += "<tr>";
                                tmpstr += "<th align='center'>項目</th>";
                                tmpstr += "<th align='center'>第" + exDv[i]["P_Period"].ToString() + "期</th>";
                                tmpstr += "</tr>";
                            }

                            tmpstr += "<tr>";
                            tmpstr += "<td align='center'>"+ getExType_Cn(exDv[i]["P_ItemName"].ToString()) + "</td>";
                            tmpstr += "<td align='center'>"+ exDv[i]["P_ExFinish"].ToString() + "</td>";
                            tmpstr += "</tr>";

                            tmpStage = exDv[i]["P_Period"].ToString();
                        }
                        tmpstr += "</table>";
                    }
                    builder.InsertHtml(tmpstr);


                    //承辦人
                    if (mmdt.Rows.Count > 0)
                    {
                        //設定檔名
                        FileName = mmdt.Rows[0]["City"].ToString() + "(" + ROC_Date(info_dt.Rows[0]["PD_StartDate"].ToString()) + "~" + ROC_Date(info_dt.Rows[0]["PD_EndDate"].ToString()) + ")";
                        //人員資料
                        doc.Range.Replace("@City", mmdt.Rows[0]["City"].ToString(), false, false);
                        doc.Range.Replace("@Office", mmdt.Rows[0]["M_Office"].ToString(), false, false);

                        doc.Range.Replace("@cName", mmdt.Rows[0]["M_Name"].ToString(), false, false);
                        doc.Range.Replace("@cTitle", mmdt.Rows[0]["M_JobTitle"].ToString(), false, false);
                        doc.Range.Replace("@cEmail", mmdt.Rows[0]["M_Email"].ToString(), false, false);
                        if (mmdt.Rows[0]["M_Ext"].ToString().Trim() != "")
                            doc.Range.Replace("@cTel", mmdt.Rows[0]["M_Tel"].ToString() + "-" + mmdt.Rows[0]["M_Ext"].ToString(), false, false);
                        else
                            doc.Range.Replace("@cTel", mmdt.Rows[0]["M_Tel"].ToString(), false, false);
                        doc.Range.Replace("@cFax", mmdt.Rows[0]["M_Fax"].ToString(), false, false);
                        doc.Range.Replace("@cPhone", mmdt.Rows[0]["M_Phone"].ToString(), false, false);
                        doc.Range.Replace("@cAddr", mmdt.Rows[0]["M_Addr"].ToString(), false, false);
                    }
                    //承辦主管
                    if (addt.Rows.Count > 0)
                    {
                        doc.Range.Replace("@mName", addt.Rows[0]["M_Name"].ToString(), false, false);
                        doc.Range.Replace("@mTitle", addt.Rows[0]["M_JobTitle"].ToString(), false, false);
                        doc.Range.Replace("@mEmail", addt.Rows[0]["M_Email"].ToString(), false, false);
                        if (addt.Rows[0]["M_Ext"].ToString().Trim() != "")
                            doc.Range.Replace("@mTel", addt.Rows[0]["M_Tel"].ToString() + "-" + addt.Rows[0]["M_Ext"].ToString(), false, false);
                        else
                            doc.Range.Replace("@mTel", addt.Rows[0]["M_Tel"].ToString(), false, false);
                        doc.Range.Replace("@mFax", addt.Rows[0]["M_Fax"].ToString(), false, false);
                        doc.Range.Replace("@mPhone", addt.Rows[0]["M_Phone"].ToString(), false, false);
                        doc.Range.Replace("@mAddr", addt.Rows[0]["M_Addr"].ToString(), false, false);
                    }

                    period1 = ROC_Date(info_dt.Rows[0]["I_1_Sdate"].ToString()) + "至" + ROC_Date(info_dt.Rows[0]["I_1_Edate"].ToString());
                    period2 = ROC_Date(info_dt.Rows[0]["I_2_Sdate"].ToString()) + "至" + ROC_Date(info_dt.Rows[0]["I_2_Edate"].ToString());
                    period3 = ROC_Date(info_dt.Rows[0]["I_3_Sdate"].ToString()) + "至" + ROC_Date(info_dt.Rows[0]["I_3_Edate"].ToString());
                }




                //查核點
                //節電基礎工作
                string bw1 = checkpoint(cp_dt, "P_Period='1' and P_Type='01'");
                string bw2 = checkpoint(cp_dt, "P_Period='2' and P_Type='01'");
                string bw3 = checkpoint(cp_dt, "P_Period='3' and P_Type='01'");
                //因地制宜
                string p1 = checkpoint(cp_dt, "P_Period='1' and P_Type='02'");
                string p2 = checkpoint(cp_dt, "P_Period='2' and P_Type='02'");
                string p3 = checkpoint(cp_dt, "P_Period='3' and P_Type='02'");
                //設備汰換與智慧用電
                string s1 = checkpoint(cp_dt, "P_Period='1' and P_Type='03'");
                string s2 = checkpoint(cp_dt, "P_Period='2' and P_Type='03'");
                string s3 = checkpoint(cp_dt, "P_Period='3' and P_Type='03'");
                //擴大補助
                string ex1 = checkpoint(cp_dt, "P_Period='1' and P_Type='04'");
                string ex2 = checkpoint(cp_dt, "P_Period='2' and P_Type='04'");
                string ex3 = checkpoint(cp_dt, "P_Period='3' and P_Type='04'");

                builder.MoveToBookmark("checkpoint", false, true);

                //第一期
                builder.Font.Bold = true;
                builder.Write("第一期");
                builder.Font.Bold = false;
                builder.Write("\n1.節電基礎工作");
                builder.InsertHtml(bw1);
                builder.Write("\n2.因地制宜");
                builder.InsertHtml(p1);
                builder.Write("\n3.設備汰換與智慧用電");
                builder.InsertHtml(s1);
                builder.Write("\n4.擴大補助");
                builder.InsertHtml(ex1);
                //第二期
                builder.Font.Bold = true;
                builder.Write("\n第二期");
                builder.Font.Bold = false;
                builder.Write("\n1.節電基礎工作");
                builder.InsertHtml(bw2);
                builder.Write("\n2.因地制宜");
                builder.InsertHtml(p2);
                builder.Write("\n3.設備汰換與智慧用電");
                builder.InsertHtml(s2);
                builder.Write("\n4.擴大補助");
                builder.InsertHtml(ex2);
                //第三期
                builder.Font.Bold = true;
                builder.Write("\n第三期");
                builder.Font.Bold = false;
                builder.Write("\n1.節電基礎工作");
                builder.InsertHtml(bw3);
                builder.Write("\n2.因地制宜");
                builder.InsertHtml(p3);
                builder.Write("\n3.設備汰換與智慧用電");
                builder.InsertHtml(s3);
                builder.Write("\n4.擴大補助");
                builder.InsertHtml(ex3);

                #region 預定工作進度

                #region 節電基礎工作
                builder.MoveToBookmark("progress", false, true);
                builder.Font.Size = 16;
                builder.Font.Bold = true;
                builder.Write("(一)節電基礎工作\n");
                builder.Font.Size = 14;
                builder.Font.Bold = false;
                string pbw1 = Progress(projectid, "P_Period='1' and P_Type='01'", "第一期(" + period1 + ")");
                if (pbw1 == "")
                    builder.Write("無第一期資料");
                else
                    builder.InsertHtml(pbw1);
                builder.Write("\n");
                string pbw2 = Progress(projectid, "P_Period='2' and P_Type='01'", "第二期(" + period2 + ")");
                if (pbw2 == "")
                    builder.Write("無第二期資料");
                else
                    builder.InsertHtml(pbw2);
                builder.Write("\n");
                string pbw3 = Progress(projectid, "P_Period='3' and P_Type='01'", "第三期(" + period3 + ")");
                if (pbw3 == "")
                    builder.Write("無第三期資料");
                else
                    builder.InsertHtml(pbw3);
                builder.Write("\n");
                #endregion

                #region 因地制宜
                //因地制宜
                builder.Font.Size = 16;
                builder.Font.Bold = true;
                builder.Write("(二)因地制宜\n");
                builder.Font.Size = 14;
                builder.Font.Bold = false;
                string pp1 = Progress(projectid, "P_Period='1' and P_Type='02'", "第一期(" + period1 + ")");
                if (pp1 == "")
                    builder.Write("無第一期資料");
                else
                    builder.InsertHtml(pp1);
                builder.Write("\n");
                string pp2 = Progress(projectid, "P_Period='2' and P_Type='02'", "第二期(" + period2 + ")");
                if (pp2 == "")
                    builder.Write("無第二期資料");
                else
                    builder.InsertHtml(pp2);
                builder.Write("\n");
                string pp3 = Progress(projectid, "P_Period='3' and P_Type='02'", "第三期(" + period3 + ")");
                if (pp3 == "")
                    builder.Write("無第三期資料");
                else
                    builder.InsertHtml(pp3);
                builder.Write("\n");
                #endregion

                #region 設備汰換及智慧用電
                builder.Font.Size = 16;
                builder.Font.Bold = true;
                builder.Write("(三)設備汰換及智慧用電\n");
                builder.Font.Size = 14;
                builder.Font.Bold = false;
                string ps1 = Progress(projectid, "P_Period='1' and P_Type='03'", "第一期(" + period1 + ")");
                if (ps1 == "")
                    builder.Write("無第一期資料");
                else
                    builder.InsertHtml(ps1);
                builder.Write("\n");
                string ps2 = Progress(projectid, "P_Period='2' and P_Type='03'", "第二期(" + period2 + ")");
                if (ps2 == "")
                    builder.Write("無第二期資料");
                else
                    builder.InsertHtml(ps2);
                builder.Write("\n");
                string ps3 = Progress(projectid, "P_Period='3' and P_Type='03'", "第三期(" + period3 + ")");
                if (ps3 == "")
                    builder.Write("無第三期資料");
                else
                    builder.InsertHtml(ps3);
                builder.Write("\n");
                #endregion

                #region 擴大補助
                builder.Font.Size = 16;
                builder.Font.Bold = true;
                builder.Write("(四)擴大補助\n");
                builder.Font.Size = 14;
                builder.Font.Bold = false;
                string pex1 = Progress(projectid, "P_Period='1' and P_Type='04'", "第一期(" + period1 + ")");
                if (pex1 == "")
                    builder.Write("無第一期資料");
                else
                    builder.InsertHtml(pex1);
                builder.Write("\n");
                string pex2 = Progress(projectid, "P_Period='2' and P_Type='04'", "第二期(" + period2 + ")");
                if (pex2 == "")
                    builder.Write("無第二期資料");
                else
                    builder.InsertHtml(pex2);
                builder.Write("\n");
                string pex3 = Progress(projectid, "P_Period='3' and P_Type='04'", "第三期(" + period3 + ")");
                if (pex3 == "")
                    builder.Write("無第三期資料");
                else
                    builder.InsertHtml(pex3);
                builder.Write("\n");
                #endregion

                #endregion

                //變換整份word字型
                foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
                {
                    run.Font.Name = "標楷體";
                }

                #region 頁碼
                // Go to the primary footer
                builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                // Add fields for current page number
                builder.InsertField("PAGE", "");
                // Add any custom text
                builder.Write(" / ");
                // Add field for total page numbers in document
                builder.InsertField("NUMPAGES", "");
                #endregion

                if (Request.QueryString["tp"].ToString() == "W")
                    doc.Save(Server.MapPath("~/Template/" + newName + ".docx"));
                else
                    doc.Save(Server.MapPath("~/Template/" + newName + ".pdf"), SaveFormat.Pdf);
            }

            
            Response.Clear();
            Response.ClearHeaders();
            Response.AddHeader("content-disposition", "attachment;filename=" + Server.UrlEncode(FileName + ext));
            Response.ContentType = "application/octet-stream";
            Response.WriteFile(Server.MapPath("~/Template/" + newName + ext));
            Response.Flush();
            File.Delete(Server.MapPath("~/Template/" + newName + ext));
            Response.End();
        }
    }
    

    private string monthDiff(string startM,string endM)
    {
        string tmpstr = string.Empty;
        DateTime dt1 = Convert.ToDateTime(startM);
        DateTime dt2 = Convert.ToDateTime(endM);

        int Year = dt2.Year - dt1.Year;
        int Month = (dt2.Year - dt1.Year) * 12 + (dt2.Month - dt1.Month) + 1;

        return Month.ToString();
    }

    private string checkpoint(DataTable dt, string scriptStr)
    {
        string tmpstr = "<table border='1' cellpadding='0' cellspacing='0' width='100%'>";
        tmpstr += "<tr><td align='center'>推動工作</td><td align='center'>查核點</td><td align='center'>預定時間</td><td align='center'>查核點概述</td></tr>";
        if (dt.Rows.Count > 0)
        {
            string tmpid = string.Empty;
            DataView dv = dt.DefaultView;
            dv.RowFilter = scriptStr;
            if (dv.Count > 0)
            {
                for (int i = 0; i < dv.Count; i++)
                {
                    string itemName = string.Empty;
                    switch (dv[i]["P_Type"].ToString())
                    {
                        default:
                            itemName = dv[i]["P_ItemName"].ToString();
                            break;
                        case "03":
                            itemName = getPItemNa(dv[i]["P_ItemName"].ToString());
                            if (dv[i]["P_ItemName"].ToString() == "99")
                                itemName += "-" + dv[i]["CP_Desc"].ToString();
                            break;
                        case "04":
                            itemName = getExType_Cn(dv[i]["P_ItemName"].ToString());
                            if (dv[i]["P_ItemName"].ToString() == "99")
                                itemName += "-" + dv[i]["CP_Desc"].ToString();
                            break;
                    }
                    tmpstr += "<tr>";
                    if (tmpid != dv[i]["P_Guid"].ToString())
                        tmpstr += "<td rowspan='" + dv[i]["ItemCount"].ToString() + "' align='center'>" + itemName + "</td>";
                    tmpstr += "<td align='center'>" + dv[i]["CP_Point"].ToString() + "</td>";
                    tmpstr += "<td align='center' width='20%'>" + dv[i]["CP_ReserveYear"].ToString() + " 年 " + dv[i]["CP_ReserveMonth"].ToString() + " 月</td>";
                    tmpstr += "<td>" + dv[i]["CP_Desc"].ToString() + "</td>";
                    tmpstr += "</tr>";
                    tmpid = dv[i]["P_Guid"].ToString();
                }
            }
            else
                tmpstr += "<tr><td colspan='4'>查無資料</td></tr>";

        }
        tmpstr += "</table>";
        return tmpstr;
    }

    private string Progress(string project_id,string scriptStr,string PeriodStr)
    {
        string tmpstr = string.Empty;
        DataSet ds = cp_db.getExportProgress(project_id, scriptStr);
        DataTable dym = ds.Tables[0];
        DataTable dt = ds.Tables[1];
        DataTable total_dt = ds.Tables[2];
        if (dym.Rows.Count > 0)
        {
            tmpstr = "<table border='1' cellpadding='0' cellspacing='0' width='100%'>";
            tmpstr += "<tr><td align='center' rowspan='3'>推動項目</td><td align='center' rowspan='3'>工作比重(%)</td><td align='center' rowspan='3'>年  月<br>進度(%)</td>";
            tmpstr += "<td align='center' colspan='" + dym.Rows.Count + "'>" + PeriodStr + "</td></tr>";
            //年
            tmpstr += "<tr>";
            string ytmp = string.Empty;
            for (int i = 0; i < dym.Rows.Count; i++)
            {
                if (ytmp != dym.Rows[i]["Y"].ToString())
                    tmpstr += "<td align='center' colspan='" + dym.Rows[i]["yc"].ToString() + "'>" + dym.Rows[i]["Y"].ToString() + "年</td>";
                ytmp = dym.Rows[i]["Y"].ToString();
            }
            tmpstr += "</tr>";
            //月
            tmpstr += "<tr>";
            for (int i = 0; i < dym.Rows.Count; i++)
            {
                tmpstr += "<td align='center'>" + dym.Rows[i]["M"].ToString() + "月</td>";
            }
            tmpstr += "</tr>";
            #region 累計預定進度合計(舊)
            //if (total_dt.Rows.Count > 0)
            //{
            //    for (int i = 0; i < dym.Rows.Count; i++)
            //    {
            //        for (int j = 0; j < total_dt.Rows.Count; j++)
            //        {
            //            if (total_dt.Rows[j]["Y"].ToString()== dym.Rows[i]["Y"].ToString() && total_dt.Rows[j]["M"].ToString() == dym.Rows[i]["M"].ToString())
            //            {
            //                PStrTotal += "<td align='center'>" + double.Parse(total_dt.Rows[j]["total"].ToString()) + "%</td>";
            //                break;
            //            }
            //        }
            //    }
            //}
            #endregion

            #region 累計預定進度合計
            string PStrTotal = string.Empty;
            DataTable pdt = cp_db.getProgressTotal(project_id);
            DataView pdv = pdt.DefaultView;
            pdv.RowFilter = scriptStr;
            List<string> idAry = new List<string>();
            List<string> valAry = new List<string>();
            if (pdv.Count > 0)
            {
                for (int i = 0; i < dym.Rows.Count; i++)
                {
                    for (int j = 0; j < pdv.Count; j++)
                    {
                        if (dym.Rows[i]["Y"].ToString() == pdv[j]["CP_ReserveYear"].ToString() && dym.Rows[i]["M"].ToString() == pdv[j]["CP_ReserveMonth"].ToString())
                        {
                            if (!idAry.Contains(pdv[j]["P_ID"].ToString()))
                            {
                                idAry.Add(pdv[j]["P_ID"].ToString());
                                valAry.Add(pdv[j]["CP_Process"].ToString());
                            }
                            else
                            {
                                valAry[idAry.IndexOf(pdv[j]["P_ID"].ToString())] = pdv[j]["CP_Process"].ToString();
                            }
                        }
                    }

                    double temp = 0;
                    for (int a = 0; a < valAry.Count; a++)
                    {
                        temp += (valAry[a] != "") ? double.Parse(valAry[a]) : 0;
                    }
                    PStrTotal += "<td align='center'>" + temp + "%</td>";
                }
            }
            #endregion


            DataView dv = dt.DefaultView;
            dv.RowFilter = scriptStr;
            if (dv.Count > 0)
            {
                string PStr1 = string.Empty;
                string PStr2 = string.Empty;
                string PStr3 = string.Empty;
                int PushItemCount = 1;
                int rnum = 0; //計算同推動項目群組目前是第幾筆
                double WorkRatioNum = 0;
                DataTable totalTab = new DataTable();
                for (int i = 0; i < dv.Count; i++)
                {
                    if (PushItemCount == 1)
                    {
                        string itemName = string.Empty;
                        switch (dv[i]["P_Type"].ToString())
                        {
                            default:
                                itemName= dv[i]["P_ItemName"].ToString();
                                break;
                            case "03":
                                itemName = getPItemNa(dv[i]["P_ItemName"].ToString());
                                if (dv[i]["P_ItemName"].ToString() == "99")
                                    itemName += "-" + dv[i]["CP_Desc"].ToString();
                                break;
                            case "04":
                                itemName = getExType_Cn(dv[i]["P_ItemName"].ToString());
                                if (dv[i]["P_ItemName"].ToString() == "99")
                                    itemName += "-" + dv[i]["CP_Desc"].ToString();
                                break;
                        }
                        PStr1 += "<tr><td rowspan='3' align='center'>" + itemName + "</td>";
                        PStr1 += "<td rowspan='3' align='center'>" + dv[i]["P_WorkRatio"].ToString().Trim() + "%</td>";
                        PStr1 += "<td align='center'>查核點</td>";
                        PStr2 += "<tr><td align='center'>累計預定進度(%)</td>";

                        for (int j = 0; j < dym.Rows.Count; j++)
                        {
                            if (dym.Rows[j]["Y"].ToString().Trim() == dv[i]["CP_ReserveYear"].ToString().Trim() && dym.Rows[j]["M"].ToString().Trim() == dv[i]["CP_ReserveMonth"].ToString().Trim())
                            {
                                PStr1 += "<td align='center'>" + dv[i]["CP_Point"].ToString().Trim() + "</td>";
                                PStr2 += "<td align='center'>" + dv[i]["CP_Process"].ToString().Trim() + "%</td>";
                                PStr3 += dv[i]["CP_Point"].ToString().Trim() + "  " + dv[i]["CP_Desc"].ToString().Trim() + "<br />";
                                rnum += 1;
                                break;
                            }
                            else
                            {
                                PStr1 += "<td></td>";
                                PStr2 += "<td></td>";
                                rnum += 1;
                            }
                        }

                        //該推動項目最後一筆
                        if (Int32.Parse(dv[i]["ItemCount"].ToString().Trim()) == PushItemCount)
                        {
                            //工作比重累加
                            WorkRatioNum += (dv[i]["P_WorkRatio"].ToString().Trim() != "") ? double.Parse(dv[i]["P_WorkRatio"].ToString().Trim()) : 0;
                            //每筆row的後面未填滿時要補滿
                            if (rnum != dym.Rows.Count)
                            {
                                var mRange = dym.Rows.Count - rnum;
                                for (int a = 0; a < mRange; a++)
                                {
                                    PStr1 += "<td></td>";
                                    PStr2 += "<td></td>";
                                }
                            }
                            //結尾
                            PStr1 += "</tr>";
                            PStr2 += "</tr>";
                            PStr3 = "<tr><td align='center'>查核點<br />進度說明</td><td colspan='" + dym.Rows.Count + "'>" + PStr3 + "</td></tr>";
                            tmpstr += PStr1 + PStr2 + PStr3;
                            //清空
                            PStr1 = string.Empty;
                            PStr2 = string.Empty;
                            PStr3 = string.Empty;
                            rnum = 0;
                            PushItemCount = 1;
                        }
                        else
                            PushItemCount++;
                    }
                    else
                    {
                        for (int j = 0; j < dym.Rows.Count; j++)
                        {
                            if (dym.Rows[j]["Y"].ToString().Trim() == dv[i]["CP_ReserveYear"].ToString().Trim() && dym.Rows[j]["M"].ToString().Trim() == dv[i]["CP_ReserveMonth"].ToString().Trim())
                            {
                                PStr1 += "<td align='center'>" + dv[i]["CP_Point"].ToString().Trim() + "</td>";
                                PStr2 += "<td align='center'>" + dv[i]["CP_Process"].ToString().Trim() + "%</td>";
                                PStr3 += dv[i]["CP_Point"].ToString().Trim() + "  " + dv[i]["CP_Desc"].ToString().Trim() + "<br />";
                                rnum += 1;
                                break;
                            }

                            //判斷是否中間有跳空的月份
                            if (rnum == j)
                            {
                                PStr1 += "<td></td>";
                                PStr2 += "<td></td>";
                                rnum += 1;
                            }
                        }
                        
                        //該推動項目最後一筆
                        if (Int32.Parse(dv[i]["ItemCount"].ToString().Trim()) == PushItemCount)
                        {
                            //工作比重累加
                            WorkRatioNum += (dv[i]["P_WorkRatio"].ToString().Trim() != "") ? double.Parse(dv[i]["P_WorkRatio"].ToString().Trim()) : 0;
                            //每筆row的後面未填滿時要補滿
                            if (rnum != dym.Rows.Count)
                            {
                                var mRange = dym.Rows.Count - rnum;
                                for (int a = 0; a < mRange; a++)
                                {
                                    PStr1 += "<td></td>";
                                    PStr2 += "<td></td>";
                                }
                            }
                            //結尾
                            PStr1 += "</tr>";
                            PStr2 += "</tr>";
                            PStr3 = "<tr><td align='center'>查核點<br />進度說明</td><td colspan='" + dym.Rows.Count + "'>" + PStr3 + "</td></tr>";
                            tmpstr += PStr1 + PStr2 + PStr3;
                            //清空
                            PStr1 = string.Empty;
                            PStr2 = string.Empty;
                            PStr3 = string.Empty;
                            rnum = 0;
                            PushItemCount = 1;
                        }
                        else
                            PushItemCount++;
                    }

                }

                //合計
                tmpstr += "<tr><td align='center'>合計</td><td align='center'>" + WorkRatioNum.ToString() + "%</td><td align='center'>累計預定進度(%)</td>" + PStrTotal + "</tr>";
            }
            tmpstr += "</table>";
        }
        return tmpstr;
    }

    private string ROC_Date(string dateStr)
    {
        string TaiwanDate = string.Empty;
        if (dateStr == "")
            TaiwanDate = "";
        else
        {
            var ROC_Calendar = new System.Globalization.TaiwanCalendar();
            TaiwanDate = ROC_Calendar.GetYear(DateTime.Parse(dateStr)).ToString().PadLeft(3, '0') + "年" +
                    DateTime.Parse(dateStr).Month.ToString().PadLeft(1, '0') + "月" +
                    DateTime.Parse(dateStr).Day.ToString().PadLeft(1, '0') + "日";
        }
        return TaiwanDate;
    }


    private string getPItemNa(string item)
    {
        string str = string.Empty;
        DataTable dt = cp_db.getPushItemName(item);
        if (dt.Rows.Count > 0)
            str = dt.Rows[0]["C_Item_cn"].ToString();
        return str;
    }

    private string getExType_Cn(string item)
    {
        string str = string.Empty;
        DataTable dt = ct_db.getCn("09", item);
        if (dt.Rows.Count > 0)
            str = dt.Rows[0]["C_Item_cn"].ToString();
        return str;
    }
}