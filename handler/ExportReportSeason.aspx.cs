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
using System.Xml;
using System.Text;

public partial class handler_ExportReportSeason : System.Web.UI.Page
{
    Member_DB mb_db=new Member_DB();
    ReportSeasonV2_DB rs_db = new ReportSeasonV2_DB();
    ExpandFinish_DB ef_db = new ExpandFinish_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        string newName= Guid.NewGuid().ToString("N").Substring(0, 10) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

        string FileName=newName;
        Aspose.Words.Document doc=new Aspose.Words.Document(Server.MapPath("~/Template/ReportSeason.docx"));
        DocumentBuilder builder=new DocumentBuilder(doc);
        builder.MoveToDocumentEnd(); //移至Word最下方

        if (Request.QueryString["v"] != null && Request.QueryString["tp"] != null)
        {
            string ext = (Request.QueryString["tp"].ToString().ToLower() == "word") ? ".docx" : ".pdf";
            string id = string.IsNullOrEmpty(Request.QueryString["v"]) ? "" : Common.Decrypt(Request.QueryString["v"].ToString());
            string f_city = string.Empty;
            string f_year = string.Empty;
            string f_season = string.Empty;
            string strHtml = string.Empty;

            rs_db._RS_ID= id;
            DataTable dt = rs_db.getExportSeasonDetail();
            if (dt.Rows.Count > 0)
            {
                f_city = dt.Rows[0]["City"].ToString().Trim();
                f_year = dt.Rows[0]["RS_Year"].ToString().Trim();
                f_season = dt.Rows[0]["RS_Season"].ToString().Trim();

                #region 基本資料
                strHtml += "<table width='100%' border='1' cellspacing='0' cellpadding='0'>";
                strHtml += "<tr>";
                strHtml += "<td style='text-align:center; width:17%;'>提報年份</td>";
                strHtml += "<td style='width:17%;'>" + dt.Rows[0]["RS_Year"].ToString().Trim() + "年</td>";
                strHtml += "<td style='text-align:center; width:17%;'>提報季別</td>";
                strHtml += "<td style='width:17%;'>第" + NumToChinese(dt.Rows[0]["RS_Season"].ToString().Trim()) + "季</td>";
                strHtml += "<td style='text-align:center; width:17%;'>期程</td>";
                strHtml += "<td style='width:17%;'>第" + NumToChinese(dt.Rows[0]["RS_Stage"].ToString().Trim()) + "期</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td style='text-align:center;'>執行機關</td>";
                strHtml += "<td>" + dt.Rows[0]["City"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:center;'>承辦局處</td>";
                strHtml += "<td colspan='3'>" + dt.Rows[0]["M_Office"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td style='text-align:center;'>執行期程</td>";
                strHtml += "<td colspan='3'>開始：" + dt.Rows[0]["RS_StartDay"].ToString().Trim() + " ~ 結束：" + dt.Rows[0]["RS_EndDay"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:center;'>合計</td>";
                strHtml += "<td>" + dt.Rows[0]["RS_TotalMonth"].ToString().Trim() + "月</td>";
                strHtml += "</tr>";
                strHtml += "</table>";
                #endregion

                #region 預算狀態
                strHtml += "<div style='font-size:16pt;'><b>預算狀態</b></div>";
                builder.InsertHtml(strHtml);
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Font.Bold = false;
                builder.Font.Size = 12;
                builder.Write(dt.Rows[0]["RS_CostDesc"].ToString().Trim());
                builder.Write("\n");
                #endregion

                #region 經費實支數
                builder.Font.Bold = true;
                builder.Font.Size = 18;
                builder.Write("經費實支數");
                builder.Write("\n");
                strHtml = "<table width='100%' border='1' cellspacing='0' cellpadding='0'>";
                strHtml += "<tr>";
                strHtml += "<td style='text-align:center; width:25%;'>推動工作</td>";
                strHtml += "<td style='text-align:center; width:25%;'>經費(千元) A</td>";
                strHtml += "<td style='text-align:center; width:25%;'>實支數(千元) B</td>";
                strHtml += "<td style='text-align:center; width:25%;'>實支率(%) C = B / A</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td style='text-align:center;'>節電基礎工作</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_Type01Money"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_Type01Real"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_Type01RealRate"].ToString().Trim() + "%</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td style='text-align:center;'>因地制宜</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_Type02Money"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_Type02Real"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_Type02RealRate"].ToString().Trim() + "%</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td style='text-align:center;'>設備汰換及智慧用電</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_Type03Money"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_Type03Real"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_Type03RealRate"].ToString().Trim() + "%</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td style='text-align:center;'>擴大補助</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_Type04Money"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_Type04Real"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_Type04RealRate"].ToString().Trim() + "%</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td style='text-align:center;'>整體工作</td>";
                strHtml += "<td style='text-align:right;'>" + CountMoney(dt, "allMoney") + "</td>";
                strHtml += "<td style='text-align:right;'>" + CountMoney(dt, "allReal") + "</td>";
                strHtml += "<td style='text-align:right;'>" + CountMoney(dt, "Ratio") + "%</td>";
                strHtml += "</tr>";
                strHtml += "</table>";
                #endregion

                #region 整體進度
                strHtml += "<div style='font-size:16pt;'><b>整體進度</b></div>";
                builder.InsertHtml(strHtml);
                builder.Font.Bold = false;
                builder.Font.Size = 12;
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                builder.Write("預定執行進度：" + dt.Rows[0]["RS_AllSchedule"].ToString().Trim() + "%");
                builder.Write("\t\t\t\t\t\t\t實際執行進度：" + dt.Rows[0]["RS_AllRealSchedule"].ToString().Trim() + "%");
                builder.Write("\n");
                #endregion

                #region 查核點&預定工作進度
                builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                builder.Font.Bold = true;
                builder.Font.Size = 16;
                builder.Write("壹、節電基礎工作");
                builder.Font.Color = Color.Blue;
                builder.Write("（本季預定：" + dt.Rows[0]["RS_01Schedule"].ToString().Trim() + "%；本季實際 " + dt.Rows[0]["RS_01RealSchedule"].ToString().Trim() + "%）");
                builder.Font.Color = Color.Black;
                builder.Write("\n");
                if (dt.Rows[0]["RS_PushItemDesc"].ToString() == "")
                    builder.InsertHtml(CheckPointHtml_Ver1(dt.Rows[0]["RS_CheckPointData"].ToString().Trim(), "01"));
                else
                    builder.InsertHtml(CheckPointHtml_Ver2(dt.Rows[0]["RS_CheckPointData"].ToString().Trim(), dt.Rows[0]["RS_PushItemDesc"].ToString(), "01"));

                builder.Write("貳、因地制宜");
                builder.Font.Color = Color.Blue;
                builder.Write("（本季預定：" + dt.Rows[0]["RS_02Schedule"].ToString().Trim() + "%；本季實際 " + dt.Rows[0]["RS_02RealSchedule"].ToString().Trim() + "%）");
                builder.Font.Color = Color.Black;
                builder.Write("\n");
                if (dt.Rows[0]["RS_PushItemDesc"].ToString() == "")
                    builder.InsertHtml(CheckPointHtml_Ver1(dt.Rows[0]["RS_CheckPointData"].ToString().Trim(), "02"));
                else
                    builder.InsertHtml(CheckPointHtml_Ver2(dt.Rows[0]["RS_CheckPointData"].ToString().Trim(), dt.Rows[0]["RS_PushItemDesc"].ToString(), "02"));

                builder.Write("參、設備汰換與智慧用電");
                builder.Font.Color = Color.Blue;
                builder.Write("（本季預定：" + dt.Rows[0]["RS_03Schedule"].ToString().Trim() + "%；本季實際 " + dt.Rows[0]["RS_03RealSchedule"].ToString().Trim() + "%）");
                builder.Font.Color = Color.Black;
                builder.Write("\n");
                if (dt.Rows[0]["RS_PushItemDesc"].ToString() == "")
                    builder.InsertHtml(CheckPointHtml_Ver1(dt.Rows[0]["RS_CheckPointData"].ToString().Trim(), "03"));
                else
                    builder.InsertHtml(CheckPointHtml_Ver2(dt.Rows[0]["RS_CheckPointData"].ToString().Trim(), dt.Rows[0]["RS_PushItemDesc"].ToString(), "03"));

                #region 設備汰換預計完成數
                strHtml = "<table width='100%' border='1' cellspacing='0' cellpadding='0'>";
                strHtml += "<thead>";
                strHtml += "<tr>";
                strHtml += "<th nowrap='nowrap' style='width:50%;'>項目</th>";
                strHtml += "<th nowrap='nowrap' style='width:25%;'>本期預計完成數</th>";
                strHtml += "<th nowrap='nowrap' style='width:25%;'>本期累計完成數</th>";
                strHtml += "</tr>";
                strHtml += "</thead>";
                strHtml += "<tbody>";
                strHtml += "<tr>";
                strHtml += "<td style='text-align:center;'>無風管冷氣(kW)註：每台冷氣約4kW</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_03Type01S"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_03Type01C"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td style='text-align:center;'>老舊辦公室照明(具)</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_03Type02S"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_03Type02C"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td style='text-align:center;'>室內停車場智慧照明(盞)</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_03Type03S"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_03Type03C"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td style='text-align:center;'>中型能管系統(套)</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_03Type04S"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_03Type04C"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "<tr>";
                strHtml += "<td style='text-align:center;'>大型能管系統(套)</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_03Type05S"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right;'>" + dt.Rows[0]["RS_03Type05C"].ToString().Trim() + "</td>";
                strHtml += "</tr>";
                strHtml += "</tbody>";
                strHtml += "</table>";
                builder.InsertHtml(strHtml);
                #endregion

                builder.Write("肆、擴大補助");
                builder.Font.Color = Color.Blue;
                builder.Write("（本季預定：" + dt.Rows[0]["RS_04Schedule"].ToString().Trim() + "%；本季實際 " + dt.Rows[0]["RS_04RealSchedule"].ToString().Trim() + "%）");
                builder.Font.Color = Color.Black;
                builder.Write("\n");
                if (dt.Rows[0]["RS_PushItemDesc"].ToString() == "")
                    builder.InsertHtml(CheckPointHtml_Ver1(dt.Rows[0]["RS_CheckPointData"].ToString().Trim(), "04"));
                else
                    builder.InsertHtml(CheckPointHtml_Ver2(dt.Rows[0]["RS_CheckPointData"].ToString().Trim(), dt.Rows[0]["RS_PushItemDesc"].ToString(), "04"));

                #region 擴大補助預計完成數
                XmlDocument xDoc = new XmlDocument();
                xDoc.LoadXml(dt.Rows[0]["RS_CheckPointData"].ToString().Trim());
                XmlNodeList pNode = xDoc.SelectNodes("/cpList/PushItem[@P_Type='04']");
                //查詢累計完成數資料
                DataTable ExDt = ef_db.GetDataByRSID(id);
                strHtml = "<table width='100%' border='1' cellspacing='0' cellpadding='0'>";
                strHtml += "<thead>";
                strHtml += "<tr>";
                strHtml += "<th nowrap='nowrap' style='width:50%;'>項目</th>";
                strHtml += "<th nowrap='nowrap' style='width:25%;'>本期預計完成數</th>";
                strHtml += "<th nowrap='nowrap' style='width:25%;'>本期累計完成數</th>";
                strHtml += "</tr>";
                strHtml += "</thead>";
                strHtml += "<tbody>";
                if (pNode.Count > 0)
                {
                    for (int i = 0; i < pNode.Count; i++)
                    {
                        strHtml += "<tr>";
                        strHtml += "<td style='text-align:center;'>" + pNode[i].Attributes[3].Value + "</td>";
                        strHtml += "<td style='text-align:right;'>" + pNode[i].Attributes[5].Value + "</td>";
                        if (ExDt.Rows.Count > 0)
                        {
                            for (int a = 0; a < ExDt.Rows.Count; a++)
                            {
                                if (ExDt.Rows[a]["EF_PushitemId"].ToString() == pNode[i].Attributes[0].Value)
                                {
                                    strHtml += "<td style='text-align:right;'>" + ExDt.Rows[a]["EF_Finish"].ToString() + "</td>";
                                }
                            }
                            strHtml += "</tr>";
                        }
                    }
                }
                strHtml += "</tbody>";
                strHtml += "</table>";
                builder.InsertHtml(strHtml);
                #endregion
                #endregion

                #region 承辦人&審核資訊
                strHtml = "<table width='100%' border='0' cellspacing='0' cellpadding='0'>";
                strHtml += "<tr>";
                strHtml += "<td style='text-align:right; width:15%;'>填表人：</td>";
                strHtml += "<td style='width:17%;'>" + dt.Rows[0]["Member"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right; width:17%;'>電話：</td>";
                strHtml += "<td style='width:17%;'>" + dt.Rows[0]["Tel"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right; width:17%;'>填表日期：</td>";
                strHtml += "<td style='width:17%;'>" + DateTime.Parse(dt.Rows[0]["WriteDate"].ToString().Trim()).ToString("yyyy/MM/dd") + "</td>";
                strHtml += "</tr><tr>";
                strHtml += "<td style='text-align:right;'>主管：</td>";
                strHtml += "<td colspan='2'>" + dt.Rows[0]["Manager"].ToString().Trim() + "</td>";
                strHtml += "<td style='text-align:right;'>審核日期：</td>";
                string checkday = (dt.Rows[0]["CheckDate"].ToString().Trim() == "") ? "" : DateTime.Parse(dt.Rows[0]["CheckDate"].ToString().Trim()).ToString("yyyy/MM/dd");
                strHtml += "<td colspan='2'>" + checkday + "</td>";
                strHtml += "</tr></table>";
                builder.InsertHtml(strHtml);
                #endregion
            }

            #region  變換整份word字型
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
            {
                run.Font.Name = "標楷體";
            }
            #endregion

            doc.Save(Server.MapPath("~/Template/" + newName + ext));
            Response.Clear();
            Response.ClearHeaders();

            FileName = "季報-(" + f_city + ")" + f_year + "年第" + f_season + "季";
            string BrowserName = Request.Browser.Browser.ToLower();
            FileName = (BrowserName != "firefox") ? Server.UrlEncode(FileName + ext) : FileName + ext; // firefox 就愛跟別人不一樣
            Response.AddHeader("content-disposition", "attachment;filename=" + FileName);
            Response.ContentType="application/octet-stream";
            Response.WriteFile(Server.MapPath("~/Template/" + newName + ext));
            Response.Flush();
            File.Delete(Server.MapPath("~/Template/" + newName + ext));
            Response.End();
        }
    }

    #region 第一版
    private string CheckPointHtml_Ver1(string cpxml, string type)
    {
        string tmpHtml = string.Empty;
        XmlDocument xDoc = new XmlDocument();
        xDoc.LoadXml(cpxml);
        XmlNodeList pNode = xDoc.SelectNodes("/cpList/PushItem[@P_Type='" + type + "']");
        for (int i = 0; i < pNode.Count; i++)
        {
            string year_str = string.Empty; //年
            string month_str = "<tr>"; //月
            string cpstr = string.Empty; // 查核點
            string pstr = string.Empty; // 預定進度
            string realstr = string.Empty; // 實際進度
            string cpdesc = string.Empty; // 查核點進度說明
            int tmpCount = 0; // 年 colspan
            tmpHtml += "<div style='font-size:14pt;'><b>" + (i + 1) + "、" + pNode[i].Attributes[2].Value + "</b></div>";
            tmpHtml += "<div style='font-size:12pt;'>(1)執行進度</div>";
            tmpHtml += "<table width='100%' border='1' cellspacing='0' cellpadding='0'>";
            tmpHtml += "<thead><tr>";
            tmpHtml += "<th nowrap='nowrap' rowspan='2' style='width:150px;'>工作比重(%)</th>";
            tmpHtml += "<th nowrap='nowrap' rowspan='2' style='width:150px;'>年月<br />進度(%)</th>";
            for (int j = 0; j < pNode[i].ChildNodes.Count; j++)
            {
                XmlNode cp = pNode[i].ChildNodes[j];
                #region 年
                if (j == 0)
                    tmpCount += 1;
                else if (pNode[i].ChildNodes[j - 1].SelectSingleNode("CP_ReserveYear").InnerText != cp.SelectSingleNode("CP_ReserveYear").InnerText) //跨年時
                {
                    year_str += "<th colspan='" + tmpCount + "' style='text-align:center;'>" + pNode[i].ChildNodes[j - 1].SelectSingleNode("CP_ReserveYear").InnerText + "年</th>";
                    tmpCount = 1; //跨年需重置
                }
                else if (pNode[i].ChildNodes.Count == (j + 1)) //最後一筆資料
                {
                    tmpCount += 1; //最後一筆也要算
                    year_str += "<th colspan='" + tmpCount + "' style='text-align:center;'>" + cp.SelectSingleNode("CP_ReserveYear").InnerText + "年</th></tr>";
                }
                else
                    tmpCount += 1;
                #endregion

                #region 月
                month_str += "<th style='text - align:center;'>" + cp.SelectSingleNode("CP_ReserveMonth").InnerText + "月</th>";
                #endregion

                #region 執行進度-Body
                if (j == 0)
                {
                    cpstr += "<tr>";
                    cpstr += "<td rowspan='3' style='text-align:center;'>" + pNode[i].Attributes[3].Value + "%</td>";
                    cpstr += "<td style='text-align:center;'>查核點</td>";
                    cpstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_Point").InnerText + "</td>";
                    pstr += "<tr>";
                    pstr += "<td style='text-align:center;'>累計預定進度(%)</td>";
                    pstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_Process").InnerText + "%</td>";
                    realstr += "<tr>";
                    realstr += "<td style='text-align:center;'>累計實際進度(%)</td>";
                    realstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_RealProcess").InnerText + "</td>";
                }
                else if (pNode[i].ChildNodes.Count == (j + 1)) //最後一筆資料
                {
                    cpstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_Point").InnerText + "</td></tr>";
                    pstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_Process").InnerText + "%</td></tr>";
                    realstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_RealProcess").InnerText + "</td>";
                }
                else
                {
                    cpstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_Point").InnerText + "</td>";
                    pstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_Process").InnerText + "%</td>";
                    realstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_RealProcess").InnerText + "</td>";
                }
                #endregion

                #region 查核點進度說明-Body
                cpdesc += "<tr>";
                cpdesc += "<td>" + cp.SelectSingleNode("CP_Point").InnerText + "  " + cp.SelectSingleNode("CP_Desc").InnerText + "</td>";
                cpdesc += "<td>" + cp.SelectSingleNode("CP_Summary").InnerText.Replace("\r\n", "<br>") + "</td>";
                cpdesc += "<td>" + cp.SelectSingleNode("CP_BackwardDesc").InnerText.Replace("\r\n", "<br>") + "</td>";
                cpdesc += "</tr>";
                #endregion
            }
            tmpHtml += year_str + month_str;
            tmpHtml += "</tr></thead>";
            tmpHtml += "<tbody>" + cpstr + pstr + realstr + "</tbody>";
            tmpHtml += "</table></div>";
            tmpHtml += "<div class='font-size3 margin10T'>(2)查核點進度說明</div>";
            tmpHtml += "<div class='stripecomplex margin5T font-normal'>";
            tmpHtml += "<table width='100%' border='1' cellspacing='0' cellpadding='0'>";
            tmpHtml += "<thead><tr>";
            tmpHtml += "<th>查核點</th>";
            tmpHtml += "<th style='width:35%;'>辦理情形</th>";
            tmpHtml += "<th style='width:35%;'>進度差異說明</th>";
            tmpHtml += "</tr></thead>";
            tmpHtml += "<tbody>" + cpdesc + "</tbody>";
            tmpHtml += "</table></div>";
        }
        return tmpHtml;
    }
    #endregion

    #region 第二版(因架構不同，故兩版都留)
    private string CheckPointHtml_Ver2(string cpxml,string pdxml, string type)
    {
        string tmpHtml = string.Empty;
        XmlDocument xDoc = new XmlDocument();
        XmlDocument xDoc2 = new XmlDocument();
        ///CheckPoint
        xDoc.LoadXml(cpxml);
        XmlNodeList pNode = xDoc.SelectNodes("/cpList/PushItem[@P_Type='" + type + "']");
        ///PushItemDesc
        xDoc2.LoadXml(pdxml);
        for (int i = 0; i < pNode.Count; i++)
        {
            string year_str = string.Empty; //年
            string month_str = "<tr>"; //月
            string cpstr = string.Empty; // 查核點
            string pstr = string.Empty; // 預定進度
            string realstr = string.Empty; // 實際進度
            string cpdesc = string.Empty; // 查核點進度說明
            int tmpCount = 0; // 年 colspan
            tmpHtml += "<div style='font-size:14pt;'><b>" + (i + 1) + "、" + pNode[i].Attributes[3].Value + "</b></div>";
            tmpHtml += "<div style='font-size:12pt;'>(1)執行進度</div>";
            tmpHtml += "<table width='100%' border='1' cellspacing='0' cellpadding='0'>";
            tmpHtml += "<thead><tr>";
            tmpHtml += "<th nowrap='nowrap' rowspan='2' style='width:150px;'>工作比重(%)</th>";
            tmpHtml += "<th nowrap='nowrap' rowspan='2' style='width:150px;'>年月<br />進度(%)</th>";
            for (int j = 0; j < pNode[i].ChildNodes.Count; j++)
            {
                XmlNode cp = pNode[i].ChildNodes[j];
                #region 年
                if (j == 0)
                    tmpCount += 1;
                else if (pNode[i].ChildNodes[j - 1].SelectSingleNode("CP_ReserveYear").InnerText != cp.SelectSingleNode("CP_ReserveYear").InnerText) //跨年時
                {
                    year_str += "<th colspan='" + tmpCount + "' style='text-align:center;'>" + pNode[i].ChildNodes[j - 1].SelectSingleNode("CP_ReserveYear").InnerText + "年</th>";
                    tmpCount = 1; //跨年需重置
                }
                else if (pNode[i].ChildNodes.Count == (j + 1)) //最後一筆資料
                {
                    tmpCount += 1; //最後一筆也要算
                    year_str += "<th colspan='" + tmpCount + "' style='text-align:center;'>" + cp.SelectSingleNode("CP_ReserveYear").InnerText + "年</th></tr>";
                }
                else
                    tmpCount += 1;
                #endregion

                #region 月
                month_str += "<th style='text - align:center;'>" + cp.SelectSingleNode("CP_ReserveMonth").InnerText + "月</th>";
                #endregion

                #region 執行進度-Body
                if (j == 0)
                {
                    cpstr += "<tr>";
                    cpstr += "<td rowspan='3' style='text-align:center;'>" + pNode[i].Attributes[3].Value + "%</td>";
                    cpstr += "<td style='text-align:center;'>查核點</td>";
                    cpstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_Point").InnerText + "</td>";
                    pstr += "<tr>";
                    pstr += "<td style='text-align:center;'>累計預定進度(%)</td>";
                    pstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_Process").InnerText + "%</td>";
                    realstr += "<tr>";
                    realstr += "<td style='text-align:center;'>累計實際進度(%)</td>";
                    realstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_RealProcess").InnerText + "</td>";
                }
                else if (pNode[i].ChildNodes.Count == (j + 1)) //最後一筆資料
                {
                    cpstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_Point").InnerText + "</td></tr>";
                    pstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_Process").InnerText + "%</td></tr>";
                    realstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_RealProcess").InnerText + "</td>";
                }
                else
                {
                    cpstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_Point").InnerText + "</td>";
                    pstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_Process").InnerText + "%</td>";
                    realstr += "<td style='text-align:center;'>" + cp.SelectSingleNode("CP_RealProcess").InnerText + "</td>";
                }
                #endregion

                #region 查核點進度說明-Body
                //cpdesc += "<tr>";
                //cpdesc += "<td>" + cp.SelectSingleNode("CP_Point").InnerText + "  " + cp.SelectSingleNode("CP_Desc").InnerText + "</td>";
                //cpdesc += "<td>" + cp.SelectSingleNode("CP_Summary").InnerText.Replace("\r\n", "<br>") + "</td>";
                //cpdesc += "<td>" + cp.SelectSingleNode("CP_BackwardDesc").InnerText.Replace("\r\n", "<br>") + "</td>";
                //cpdesc += "</tr>";

                cpdesc += cp.SelectSingleNode("CP_Point").InnerText + "  " + cp.SelectSingleNode("CP_Desc").InnerText + "<br>";
                #endregion
            }
            tmpHtml += year_str + month_str;
            tmpHtml += "</tr></thead>";
            tmpHtml += "<tbody>" + cpstr + pstr + realstr + "</tbody>";
            tmpHtml += "</table></div>";
            tmpHtml += "<div class='font-size3 margin10T'>(2)查核點進度說明</div>";
            tmpHtml += "<div class='stripecomplex margin5T font-normal'>";
            tmpHtml += "<table width='100%' border='1' cellspacing='0' cellpadding='0'>";
            tmpHtml += "<thead><tr>";
            tmpHtml += "<th>查核點</th>";
            tmpHtml += "<th>年 季</th>";
            tmpHtml += "<th style='width:35%;'>辦理情形</th>";
            tmpHtml += "<th style='width:35%;'>進度差異說明</th>";
            tmpHtml += "</tr></thead>";
            /// 進度說明
            string pdstr = "";
            XmlNodeList pdNode = xDoc2.SelectNodes("/pdList/pd_item[@PD_PushitemGuid='" + pNode[i].Attributes[0].Value + "']");
            ///Rowspan
            int rspan_tmp = pdNode.Count;
            for (int j = 0; j < pdNode.Count; j++)
            {
                XmlNode pd = pdNode[j];
                pdstr += "<tr>";
                if (j == 0)
                    pdstr += "<td rowspan=" + rspan_tmp + ">" + cpdesc + "</td>";
                pdstr += "<td nowrap='nowrap' style='text-align:center;'>" + pd.Attributes[5].Value + "年<br>第" + pd.Attributes[6].Value + "季</td>";
                pdstr += "<td>" + pd.Attributes[7].Value + "</td>";
                pdstr += "<td>" + pd.Attributes[8].Value + "</td>";
                pdstr += "</tr>";
            }
            tmpHtml += "<tbody>" + pdstr + "</tbody>";
            tmpHtml += "</table></div>";
        }
        return tmpHtml;
    }
    #endregion

    //阿拉伯數字轉中文
    private string NumToChinese(string str)
    {
        string recVal = string.Empty;
        switch (str)
        {
            case "1":
                recVal = "一";
                break;
            case "2":
                recVal = "二";
                break;
            case "3":
                recVal = "三";
                break;
            case "4":
                recVal = "四";
                break;
        }
        return recVal;
    }

    private string CountMoney(DataTable dt,string item)
    {
        string rVal = string.Empty;
        //因為有些人擴大補助的部分不沒有填數字，不向設備汰換都是必填，所以才要先判斷掉
        string RS_Type01Money = (dt.Rows[0]["RS_Type01Money"].ToString() == "") ? "0" : dt.Rows[0]["RS_Type01Money"].ToString();
        string RS_Type02Money = (dt.Rows[0]["RS_Type02Money"].ToString() == "") ? "0" : dt.Rows[0]["RS_Type02Money"].ToString();
        string RS_Type03Money = (dt.Rows[0]["RS_Type03Money"].ToString() == "") ? "0" : dt.Rows[0]["RS_Type03Money"].ToString();
        string RS_Type04Money = (dt.Rows[0]["RS_Type04Money"].ToString() == "") ? "0" : dt.Rows[0]["RS_Type04Money"].ToString();
        string RS_Type01Real = (dt.Rows[0]["RS_Type01Real"].ToString() == "") ? "0" : dt.Rows[0]["RS_Type01Real"].ToString();
        string RS_Type02Real = (dt.Rows[0]["RS_Type02Real"].ToString() == "") ? "0" : dt.Rows[0]["RS_Type02Real"].ToString();
        string RS_Type03Real = (dt.Rows[0]["RS_Type03Real"].ToString() == "") ? "0" : dt.Rows[0]["RS_Type03Real"].ToString();
        string RS_Type04Real = (dt.Rows[0]["RS_Type04Real"].ToString() == "") ? "0" : dt.Rows[0]["RS_Type04Real"].ToString();
        //double totalMoney = double.Parse(dt.Rows[0]["RS_Type01Money"].ToString()) + double.Parse(dt.Rows[0]["RS_Type02Money"].ToString()) + double.Parse(dt.Rows[0]["RS_Type03Money"].ToString()) + double.Parse(dt.Rows[0]["RS_Type04Money"].ToString());
        //double totalReal = double.Parse(dt.Rows[0]["RS_Type01Real"].ToString()) + double.Parse(dt.Rows[0]["RS_Type02Real"].ToString()) + double.Parse(dt.Rows[0]["RS_Type03Real"].ToString()) + double.Parse(dt.Rows[0]["RS_Type04Real"].ToString());
        double totalMoney = double.Parse(RS_Type01Money) + double.Parse(RS_Type02Money) + double.Parse(RS_Type03Money) + double.Parse(RS_Type04Money);
        double totalReal = double.Parse(RS_Type01Real) + double.Parse(RS_Type02Real) + double.Parse(RS_Type03Real) + double.Parse(RS_Type04Real);

        double RealRatio = (totalReal / totalMoney) * 100;
        switch (item)
        {
            case "allMoney":
                rVal = totalMoney.ToString();
                break;
            case "allReal":
                rVal = totalReal.ToString();
                break;
            case "Ratio":
                rVal = Math.Round(RealRatio, 0).ToString();
                break;
        }
        return rVal;
    }
}