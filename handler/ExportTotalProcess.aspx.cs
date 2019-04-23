using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;
using NPOI.XSSF.UserModel;//-- XSSF 用來產生Excel 2007檔案（.xlsx）
using NPOI.SS.UserModel;//-- v.1.2.4起 新增的。

public partial class handler_ExportTotalProcess : System.Web.UI.Page
{
    Chart_DB ch_db = new Chart_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        DataTable dt = new DataTable();
        if (Request.QueryString["s"] != null)
        {
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            XSSFWorkbook workbook = new XSSFWorkbook();//-- XSSF 用來產生Excel 2007檔案（.xlsx）
            ISheet u_sheet = workbook.CreateSheet("工作表一");
            MemoryStream ms = new MemoryStream();

            //設定共用style
            //style 置中
            XSSFCellStyle cs_center = (XSSFCellStyle)workbook.CreateCellStyle();
            cs_center.VerticalAlignment = VerticalAlignment.Center;//垂直
            cs_center.Alignment = HorizontalAlignment.Center;//水平
            //style 靠右
            XSSFCellStyle cs_right = (XSSFCellStyle)workbook.CreateCellStyle();
            cs_right.VerticalAlignment = VerticalAlignment.Center;//垂直
            cs_right.Alignment = HorizontalAlignment.Right;//水平
            //style 自動換行
            XSSFCellStyle cs_break = (XSSFCellStyle)workbook.CreateCellStyle();
            cs_break.VerticalAlignment = VerticalAlignment.Center;//垂直
            cs_break.Alignment = HorizontalAlignment.Left;//水平
            cs_break.WrapText = true;//自動換行

            //******************* 表頭 ******************//
            IRow u_row = u_sheet.CreateRow(0);// 在工作表裡面，產生一列。
            //(int)((50 + 0.72) * 256) 50
            //(int)((100 + 0.72) * 256)) 100
            //設定欄寬
            u_sheet.SetColumnWidth(0, (int)((10 + 0.72) * 256));
            u_sheet.SetColumnWidth(1, (int)((13 + 0.72) * 256));
            u_sheet.SetColumnWidth(2, (int)((8 + 0.72) * 256));
            u_sheet.SetColumnWidth(3, (int)((8 + 0.72) * 256));
            u_sheet.SetColumnWidth(4, (int)((8 + 0.72) * 256));
            u_sheet.SetColumnWidth(5, (int)((8 + 0.72) * 256));
            u_sheet.SetColumnWidth(6, (int)((8 + 0.72) * 256));
            u_sheet.SetColumnWidth(7, (int)((8 + 0.72) * 256));
            u_sheet.SetColumnWidth(8, (int)((8 + 0.72) * 256));
            u_sheet.SetColumnWidth(9, (int)((8 + 0.72) * 256));
            u_sheet.SetColumnWidth(10, (int)((8 + 0.72) * 256));
            u_sheet.SetColumnWidth(11, (int)((8 + 0.72) * 256));
            u_sheet.SetColumnWidth(12, (int)((8 + 0.72) * 256));
            u_sheet.SetColumnWidth(13, (int)((8 + 0.72) * 256));

            //設定表頭
            u_row.CreateCell(0).SetCellValue("縣市");
            u_row.CreateCell(1).SetCellValue("季");
            u_row.CreateCell(2).SetCellValue("節電基礎工作%");
            u_row.CreateCell(5).SetCellValue("因地制宜%");
            u_row.CreateCell(8).SetCellValue("設備汰換與智慧用電%");
            u_row.CreateCell(11).SetCellValue("擴大補助%");
            u_row.CreateCell(14).SetCellValue("整體%");

            //u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 3, 6));
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 0, 0));//建立跨越2列(共2列 1~2)  ，跨越0欄(共0欄 A~A)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 1, 1));//建立跨越2列(共2列 1~2)  ，跨越0欄(共0欄 B~B)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 2, 4));//建立跨越2列(共0列 1~1)  ，跨越3欄(共3欄 C~E)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 5, 7));//建立跨越2列(共0列 1~1)  ，跨越3欄(共3欄 F~H)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 8, 10));//建立跨越2列(共0列 1~1)  ，跨越3欄(共3欄 I~K)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 11, 13));//建立跨越2列(共0列 1~1)  ，跨越3欄(共3欄 L~N)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 14, 16));//建立跨越2列(共0列 1~1)  ，跨越3欄(共3欄 O~Q)

            //指定欄位style
            u_sheet.GetRow(0).GetCell(0).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(1).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(2).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(5).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(8).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(11).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(14).CellStyle = cs_center;

            IRow u_row2 = u_sheet.CreateRow(1);// 在工作表裡面，產生一列。
            u_row2.CreateCell(2).SetCellValue("預定");
            u_row2.CreateCell(3).SetCellValue("實際");
            u_row2.CreateCell(4).SetCellValue("差異");
            u_row2.CreateCell(5).SetCellValue("預定");
            u_row2.CreateCell(6).SetCellValue("實際");
            u_row2.CreateCell(7).SetCellValue("差異");
            u_row2.CreateCell(8).SetCellValue("預定");
            u_row2.CreateCell(9).SetCellValue("實際");
            u_row2.CreateCell(10).SetCellValue("差異");
            u_row2.CreateCell(11).SetCellValue("預定");
            u_row2.CreateCell(12).SetCellValue("實際");
            u_row2.CreateCell(13).SetCellValue("差異");
            u_row2.CreateCell(14).SetCellValue("預定");
            u_row2.CreateCell(15).SetCellValue("實際");
            u_row2.CreateCell(16).SetCellValue("差異");

            //指定欄位style
            u_sheet.GetRow(1).GetCell(2).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(3).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(4).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(5).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(6).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(7).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(8).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(9).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(10).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(11).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(12).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(13).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(14).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(15).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(16).CellStyle = cs_center;


            //******************* 內容 star *******************//
            string sum01S = "", sum01F = "", sum01SF = "";
            string sum02S = "", sum02F = "", sum02SF = "";
            string sum03S = "", sum03F = "", sum03SF = "";
            string sum04S = "", sum04F = "", sum04SF = "";
            int j = 1;
            string strStage = Request.QueryString["s"].ToString().Trim();
            ch_db._strStage = strStage;
            dt = ch_db.getReportProcess();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i]["RS_Year"].ToString().Trim() != "" && dt.Rows[i]["RS_Season"].ToString().Trim() != "")
                    {
                        j = j + 1;
                        sum01S = dt.Rows[i]["RS_Sum01S"].ToString().Trim();
                        sum02S = dt.Rows[i]["RS_Sum02S"].ToString().Trim();
                        sum03S = dt.Rows[i]["RS_Sum03S"].ToString().Trim();
                        sum04S = dt.Rows[i]["RS_Sum04S"].ToString().Trim();
                        sum01F = dt.Rows[i]["RS_Sum01F"].ToString().Trim();
                        sum02F = dt.Rows[i]["RS_Sum02F"].ToString().Trim();
                        sum03F = dt.Rows[i]["RS_Sum03F"].ToString().Trim();
                        sum04F = dt.Rows[i]["RS_Sum04F"].ToString().Trim();
                        sum01SF = dt.Rows[i]["RS_Sum01S_F"].ToString().Trim();
                        sum02SF = dt.Rows[i]["RS_Sum02S_F"].ToString().Trim();
                        sum03SF = dt.Rows[i]["RS_Sum03S_F"].ToString().Trim();
                        sum04SF = dt.Rows[i]["RS_Sum04S_F"].ToString().Trim();

                        u_row = u_sheet.CreateRow(j);    // 在工作表裡面，產生一列。
                        u_row.CreateCell(0).SetCellValue(dt.Rows[i]["citycn"].ToString().Trim());

                        //沒有季的資料 畫面就都只SHOW空白
                        u_row.CreateCell(1).SetCellValue(dt.Rows[i]["RS_Year"].ToString().Trim() + "年第" + dt.Rows[i]["RS_Season"].ToString().Trim() + "季");
                        if (sum01S != "0" && sum01S != "0.0" && sum01S != "0.00")
                        {
                            u_row.CreateCell(2).SetCellValue(sum01S.ToString());
                            u_row.CreateCell(3).SetCellValue(sum01F.ToString());
                            u_row.CreateCell(4).SetCellValue(sum01SF.ToString());
                        }else
                        {
                            u_row.CreateCell(2).SetCellValue("");
                            u_row.CreateCell(3).SetCellValue("");
                            u_row.CreateCell(4).SetCellValue("");
                        }
                        if (sum02S != "0" && sum02S != "0.0" && sum02S != "0.00")
                        {
                            u_row.CreateCell(5).SetCellValue(sum02S.ToString());
                            u_row.CreateCell(6).SetCellValue(sum02F.ToString());
                            u_row.CreateCell(7).SetCellValue(sum02SF.ToString());
                        }
                        else {
                            u_row.CreateCell(5).SetCellValue("");
                            u_row.CreateCell(6).SetCellValue("");
                            u_row.CreateCell(7).SetCellValue("");
                        }
                        if (sum03S != "0" && sum03S != "0.0" && sum03S != "0.00")
                        {
                            u_row.CreateCell(8).SetCellValue(sum03S.ToString());
                            u_row.CreateCell(9).SetCellValue(sum03F.ToString());
                            u_row.CreateCell(10).SetCellValue(sum03SF.ToString());
                        }
                        else {
                            u_row.CreateCell(8).SetCellValue("");
                            u_row.CreateCell(9).SetCellValue("");
                            u_row.CreateCell(10).SetCellValue("");
                        }
                        if (sum04S != "0" && sum04S != "0.0" && sum04S != "0.00")
                        {
                            u_row.CreateCell(11).SetCellValue(sum04S.ToString());
                            u_row.CreateCell(12).SetCellValue(sum04F.ToString());
                            u_row.CreateCell(13).SetCellValue(sum04SF.ToString());
                        }
                        else
                        {
                            u_row.CreateCell(11).SetCellValue("");
                            u_row.CreateCell(12).SetCellValue("");
                            u_row.CreateCell(13).SetCellValue("");
                        }

                        if ((sum01S == "0" || sum01S == "0.0" || sum01S == "0.00") && (sum02S == "0" || sum02S == "0.0" || sum02S == "0.00") && (sum03S == "0" || sum03S == "0.0" || sum03S == "0.00") && (sum04S == "0" || sum04S == "0.0" || sum04S == "0.00"))
                        {
                            u_row.CreateCell(14).SetCellValue("");
                            u_row.CreateCell(15).SetCellValue("");
                            u_row.CreateCell(16).SetCellValue("");
                        }
                        else
                        {
                            int intall = 0;
                            if (Convert.ToDouble(sum01S) > 0)
                            {
                                intall = intall + 1;
                            }
                            if (Convert.ToDouble(sum02S) > 0)
                            {
                                intall = intall + 1;
                            }
                            if (Convert.ToDouble(sum03S) > 0)
                            {
                                intall = intall + 1;
                            }
                            if (Convert.ToDouble(sum04S) > 0)
                            {
                                intall = intall + 1;
                            }

                            sum01S = (sum01S == "") ? "0" : sum01S;
                            sum02S = (sum02S == "") ? "0" : sum02S;
                            sum03S = (sum03S == "") ? "0" : sum03S;
                            sum04S = (sum04S == "") ? "0" : sum04S;
                            sum01F = (sum01F == "") ? "0" : sum01F;
                            sum02F = (sum02F == "") ? "0" : sum02F;
                            sum03S = (sum03S == "") ? "0" : sum03S;
                            sum04F = (sum04F == "") ? "0" : sum04F;
                            sum01SF = (sum01SF == "") ? "0" : sum01SF;
                            sum02SF = (sum02SF == "") ? "0" : sum02SF;
                            sum03SF = (sum03SF == "") ? "0" : sum03SF;
                            sum04SF = (sum04SF == "") ? "0" : sum04SF;

                            string vs = Math.Round(((Convert.ToDouble(sum01S) + Convert.ToDouble(sum02S) + Convert.ToDouble(sum03S) + Convert.ToDouble(sum04S)) / Convert.ToDouble(intall)), 2).ToString();//四捨五入到小數點第2位
                            string vf = Math.Round(((Convert.ToDouble(sum01F) + Convert.ToDouble(sum02F) + Convert.ToDouble(sum03F) + Convert.ToDouble(sum04F)) / Convert.ToDouble(intall)), 2).ToString();//四捨五入到小數點第2位
                            u_row.CreateCell(14).SetCellValue(vs.ToString());
                            u_row.CreateCell(15).SetCellValue(vf.ToString());
                            u_row.CreateCell(16).SetCellValue(Math.Round(Convert.ToDouble(vs) - Convert.ToDouble(vf), 2).ToString());//四捨五入到小數點第2位
                        }

                        u_sheet.GetRow(j).GetCell(0).CellStyle = cs_center;
                        u_sheet.GetRow(j).GetCell(1).CellStyle = cs_center;
                        u_sheet.GetRow(j).GetCell(2).CellStyle = cs_break;
                        u_sheet.GetRow(j).GetCell(3).CellStyle = cs_right;
                        u_sheet.GetRow(j).GetCell(4).CellStyle = cs_right;
                        u_sheet.GetRow(j).GetCell(5).CellStyle = cs_right;
                        u_sheet.GetRow(j).GetCell(6).CellStyle = cs_right;
                        u_sheet.GetRow(j).GetCell(7).CellStyle = cs_right;
                        u_sheet.GetRow(j).GetCell(8).CellStyle = cs_right;
                        u_sheet.GetRow(j).GetCell(9).CellStyle = cs_right;
                        u_sheet.GetRow(j).GetCell(10).CellStyle = cs_right;
                        u_sheet.GetRow(j).GetCell(11).CellStyle = cs_right;
                        u_sheet.GetRow(j).GetCell(12).CellStyle = cs_right;
                        u_sheet.GetRow(j).GetCell(13).CellStyle = cs_right;
                        u_sheet.GetRow(j).GetCell(14).CellStyle = cs_right;
                        u_sheet.GetRow(j).GetCell(15).CellStyle = cs_right;
                        u_sheet.GetRow(j).GetCell(16).CellStyle = cs_right;

                    }
                    //else
                    //{
                    //    u_row.CreateCell(1).SetCellValue("");
                    //    u_row.CreateCell(2).SetCellValue("");
                    //    u_row.CreateCell(3).SetCellValue("");
                    //    u_row.CreateCell(4).SetCellValue("");
                    //    u_row.CreateCell(5).SetCellValue("");
                    //    u_row.CreateCell(6).SetCellValue("");
                    //    u_row.CreateCell(7).SetCellValue("");
                    //    u_row.CreateCell(8).SetCellValue("");
                    //    u_row.CreateCell(9).SetCellValue("");
                    //    u_row.CreateCell(10).SetCellValue("");
                    //    u_row.CreateCell(11).SetCellValue("");
                    //    u_row.CreateCell(12).SetCellValue("");
                    //    u_row.CreateCell(13).SetCellValue("");
                    //}
                    
                }
            }
            //******************* 內容 end *******************//
            workbook.Write(ms);
            string fileName = "節電基礎及因地制宜工作進度摘要第" + strStage + "期" + DateTime.Now.ToString("yyyyMMddHHmmss");
            Response.AddHeader("Content-Disposition", "attachment;filename=\"" + HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8) + ".xlsx\"");//設定utf8 防止中文檔名亂碼
            //Response.AddHeader("Content-Disposition", String.Format("attachment;filename=" + fileName));
            Response.BinaryWrite(ms.ToArray());

            //== 釋放資源
            workbook = null;
            ms.Close();
            ms.Dispose();

            Response.Flush();
            Response.End();

        }
    }
}