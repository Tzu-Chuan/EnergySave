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

public partial class handler_ExportTotalBehind : System.Web.UI.Page
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

            //******************* 表頭 start ******************//
            IRow u_row = u_sheet.CreateRow(0);// 在工作表裡面，產生一列。
            //(int)((50 + 0.72) * 256) 50
            //(int)((100 + 0.72) * 256)) 100
            //設定欄寬
            u_sheet.SetColumnWidth(0, (int)((10 + 0.72) * 256));
            u_sheet.SetColumnWidth(1, (int)((13 + 0.72) * 256));
            u_sheet.SetColumnWidth(2, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(3, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(4, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(5, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(6, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(7, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(8, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(9, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(10, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(11, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(12, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(13, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(14, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(15, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(16, (int)((11 + 0.72) * 256));
            //設定表頭
            u_row.CreateCell(0).SetCellValue("縣市");
            u_row.CreateCell(1).SetCellValue("季");
            u_row.CreateCell(2).SetCellValue("無風館冷氣(KW)");
            u_row.CreateCell(5).SetCellValue("老舊辦公室照明(具)");
            u_row.CreateCell(8).SetCellValue("室內停車場智慧照明(盞)");
            u_row.CreateCell(11).SetCellValue("中型能管系統(套)");
            u_row.CreateCell(14).SetCellValue("大型能管系統(套)");

            //u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 3, 6));
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 0, 0));//建立跨越2列(共2列 1~2)  ，跨越0欄(共0欄 A~A)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 1, 1));//建立跨越2列(共2列 1~2)  ，跨越0欄(共0欄 B~B)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 2, 4));//建立跨越1列(共1列 1~1)  ，跨越3欄(共3欄 C~E)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 5, 7));//建立跨越1列(共1列 1~1)  ，跨越3欄(共3欄 F~H)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 8, 10));//建立跨越1列(共1列 1~1)  ，跨越3欄(共3欄 I~K)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 11, 13));//建立跨越1列(共1列 1~1)  ，跨越3欄(共3欄 L~N)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 14, 16));//建立跨越1列(共1列 1~1)  ，跨越3欄(共3欄 O~Q)

            //指定欄位style
            u_sheet.GetRow(0).GetCell(0).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(1).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(2).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(5).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(8).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(11).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(14).CellStyle = cs_center;
            IRow u_row2 = u_sheet.CreateRow(1);// 在工作表裡面，產生一列。
            u_row2.CreateCell(2).SetCellValue("規劃數");
            u_row2.CreateCell(3).SetCellValue("申請數");
            u_row2.CreateCell(4).SetCellValue("完成數");
            u_row2.CreateCell(5).SetCellValue("規劃數");
            u_row2.CreateCell(6).SetCellValue("申請數");
            u_row2.CreateCell(7).SetCellValue("完成數");
            u_row2.CreateCell(8).SetCellValue("規劃數");
            u_row2.CreateCell(9).SetCellValue("申請數");
            u_row2.CreateCell(10).SetCellValue("完成數");
            u_row2.CreateCell(11).SetCellValue("規劃數");
            u_row2.CreateCell(12).SetCellValue("申請數");
            u_row2.CreateCell(13).SetCellValue("完成數");
            u_row2.CreateCell(14).SetCellValue("規劃數");
            u_row2.CreateCell(15).SetCellValue("申請數");
            u_row2.CreateCell(16).SetCellValue("完成數");

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


            //******************* 表頭 end ******************//

            //******************* 內容 star *******************//
            string strStage = Request.QueryString["s"].ToString().Trim();
            ch_db._strStage = strStage;
            dt = ch_db.getReportTotalBehindg();
            double sum01 = 0.0, sum02 = 0.0, sum03 = 0.0, sum04 = 0, sum05 = 0;
            double sumC01 = 0.0, sumC02 = 0.0, sumC03 = 0.0, sumC04 = 0, sumC05 = 0;
            double sumF01 = 0.0, sumF02 = 0.0, sumF03 = 0.0, sumF04 = 0, sumF05 = 0;
            int j = 1;
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i]["city_Year"].ToString().Trim() != "" && dt.Rows[i]["city_Season"].ToString().Trim() != "")
                    {
                        j = j + 1;
                        u_row = u_sheet.CreateRow(j);    // 在工作表裡面，產生一列。
                        u_row.CreateCell(0).SetCellValue(dt.Rows[i]["city_Item_cn"].ToString().Trim());
                        //沒有季的資料 畫面就都只SHOW空白
                        u_row.CreateCell(1).SetCellValue(dt.Rows[i]["city_Year"].ToString().Trim() + "年第" + dt.Rows[i]["city_Season"].ToString().Trim() + "季");
                        u_row.CreateCell(2).SetCellValue(dt.Rows[i]["RM_Sum01"].ToString().Trim());
                        u_row.CreateCell(3).SetCellValue(dt.Rows[i]["RM_SumC01"].ToString().Trim());
                        u_row.CreateCell(4).SetCellValue(dt.Rows[i]["RM_SumF01"].ToString().Trim());
                        u_row.CreateCell(5).SetCellValue(dt.Rows[i]["RM_Sum02"].ToString().Trim());
                        u_row.CreateCell(6).SetCellValue(dt.Rows[i]["RM_SumC02"].ToString().Trim());
                        u_row.CreateCell(7).SetCellValue(dt.Rows[i]["RM_SumF02"].ToString().Trim());
                        u_row.CreateCell(8).SetCellValue(dt.Rows[i]["RM_Sum03"].ToString().Trim());
                        u_row.CreateCell(9).SetCellValue(dt.Rows[i]["RM_SumC03"].ToString().Trim());
                        u_row.CreateCell(10).SetCellValue(dt.Rows[i]["RM_SumF03"].ToString().Trim());
                        u_row.CreateCell(11).SetCellValue(dt.Rows[i]["RM_Sum04"].ToString().Trim());
                        u_row.CreateCell(12).SetCellValue(dt.Rows[i]["RM_SumC04"].ToString().Trim());
                        u_row.CreateCell(13).SetCellValue(dt.Rows[i]["RM_SumF04"].ToString().Trim());
                        u_row.CreateCell(14).SetCellValue(dt.Rows[i]["RM_Sum05"].ToString().Trim());
                        u_row.CreateCell(15).SetCellValue(dt.Rows[i]["RM_SumC05"].ToString().Trim());
                        u_row.CreateCell(16).SetCellValue(dt.Rows[i]["RM_SumF05"].ToString().Trim());

                        sum01 += (dt.Rows[i]["RM_Sum01"].ToString().Trim() == "") ? 0.0 : Convert.ToDouble(dt.Rows[i]["RM_Sum01"].ToString().Trim());
                        sum02 += (dt.Rows[i]["RM_Sum02"].ToString().Trim() == "") ? 0.0 : Convert.ToDouble(dt.Rows[i]["RM_Sum02"].ToString().Trim());
                        sum03 += (dt.Rows[i]["RM_Sum03"].ToString().Trim() == "") ? 0.0 : Convert.ToDouble(dt.Rows[i]["RM_Sum03"].ToString().Trim());
                        sum04 += (dt.Rows[i]["RM_Sum04"].ToString().Trim() == "") ? 0 : Convert.ToDouble(dt.Rows[i]["RM_Sum04"].ToString().Trim());
                        sum05 += (dt.Rows[i]["RM_Sum05"].ToString().Trim() == "") ? 0 : Convert.ToDouble(dt.Rows[i]["RM_Sum05"].ToString().Trim());
                        sumC01 += (dt.Rows[i]["RM_SumC01"].ToString().Trim() == "") ? 0.0 : Convert.ToDouble(dt.Rows[i]["RM_SumC01"].ToString().Trim());
                        sumC02 += (dt.Rows[i]["RM_SumC02"].ToString().Trim() == "") ? 0.0 : Convert.ToDouble(dt.Rows[i]["RM_SumC02"].ToString().Trim());
                        sumC03 += (dt.Rows[i]["RM_SumC03"].ToString().Trim() == "") ? 0.0 : Convert.ToDouble(dt.Rows[i]["RM_SumC03"].ToString().Trim());
                        sumC04 += (dt.Rows[i]["RM_SumC04"].ToString().Trim() == "") ? 0 : Convert.ToDouble(dt.Rows[i]["RM_SumC04"].ToString().Trim());
                        sumC05 += (dt.Rows[i]["RM_SumC05"].ToString().Trim() == "") ? 0 : Convert.ToDouble(dt.Rows[i]["RM_SumC05"].ToString().Trim());
                        sumF01 += (dt.Rows[i]["RM_SumF01"].ToString().Trim() == "") ? 0.0 : Convert.ToDouble(dt.Rows[i]["RM_SumF01"].ToString().Trim());
                        sumF02 += (dt.Rows[i]["RM_SumF02"].ToString().Trim() == "") ? 0.0 : Convert.ToDouble(dt.Rows[i]["RM_SumF02"].ToString().Trim());
                        sumF03 += (dt.Rows[i]["RM_SumF03"].ToString().Trim() == "") ? 0.0 : Convert.ToDouble(dt.Rows[i]["RM_SumF03"].ToString().Trim());
                        sumF04 += (dt.Rows[i]["RM_SumF04"].ToString().Trim() == "") ? 0 : Convert.ToDouble(dt.Rows[i]["RM_SumF04"].ToString().Trim());
                        sumF05 += (dt.Rows[i]["RM_SumF05"].ToString().Trim() == "") ? 0 : Convert.ToDouble(dt.Rows[i]["RM_SumF05"].ToString().Trim());

                        //指定欄位style
                        u_sheet.GetRow(j).GetCell(0).CellStyle = cs_center;
                        u_sheet.GetRow(j).GetCell(1).CellStyle = cs_center;
                        u_sheet.GetRow(j).GetCell(2).CellStyle = cs_right;
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

                        //最後一筆要合計
                        if (i == dt.Rows.Count - 1)
                        {
                            u_row = u_sheet.CreateRow(j+1);
                            u_row.CreateCell(0).SetCellValue("合計");
                            u_row.CreateCell(1).SetCellValue("");
                            u_row.CreateCell(2).SetCellValue(sum01);
                            u_row.CreateCell(3).SetCellValue(sumC01);
                            u_row.CreateCell(4).SetCellValue(sumF01);
                            u_row.CreateCell(5).SetCellValue(sum02);
                            u_row.CreateCell(6).SetCellValue(sumC02);
                            u_row.CreateCell(7).SetCellValue(sumF02);
                            u_row.CreateCell(8).SetCellValue(sum03);
                            u_row.CreateCell(9).SetCellValue(sumC03);
                            u_row.CreateCell(10).SetCellValue(sumF03);
                            u_row.CreateCell(11).SetCellValue(sum04);
                            u_row.CreateCell(12).SetCellValue(sumC04);
                            u_row.CreateCell(13).SetCellValue(sumF04);
                            u_row.CreateCell(14).SetCellValue(sum05);
                            u_row.CreateCell(15).SetCellValue(sumC05);
                            u_row.CreateCell(16).SetCellValue(sumF05);

                            //指定欄位style
                            u_sheet.GetRow(j+1).GetCell(0).CellStyle = cs_center;
                            u_sheet.GetRow(j+1).GetCell(1).CellStyle = cs_center;
                            u_sheet.GetRow(j+1).GetCell(2).CellStyle = cs_right;
                            u_sheet.GetRow(j+1).GetCell(3).CellStyle = cs_right;
                            u_sheet.GetRow(j+1).GetCell(4).CellStyle = cs_right;
                            u_sheet.GetRow(j+1).GetCell(5).CellStyle = cs_right;
                            u_sheet.GetRow(j+1).GetCell(6).CellStyle = cs_right;
                            u_sheet.GetRow(j+1).GetCell(7).CellStyle = cs_right;
                            u_sheet.GetRow(j+1).GetCell(8).CellStyle = cs_right;
                            u_sheet.GetRow(j+1).GetCell(9).CellStyle = cs_right;
                            u_sheet.GetRow(j+1).GetCell(10).CellStyle = cs_right;
                            u_sheet.GetRow(j+1).GetCell(11).CellStyle = cs_right;
                            u_sheet.GetRow(j+1).GetCell(12).CellStyle = cs_right;
                            u_sheet.GetRow(j+1).GetCell(13).CellStyle = cs_right;
                            u_sheet.GetRow(j+1).GetCell(14).CellStyle = cs_right;
                            u_sheet.GetRow(j+1).GetCell(15).CellStyle = cs_right;
                            u_sheet.GetRow(j+1).GetCell(16).CellStyle = cs_right;
                        }
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
                    //    u_row.CreateCell(14).SetCellValue("");
                    //    u_row.CreateCell(15).SetCellValue("");
                    //    u_row.CreateCell(16).SetCellValue("");
                    //}

                    
                }
            }
            //******************* 內容 end *******************//

            workbook.Write(ms);
            string fileName = "各縣市申請數" + strStage + "期" + DateTime.Now.ToString("yyyyMMddHHmmss");
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