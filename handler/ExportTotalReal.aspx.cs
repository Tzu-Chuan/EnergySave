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

public partial class handler_ExportTotalReal : System.Web.UI.Page
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
            u_sheet.SetColumnWidth(2, (int)((58 + 0.72) * 256));
            u_sheet.SetColumnWidth(3, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(4, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(5, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(6, (int)((11 + 0.72) * 256));
            u_sheet.SetColumnWidth(7, (int)((11 + 0.72) * 256));

            //設定表頭
            u_row.CreateCell(0).SetCellValue("縣市");
            u_row.CreateCell(1).SetCellValue("季");
            u_row.CreateCell(2).SetCellValue("預算狀態");
            u_row.CreateCell(3).SetCellValue("經費實支率%");
            
            //u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 3, 6));
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 0, 0));//建立跨越2列(共2列 1~2)  ，跨越0欄(共0欄 A~A)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 1, 1));//建立跨越2列(共2列 1~2)  ，跨越0欄(共0欄 B~B)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 2, 2));//建立跨越2列(共2列 1~2)  ，跨越0欄(共0欄 C~C)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 3, 7));//建立跨越2列(共0列 1~1)  ，跨越4欄(共5欄 D~H)

            //指定欄位style
            u_sheet.GetRow(0).GetCell(0).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(1).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(2).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(3).CellStyle = cs_center;

            IRow u_row2 = u_sheet.CreateRow(1);// 在工作表裡面，產生一列。
            u_row2.CreateCell(3).SetCellValue("節電基礎");
            u_row2.CreateCell(4).SetCellValue("因地制宜");
            u_row2.CreateCell(5).SetCellValue("設備汰換");
            u_row2.CreateCell(6).SetCellValue("擴大補助");
            u_row2.CreateCell(7).SetCellValue("整體");

            //指定欄位style
            u_sheet.GetRow(1).GetCell(3).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(4).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(5).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(6).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(7).CellStyle = cs_center;


            //******************* 內容 star *******************//
            double sumM = 0, sumR = 0;
            string sum01RR = "", sum02RR = "", sum03RR = "", sum04RR = "", sM = "", sR = "";
            string strStage = Request.QueryString["s"].ToString().Trim();
            ch_db._strStage = strStage;
            dt = ch_db.getReportReal();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i]["RS_Type01Money"].ToString().Trim() != "") {
                        sumM += Convert.ToDouble(dt.Rows[i]["RS_Type01Money"].ToString().Trim());
                        sM = sumM.ToString().Trim();
                    }

                    if (dt.Rows[i]["RS_Type02Money"].ToString().Trim() != "") {
                        sumM += Convert.ToDouble(dt.Rows[i]["RS_Type02Money"].ToString().Trim());
                        sM = sumM.ToString().Trim();
                    }

                    if (dt.Rows[i]["RS_Type03Money"].ToString().Trim() != "") {
                        sumM += Convert.ToDouble(dt.Rows[i]["RS_Type03Money"].ToString().Trim());
                        sM = sumM.ToString().Trim();
                    }
                    if (dt.Rows[i]["RS_Type04Money"].ToString().Trim() != "")
                    {
                        sumM += Convert.ToDouble(dt.Rows[i]["RS_Type04Money"].ToString().Trim());
                        sM = sumM.ToString().Trim();
                    }


                    if (dt.Rows[i]["RS_Type01Real"].ToString().Trim() != "") {
                        sumR += Convert.ToDouble(dt.Rows[i]["RS_Type01Real"].ToString().Trim());
                        sR = sumR.ToString().Trim();
                    }
                        
                    if (dt.Rows[i]["RS_Type02Real"].ToString().Trim() != "") {
                        sumR += Convert.ToDouble(dt.Rows[i]["RS_Type02Real"].ToString().Trim());
                        sR = sumR.ToString().Trim();
                    }
                        
                    if (dt.Rows[i]["RS_Type03Real"].ToString().Trim() != "") {
                        sumR += Convert.ToDouble(dt.Rows[i]["RS_Type03Real"].ToString().Trim());
                        sR = sumR.ToString().Trim();
                    }

                    if (dt.Rows[i]["RS_Type04Real"].ToString().Trim() != "")
                    {
                        sumR += Convert.ToDouble(dt.Rows[i]["RS_Type04Real"].ToString().Trim());
                        sR = sumR.ToString().Trim();
                    }

                    if (dt.Rows[i]["RS_Type01RealRate"].ToString().Trim() != "")
                        sum01RR = dt.Rows[i]["RS_Type01RealRate"].ToString().Trim();
                    if (dt.Rows[i]["RS_Type02RealRate"].ToString().Trim() != "")
                        sum02RR = dt.Rows[i]["RS_Type02RealRate"].ToString().Trim();
                    if (dt.Rows[i]["RS_Type03RealRate"].ToString().Trim() != "")
                        sum03RR = dt.Rows[i]["RS_Type03RealRate"].ToString().Trim();
                    if (dt.Rows[i]["RS_Type04RealRate"].ToString().Trim() != "")
                        sum04RR = dt.Rows[i]["RS_Type04RealRate"].ToString().Trim();

                    u_row = u_sheet.CreateRow(i + 2);    // 在工作表裡面，產生一列。
                    u_row.CreateCell(0).SetCellValue(dt.Rows[i]["citycn"].ToString().Trim());
                    if (dt.Rows[i]["RS_Year"].ToString().Trim() != "" && dt.Rows[i]["RS_Season"].ToString().Trim() != "")
                    {
                        u_row.CreateCell(1).SetCellValue(dt.Rows[i]["RS_Year"].ToString().Trim() + "年第" + dt.Rows[i]["RS_Season"].ToString().Trim() + "季");
                    }
                    else
                    {
                        u_row.CreateCell(1).SetCellValue("");
                    }
                    u_row.CreateCell(2).SetCellValue(dt.Rows[i]["RS_CostDesc"].ToString().Trim());
                    u_row.CreateCell(3).SetCellValue(sum01RR);
                    u_row.CreateCell(4).SetCellValue(sum02RR);
                    u_row.CreateCell(5).SetCellValue(sum03RR);
                    u_row.CreateCell(6).SetCellValue(sum04RR);
                    if (sM!="" && Convert.ToDouble(sM)>0)
                    {
                        u_row.CreateCell(7).SetCellValue(Math.Round((Convert.ToDouble(sR) / Convert.ToDouble(sM))*100, 0));//四捨五入到整數位
                    }
                    else
                    {
                        u_row.CreateCell(7).SetCellValue("");
                    }
                    u_sheet.GetRow(i + 2).GetCell(0).CellStyle = cs_center;
                    u_sheet.GetRow(i + 2).GetCell(1).CellStyle = cs_center;
                    u_sheet.GetRow(i + 2).GetCell(2).CellStyle = cs_break;
                    u_sheet.GetRow(i + 2).GetCell(3).CellStyle = cs_right;
                    u_sheet.GetRow(i + 2).GetCell(4).CellStyle = cs_right;
                    u_sheet.GetRow(i + 2).GetCell(5).CellStyle = cs_right;
                    u_sheet.GetRow(i + 2).GetCell(6).CellStyle = cs_right;
                    u_sheet.GetRow(i + 2).GetCell(7).CellStyle = cs_right;
                    sum01RR = "";
                    sum02RR = "";
                    sum03RR = "";
                    sum04RR = "";
                    sM = "";
                    sR = "";
                    sumM = 0;
                    sumR = 0;
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