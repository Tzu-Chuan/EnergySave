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

public partial class handler_ExportTotalLag : System.Web.UI.Page
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
            //自動換行
            ICellStyle notesStyle = workbook.CreateCellStyle();
            notesStyle.WrapText = true;//設置換行這個要先設置
            notesStyle.VerticalAlignment = VerticalAlignment.Center;//垂直
            notesStyle.Alignment = HorizontalAlignment.Left;//水平

            //******************* 表頭 ******************//
            IRow u_row = u_sheet.CreateRow(0);// 在工作表裡面，產生一列。
            //(int)((50 + 0.72) * 256) 50
            //(int)((100 + 0.72) * 256)) 100
            //設定欄寬
            u_sheet.SetColumnWidth(0, (int)((10 + 0.72) * 256));
            u_sheet.SetColumnWidth(1, (int)((13 + 0.72) * 256));
            u_sheet.SetColumnWidth(2, (int)((50 + 0.72) * 256));
            u_sheet.SetColumnWidth(3, (int)((50 + 0.72) * 256));
            u_sheet.SetColumnWidth(4, (int)((50 + 0.72) * 256));
            u_sheet.SetColumnWidth(5, (int)((50 + 0.72) * 256));

            //設定表頭
            u_row.CreateCell(0).SetCellValue("縣市");
            u_row.CreateCell(1).SetCellValue("季");
            u_row.CreateCell(2).SetCellValue("節電基礎工作");
            u_row.CreateCell(3).SetCellValue("因地制宜");
            u_row.CreateCell(4).SetCellValue("設備汰換與智慧用電");
            u_row.CreateCell(5).SetCellValue("擴大補助");

            //指定欄位style
            u_sheet.GetRow(0).GetCell(0).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(1).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(2).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(3).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(4).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(5).CellStyle = cs_center;

            //******************* 內容 star *******************//
            string strStage = Request.QueryString["s"].ToString().Trim();
            ch_db._strStage = strStage;
            dt = ch_db.getReportTotalLog();
            string why1 = string.Empty;
            string why2 = string.Empty;
            string why3 = string.Empty;
            string whyEx = string.Empty;
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    u_row = u_sheet.CreateRow(i + 1);    // 在工作表裡面，產生一列。
                    u_row.CreateCell(0).SetCellValue(dt.Rows[i]["C_Item_cn"].ToString().Trim());
                    if (dt.Rows[i]["RS_Year"].ToString().Trim() != "" && dt.Rows[i]["RS_Season"].ToString().Trim() != "")
                    {
                        u_row.CreateCell(1).SetCellValue(dt.Rows[i]["RS_Year"].ToString().Trim() + "年第" + dt.Rows[i]["RS_Season"].ToString().Trim() + "季");
                    }
                    else
                    {
                        u_row.CreateCell(1).SetCellValue("");
                    }
                    //str1 = dt.Rows[i]["書審"].ToString().Trim().Replace("\\n", Environment.NewLine);//遇到\n就換行

                    why1= splitval(dt.Rows[i]["RS_01Why"].ToString().Trim()).Replace("\\n", Environment.NewLine);//遇到\n就換行
                    why2= splitval(dt.Rows[i]["RS_02Why"].ToString().Trim()).Replace("\\n", Environment.NewLine);//遇到\n就換行
                    why3= splitval(dt.Rows[i]["RS_03Why"].ToString().Trim()).Replace("\\n", Environment.NewLine);//遇到\n就換行
                    whyEx= splitval(dt.Rows[i]["RS_ExWhy"].ToString().Trim()).Replace("\\n", Environment.NewLine);//遇到\n就換行
                    u_row.CreateCell(2).SetCellValue(why1);
                    u_row.CreateCell(3).SetCellValue(why2);
                    u_row.CreateCell(4).SetCellValue(why3);
                    u_row.CreateCell(5).SetCellValue(whyEx);
                    u_sheet.GetRow(i + 1).GetCell(0).CellStyle = cs_center;
                    u_sheet.GetRow(i + 1).GetCell(1).CellStyle = cs_center;
                    u_sheet.GetRow(i + 1).GetCell(2).CellStyle = notesStyle;
                    u_sheet.GetRow(i + 1).GetCell(3).CellStyle = notesStyle;
                    u_sheet.GetRow(i + 1).GetCell(4).CellStyle = notesStyle;
                    u_sheet.GetRow(i + 1).GetCell(5).CellStyle = notesStyle;


                }
            }
            //******************* 內容 end *******************//

            workbook.Write(ms);
            string fileName = "計畫執行進度遭遇困難第" + strStage + "期" + DateTime.Now.ToString("yyyyMMddHHmmss");
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

    private string splitval(string str){
        string strVal = "";//回傳回去的字串
        string[] splitVal ;//split的字串
        if (str != "" && str != null)
        {
            str = str.Replace("\\n","\n");
            splitVal = str.Split('\n');
            for (var i = 0; i < splitVal.Length; i++)
            {
                if (splitVal[i].ToString().Trim() != "")
                {
                    if (strVal == "")
                    {
                        strVal += splitVal[i].ToString().Trim();
                    }
                    else
                    {
                        strVal += "\n" + splitVal[i].ToString().Trim();
                    }
                }
            }
        }
        strVal=strVal.Replace("&#x0D;", "");

        return strVal;
    }
}