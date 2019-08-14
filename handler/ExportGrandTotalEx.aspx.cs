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

public partial class handler_ExportGrandTotalEx : System.Web.UI.Page
{
    Chart_DB ch_db = new Chart_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        string strStage = string.Empty;//期
        string strSdate = string.Empty;//開始年月
        string strEdate = string.Empty;//結束年月
        
        strStage = string.IsNullOrEmpty(Request.QueryString["s"]) ? "" : Request.QueryString["s"].ToString().Trim();
        strSdate = string.IsNullOrEmpty(Request.QueryString["sd"]) ? "" : Request.QueryString["sd"].ToString().Trim();
        strEdate = string.IsNullOrEmpty(Request.QueryString["ed"]) ? "" : Request.QueryString["ed"].ToString().Trim();

        //三個參數都必填 一定要有值
        if (strStage != "" && strSdate != "" && strEdate != "")
        {
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            //******************* 表頭 ******************//
            //string path = "../Template/擴大補助.xlsx";
            string path = Server.MapPath("~/Template/月報統計_擴大補助.xlsx");
            IWorkbook xlsfile = new XSSFWorkbook();

            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(file);//-- XSSF 用來產生Excel 2007檔案（.xlsx）
                ISheet u_sheet = workbook.GetSheetAt(0);
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
                notesStyle.Alignment = HorizontalAlignment.Center;//水平
                //小數點1位
                XSSFCellStyle cs_num1 = (XSSFCellStyle)workbook.CreateCellStyle();
                IDataFormat format1 = workbook.CreateDataFormat();
                cs_num1.DataFormat = format1.GetFormat("#,##0.0");
                cs_num1.VerticalAlignment = VerticalAlignment.Center;//垂直
                cs_num1.Alignment = HorizontalAlignment.Right;//水平
                //cs_num1.VerticalAlignment = VerticalAlignment.Center;//垂直
                //cs_num1.Alignment = HorizontalAlignment.Right;//水平

                DataSet ds = new DataSet();
                ch_db._strStage = strStage;
                ch_db._strSdate = strSdate;
                ch_db._strEdate = strEdate;
                ds = ch_db.getReportGrandTotalEx();

                #region 標題
                //月報類別
                u_sheet.GetRow(0).GetCell(1).SetCellValue("擴大補助");
                //查詢月份區間
                u_sheet.GetRow(1).GetCell(1).SetCellValue(ds.Tables[0].Rows[0]["strSdate"].ToString().Trim() + " ~ " + ds.Tables[0].Rows[0]["strEdate"].ToString().Trim() + "");

                int titleSrow = 3,titleScol=1;//從第四列第二欄開始塞title
                //Title
                if (ds.Tables[1].Rows.Count > 0)
                {
                    for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                    {
                        u_sheet.GetRow(titleSrow).GetCell(titleScol).SetCellValue(ds.Tables[1].Rows[i]["C_Item_cn"].ToString().Trim());//推動項目名稱
                        u_sheet.GetRow(titleSrow + 1).GetCell(titleScol).SetCellValue("規劃數");//規劃數
                        u_sheet.GetRow(titleSrow + 1).GetCell(titleScol + 1).SetCellValue("申請數");//申請數
                        u_sheet.GetRow(titleSrow + 1).GetCell(titleScol + 2).SetCellValue("完成數");//完成數
                        titleScol = titleScol + 3;
                    }
                }
                #endregion

                #region 縣市填報資料
                int dataSrow = 5, dataScol = 0;//excel從第六列開始塞
                if (ds.Tables[2].Rows.Count > 0)
                {
                    for (int i = 0; i < ds.Tables[2].Rows.Count; i++)
                    {
                        u_sheet.CreateRow(i + dataSrow);
                        for (int j = 0; j < ds.Tables[2].Columns.Count; j++)
                        {
                            
                            if (j == 0)
                            {
                                //縣市名稱 文字
                                u_sheet.GetRow(i + dataSrow).CreateCell(j + dataScol).SetCellValue(ds.Tables[2].Rows[i][j].ToString().Trim());
                            }
                            else {
                                //填報數 數字
                                u_sheet.GetRow(i + dataSrow).CreateCell(j + dataScol).SetCellValue(Convert.ToDouble(ds.Tables[2].Rows[i][j].ToString().Trim()));//轉乘小樹 這樣在Excel裡面會自動變數字
                                u_sheet.GetRow(i + dataSrow).GetCell(j + dataScol).CellStyle = cs_num1;//套用數字靠右Cell style
                                //u_sheet.AutoSizeColumn(j + dataScol);//設定自動調整欄寬  (加了自動調整欄寬 匯出會多好幾秒)
                            }
                        }
                    }
                }
                #endregion

                workbook.Write(ms);
                string fileName = "擴大補助-月報統計表第" + strStage + "期" + strSdate + "_" + strEdate;// DateTime.Now.ToString("yyyyMMddHHmmss")
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
}