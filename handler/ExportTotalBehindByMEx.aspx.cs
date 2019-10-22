using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;
using NPOI.XSSF.UserModel;//-- XSSF 用來產生Excel 2007檔案（.xlsx）
using NPOI.SS.UserModel;//-- v.1.2.4起 新增的。

public partial class handler_ExportTotalBehindByMEx : System.Web.UI.Page
{
    Chart_DB ch_db = new Chart_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        string strStage = string.IsNullOrEmpty(Request.QueryString["s"]) ? "" : Request.QueryString["s"].ToString().Trim();
        string strType = string.IsNullOrEmpty(Request.QueryString["t"]) ? "" : Request.QueryString["t"].ToString().Trim();
        if (strStage != "") 
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
            notesStyle.Alignment = HorizontalAlignment.Center;//水平
            //小數點1位
            XSSFCellStyle cs_num1 = (XSSFCellStyle)workbook.CreateCellStyle();
            IDataFormat format1 = workbook.CreateDataFormat();
            cs_num1.DataFormat = format1.GetFormat("#,##0.0");
            //cs_num1.VerticalAlignment = VerticalAlignment.Center;//垂直
            //cs_num1.Alignment = HorizontalAlignment.Right;//水平

            DataSet ds = new DataSet();

            ch_db._strStage = strStage;
            ch_db._strExType = strType;
            ds = ch_db.getReportTotalBehindByMForEx();


            //******************* 表頭 ******************//
            IRow u_row = u_sheet.CreateRow(0);// 在工作表裡面，產生一列。
            //設定欄寬 設定表頭 指定欄位style
            //(int)((50 + 0.72) * 256) 50
            //(int)((100 + 0.72) * 256)) 100
            u_sheet.SetColumnWidth(0, (int)((10 + 0.72) * 256));
            u_sheet.SetColumnWidth(1, (int)((13 + 0.72) * 256));
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 0, 0));//建立跨越2列(共2列 1~2)  ，跨越0欄(共0欄 A~A)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 1, 1));//建立跨越2列(共2列 1~2)  ，跨越0欄(共0欄 A~A)
            u_row.CreateCell(0).SetCellValue("縣市");
            u_row.CreateCell(1).SetCellValue("月");
            u_sheet.GetRow(0).GetCell(0).CellStyle = cs_center;
            u_sheet.GetRow(0).GetCell(1).CellStyle = cs_center;

            int j = 2;//從第3欄開始
            if (ds.Tables[0].Rows.Count > 0)
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    u_sheet.SetColumnWidth(j, (int)((7 + 0.72) * 256));
                    u_row.CreateCell(j).SetCellValue(ds.Tables[0].Rows[i]["C_Item_cn"].ToString().Trim());
                    u_sheet.GetRow(0).GetCell(j).CellStyle = notesStyle;
                    u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, j, j + 2));//建立跨越0列  ，跨越3欄
                    j = j + 3;
                }

                u_row = u_sheet.CreateRow(1);// 在工作表裡面，產生一列。
                j = 2;
                for (int k = 0; k < ds.Tables[0].Rows.Count; k++)
                {
                    u_row.CreateCell(j).SetCellValue("規劃數");
                    u_row.CreateCell(j + 1).SetCellValue("申請數");
                    u_row.CreateCell(j + 2).SetCellValue("完成數");
                    u_sheet.GetRow(1).GetCell(j).CellStyle = cs_center;
                    u_sheet.GetRow(1).GetCell(j + 1).CellStyle = cs_center;
                    u_sheet.GetRow(1).GetCell(j + 2).CellStyle = cs_center;
                    j = j + 3;
                }
            }
            //******************* 內容 star *******************//

            int nowColumn = 2, nowRow = 1;
            string tmpType = "", tmpCity = "", tmpYear = "", tmpMonth = "";

            //合計 動態產生變數
            Dictionary<string, double> allKeys = new Dictionary<string, double>();

            if (ds.Tables[1].Rows.Count > 0)
            {
                for (int knum = 0; knum < ds.Tables[0].Rows.Count; knum++)
                {
                    string strkey = String.Format("sum_{0}", ds.Tables[0].Rows[knum]["C_Item"].ToString().Trim());
                    if (!allKeys.ContainsKey(strkey))
                    {
                        string strkey1 = String.Format("sum_{0}_1", ds.Tables[0].Rows[knum]["C_Item"].ToString().Trim());
                        string strkey2 = String.Format("sum_{0}_2", ds.Tables[0].Rows[knum]["C_Item"].ToString().Trim());
                        string strkey3 = String.Format("sum_{0}_3", ds.Tables[0].Rows[knum]["C_Item"].ToString().Trim());
                        allKeys.Add(strkey1, 0.0);
                        allKeys.Add(strkey2, 0.0);
                        allKeys.Add(strkey3, 0.0);
                    }
                }

                for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                {
                    tmpType = ds.Tables[1].Rows[i]["RM_CPType"].ToString().Trim();
                    if (tmpCity != "" && ds.Tables[1].Rows[i]["city_I_Guid"].ToString().Trim() == tmpCity && tmpYear == ds.Tables[1].Rows[i]["city_Year"].ToString().Trim() && tmpMonth == ds.Tables[1].Rows[i]["city_Month"].ToString().Trim())
                    {
                        nowColumn = 2;
                        tmpCity = ds.Tables[1].Rows[i]["city_I_Guid"].ToString().Trim();
                        for (int n = 0; n < ds.Tables[0].Rows.Count; n++)
                        {
                            string strkey1 = String.Format("sum_{0}_1", ds.Tables[0].Rows[n]["C_Item"].ToString().Trim());
                            string strkey2 = String.Format("sum_{0}_2", ds.Tables[0].Rows[n]["C_Item"].ToString().Trim());
                            string strkey3 = String.Format("sum_{0}_3", ds.Tables[0].Rows[n]["C_Item"].ToString().Trim());
                            if (ds.Tables[0].Rows[n]["C_Item"].ToString().Trim() == ds.Tables[1].Rows[i]["RM_CPType"].ToString().Trim())
                            {
                                u_row.CreateCell(nowColumn).SetCellValue(Convert.ToDouble(ds.Tables[1].Rows[i]["RM_Planning"].ToString().Trim()));
                                allKeys[strkey1] = allKeys[strkey1] + Convert.ToDouble(ds.Tables[1].Rows[i]["RM_Planning"].ToString().Trim());
                                //if (tmpType == "01" || tmpType == "03")
                                if (tmpType == "01")
                                    {
                                    u_row.CreateCell(nowColumn + 1).SetCellValue(Convert.ToDouble(ds.Tables[1].Rows[i]["sum3"].ToString().Trim()));
                                    u_row.CreateCell(nowColumn + 2).SetCellValue(Convert.ToDouble(ds.Tables[1].Rows[i]["sum4"].ToString().Trim()));
                                    allKeys[strkey2] = allKeys[strkey2] + Convert.ToDouble(ds.Tables[1].Rows[i]["sum3"].ToString().Trim());
                                    allKeys[strkey3] = allKeys[strkey3] + Convert.ToDouble(ds.Tables[1].Rows[i]["sum4"].ToString().Trim());
                                }
                                else
                                {
                                    u_row.CreateCell(nowColumn + 1).SetCellValue(Convert.ToDouble(ds.Tables[1].Rows[i]["sum1"].ToString().Trim()));
                                    u_row.CreateCell(nowColumn + 2).SetCellValue(Convert.ToDouble(ds.Tables[1].Rows[i]["sum2"].ToString().Trim()));
                                    allKeys[strkey2] = allKeys[strkey2] + Convert.ToDouble(ds.Tables[1].Rows[i]["sum1"].ToString().Trim());
                                    allKeys[strkey3] = allKeys[strkey3] + Convert.ToDouble(ds.Tables[1].Rows[i]["sum2"].ToString().Trim());
                                }

                            }
                            //else
                            //{
                            //    u_row.CreateCell(nowColumn).SetCellValue(0);
                            //    u_row.CreateCell(nowColumn + 1).SetCellValue(0);
                            //    u_row.CreateCell(nowColumn + 2).SetCellValue(0);
                            //}
                            //u_row.GetCell(nowColumn).CellStyle = cs_num1;
                            //u_row.GetCell(nowColumn + 1).CellStyle = cs_num1;
                            //u_row.GetCell(nowColumn + 2).CellStyle = cs_num1;
                            nowColumn = nowColumn + 3;
                        }
                    }
                    else
                    {
                        nowColumn = 2;
                        nowRow = nowRow + 1;
                        tmpCity = ds.Tables[1].Rows[i]["city_I_Guid"].ToString().Trim();
                        tmpYear = ds.Tables[1].Rows[i]["city_Year"].ToString().Trim();
                        tmpMonth = ds.Tables[1].Rows[i]["city_Month"].ToString().Trim();
                        u_row = u_sheet.CreateRow(nowRow);// 在工作表裡面，產生一列。
                        u_row.CreateCell(0).SetCellValue(ds.Tables[1].Rows[i]["city_Item_cn"].ToString().Trim());
                        if (ds.Tables[1].Rows[i]["city_Year"].ToString().Trim()=="" && ds.Tables[1].Rows[i]["city_Month"].ToString().Trim()=="") {
                            u_row.CreateCell(1).SetCellValue("");
                        }
                        else {
                            u_row.CreateCell(1).SetCellValue(ds.Tables[1].Rows[i]["city_Year"].ToString().Trim() + "年" + ds.Tables[1].Rows[i]["city_Month"].ToString().Trim() + "月");
                        }
                        
                        for (int n = 0; n < ds.Tables[0].Rows.Count; n++)
                        {
                            string strkey1 = String.Format("sum_{0}_1", ds.Tables[0].Rows[n]["C_Item"].ToString().Trim());
                            string strkey2 = String.Format("sum_{0}_2", ds.Tables[0].Rows[n]["C_Item"].ToString().Trim());
                            string strkey3 = String.Format("sum_{0}_3", ds.Tables[0].Rows[n]["C_Item"].ToString().Trim());
                            if (ds.Tables[0].Rows[n]["C_Item"].ToString().Trim() == ds.Tables[1].Rows[i]["RM_CPType"].ToString().Trim())
                            {
                                u_row.CreateCell(nowColumn).SetCellValue(Convert.ToDouble(ds.Tables[1].Rows[i]["RM_Planning"].ToString().Trim()));
                                allKeys[strkey1] = allKeys[strkey1] + Convert.ToDouble(ds.Tables[1].Rows[i]["RM_Planning"].ToString().Trim());
                                //if (tmpType == "01" || tmpType == "03")
                                if (tmpType == "01")
                                    {
                                    u_row.CreateCell(nowColumn + 1).SetCellValue(Convert.ToDouble(ds.Tables[1].Rows[i]["sum3"].ToString().Trim()));
                                    u_row.CreateCell(nowColumn + 2).SetCellValue(Convert.ToDouble(ds.Tables[1].Rows[i]["sum4"].ToString().Trim()));
                                    allKeys[strkey2] = allKeys[strkey2] + Convert.ToDouble(ds.Tables[1].Rows[i]["sum3"].ToString().Trim());
                                    allKeys[strkey3] = allKeys[strkey3] + Convert.ToDouble(ds.Tables[1].Rows[i]["sum4"].ToString().Trim());
                                }
                                else
                                {
                                    u_row.CreateCell(nowColumn + 1).SetCellValue(Convert.ToDouble(ds.Tables[1].Rows[i]["sum1"].ToString().Trim()));
                                    u_row.CreateCell(nowColumn + 2).SetCellValue(Convert.ToDouble(ds.Tables[1].Rows[i]["sum2"].ToString().Trim()));
                                    allKeys[strkey2] = allKeys[strkey2] + Convert.ToDouble(ds.Tables[1].Rows[i]["sum1"].ToString().Trim());
                                    allKeys[strkey3] = allKeys[strkey3] + Convert.ToDouble(ds.Tables[1].Rows[i]["sum2"].ToString().Trim());
                                }

                            }
                            else
                            {
                                u_row.CreateCell(nowColumn).SetCellValue(0);
                                u_row.CreateCell(nowColumn + 1).SetCellValue(0);
                                u_row.CreateCell(nowColumn + 2).SetCellValue(0);
                            }
                            u_row.GetCell(nowColumn).CellStyle = cs_num1;
                            u_row.GetCell(nowColumn + 1).CellStyle = cs_num1;
                            u_row.GetCell(nowColumn + 2).CellStyle = cs_num1;
                            nowColumn = nowColumn + 3;
                        }

                    }
                }
            }


            workbook.Write(ms);
            string fileName = "擴大補助-各縣市申請數(月累計)第" + strStage + "期" + DateTime.Now.ToString("yyyyMMddHHmmss");
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