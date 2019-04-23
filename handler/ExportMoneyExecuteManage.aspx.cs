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

public partial class handler_ExportMoneyExecuteManage : System.Web.UI.Page
{
    MoneyExecute_DB me_db = new MoneyExecute_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        DataSet ds = new DataSet();
        string stage_cn = string.Empty;
        string file_cityname = string.Empty;
        if (Request.QueryString["s"] != null && Request.QueryString["c"] != null)
        {
            switch (Request.QueryString["s"].ToString().Trim()) {
                case "1":
                    stage_cn = "一";
                    break;
                case "2":
                    stage_cn = "二";
                    break;
                case "3":
                    stage_cn = "三";
                    break;
            }
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
            cs_break.Alignment = HorizontalAlignment.Center;//水平
            cs_break.WrapText = true;//自動換行
            //style 自動換行
            XSSFCellStyle cs_breakLeft = (XSSFCellStyle)workbook.CreateCellStyle();
            cs_breakLeft.VerticalAlignment = VerticalAlignment.Center;//垂直
            cs_breakLeft.Alignment = HorizontalAlignment.Left;//水平
            cs_breakLeft.WrapText = true;//自動換行

            //同時三位一逗點不保留小數
            XSSFCellStyle cs_num = (XSSFCellStyle)workbook.CreateCellStyle();
            IDataFormat format = workbook.CreateDataFormat();
            cs_num.DataFormat = format.GetFormat("#,##0");


            //******************* 表頭 ******************//
            IRow u_row = u_sheet.CreateRow(0);// 在工作表裡面，產生一列。
            //(int)((50 + 0.72) * 256) 50
            //(int)((100 + 0.72) * 256)) 100
            //設定欄寬
            u_sheet.SetColumnWidth(0, (int)((7 + 0.72) * 256));
            u_sheet.SetColumnWidth(1, (int)((15 + 0.72) * 256));
            u_sheet.SetColumnWidth(2, (int)((25 + 0.72) * 256));
            u_sheet.SetColumnWidth(3, (int)((26 + 0.72) * 256));
            u_sheet.SetColumnWidth(4, (int)((15 + 0.72) * 256));
            u_sheet.SetColumnWidth(5, (int)((15 + 0.72) * 256));
            u_sheet.SetColumnWidth(6, (int)((15 + 0.72) * 256));
            u_sheet.SetColumnWidth(7, (int)((15 + 0.72) * 256));
            u_sheet.SetColumnWidth(8, (int)((40 + 0.72) * 256));
            //設定表頭
            u_row.CreateCell(0).SetCellValue("縣市");
            u_row.CreateCell(1).SetCellValue("");
            u_row.CreateCell(2).SetCellValue("期別");
            u_row.CreateCell(3).SetCellValue("");
            u_row.CreateCell(4).SetCellValue("");
            //u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 3, 6));
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 0, 4, 8));//建立跨越4欄(共0列 0~0)  ，跨越3欄(共0欄 E~H)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(1, 1, 5, 7));//建立跨越3欄(共0列 1~1)  ，跨越3欄(共0欄 E~G)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(2, 2, 5, 6));//建立跨越2欄(共0列 2~2)  ，跨越2欄(共0欄 E~F)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(1, 3, 0, 0));//建立跨越2列(共3列 2~4)  ，跨越0欄(共0欄 A~A)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(1, 3, 1, 1));//建立跨越2列(共3列 2~4)  ，跨越0欄(共0欄 B~B)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(1, 3, 2, 2));//建立跨越2列(共3列 2~4)  ，跨越0欄(共0欄 C~C)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(1, 3, 3, 3));//建立跨越2列(共3列 2~4)  ，跨越0欄(共0欄 D~D)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(1, 3, 4, 4));//建立跨越2列(共3列 2~4)  ，跨越0欄(共0欄 E~E)
            u_sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(1, 3, 8, 8));//建立跨越2列(共3列 2~4)  ，跨越0欄(共0欄 G~G)

            IRow u_row1 = u_sheet.CreateRow(1);// 在工作表裡面，產生一列。
            u_row1.CreateCell(0).SetCellValue("編號");
            u_row1.CreateCell(1).SetCellValue("執行機關");
            u_row1.CreateCell(2).SetCellValue("主責局處");
            u_row1.CreateCell(3).SetCellValue("計畫項目");
            u_row1.CreateCell(4).SetCellValue("金額\n(A)=(B)+(C)");
            u_row1.CreateCell(5).SetCellValue("處理方式");
            u_row1.CreateCell(8).SetCellValue("涉及措施");
            
            //指定欄位style
            u_sheet.GetRow(1).GetCell(0).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(1).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(2).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(3).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(4).CellStyle = cs_break;
            u_sheet.GetRow(1).GetCell(5).CellStyle = cs_center;
            u_sheet.GetRow(1).GetCell(8).CellStyle = cs_center;

            IRow u_row2 = u_sheet.CreateRow(2);// 在工作表裡面，產生一列。
            u_row2.CreateCell(5).SetCellValue("標案");
            u_row2.CreateCell(7).SetCellValue("自辦");

            //指定欄位style
            u_sheet.GetRow(2).GetCell(5).CellStyle = cs_center;
            u_sheet.GetRow(2).GetCell(7).CellStyle = cs_center;

            IRow u_row3 = u_sheet.CreateRow(3);// 在工作表裡面，產生一列。
            u_row3.CreateCell(5).SetCellValue("標案名稱");
            u_row3.CreateCell(6).SetCellValue("發包金額(元)(B)");
            u_row3.CreateCell(7).SetCellValue("金額(元)(C)");

            //指定欄位style
            u_sheet.GetRow(3).GetCell(5).CellStyle = cs_center;
            u_sheet.GetRow(3).GetCell(6).CellStyle = cs_center;
            u_sheet.GetRow(3).GetCell(7).CellStyle = cs_center;


            //******************* 內容 star *******************//
            int sumA = 0, sumB = 0, sumC = 0, intA = 0, intB = 0, intC = 0;
            string strStage = Request.QueryString["s"].ToString().Trim();
            string strCity = Request.QueryString["c"].ToString().Trim();
            me_db._PR_Stage = strStage;
            me_db._PR_City = strCity;
            ds = me_db.getMoneyList();

            //開始塞資料
            if (ds.Tables[0].Rows.Count > 0)
            {
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    if (i == 0)
                    {
                        if (Request.QueryString["c"].ToString().Trim() != "")
                        {
                            file_cityname = ds.Tables[0].Rows[0]["CityName"].ToString().Trim();
                        }else
                        {
                            file_cityname = "全部";
                        }
                        //縣市不是選全部才顯示經費
                        if (Request.QueryString["c"].ToString().Trim() != "")
                        {
                            u_row.CreateCell(1).SetCellValue(ds.Tables[0].Rows[i]["CityName"].ToString().Trim());
                            //String.Format("{0:0,0}",數字) = 加入三位一逗點
                            u_row.CreateCell(4).SetCellValue("經費：" + String.Format("{0:0,0}", Convert.ToDouble(ds.Tables[1].Rows[0]["MoneyAll"].ToString().Trim()) * 1000).ToString() + "(元)");
                        }
                        else
                        {
                            u_row.CreateCell(1).SetCellValue("全部");
                            u_row.CreateCell(3).SetCellValue("");
                        }
                        u_row.CreateCell(3).SetCellValue("第" + stage_cn + "期");
                    }

                    u_row = u_sheet.CreateRow(i + 4);    // 在工作表裡面，產生一列。
                    intA = 0;
                    intB = 0;
                    intC = 0;
                    if (ds.Tables[0].Rows[i]["PR_Money"].ToString().Trim() != "")
                    {
                        intA = Convert.ToInt32(ds.Tables[0].Rows[i]["PR_Money"].ToString().Trim());
                    }
                    if (ds.Tables[0].Rows[i]["PR_CaseMoney"].ToString().Trim() != "")
                    {
                        intB = Convert.ToInt32(ds.Tables[0].Rows[i]["PR_CaseMoney"].ToString().Trim());
                    }
                    if (ds.Tables[0].Rows[i]["PR_SelfMoney"].ToString().Trim() != "")
                    {
                        intC = Convert.ToInt32(ds.Tables[0].Rows[i]["PR_SelfMoney"].ToString().Trim());
                    }

                    u_row.CreateCell(0).SetCellValue((i + 1).ToString().Trim());//編號
                    u_row.CreateCell(1).SetCellValue(ds.Tables[0].Rows[i]["CityName"].ToString().Trim());
                    u_row.CreateCell(2).SetCellValue(ds.Tables[0].Rows[i]["PR_Office"].ToString().Trim());
                    u_row.CreateCell(3).SetCellValue(ds.Tables[0].Rows[i]["planNamer"].ToString().Trim());
                    u_row.CreateCell(4).SetCellValue(intA);
                    u_row.CreateCell(5).SetCellValue(ds.Tables[0].Rows[i]["PR_CaseName"].ToString().Trim());
                    u_row.CreateCell(6).SetCellValue(intB);
                    u_row.CreateCell(7).SetCellValue(intC);
                    u_row.CreateCell(8).SetCellValue(ds.Tables[0].Rows[i]["PR_Steps"].ToString().Trim());
                    
                    u_row.GetCell(0).CellStyle = cs_center;
                    u_row.GetCell(1).CellStyle = cs_center;
                    u_row.GetCell(2).CellStyle = cs_center;
                    u_row.GetCell(3).CellStyle = cs_center;
                    u_row.GetCell(4).CellStyle = cs_num;
                    u_row.GetCell(5).CellStyle = cs_center;
                    u_row.GetCell(6).CellStyle = cs_num;
                    u_row.GetCell(7).CellStyle = cs_num;
                    u_row.GetCell(8).CellStyle = cs_break;

                    sumA += intA;
                    sumB += intB;
                    sumC += intC;

                    //會後一筆資料顯示合計
                    if (i == ds.Tables[0].Rows.Count - 1)
                    {
                        u_row = u_sheet.CreateRow(i + 5);    // 在工作表裡面，產生一列。
                        u_row.CreateCell(3).SetCellValue("合計");
                        u_row.CreateCell(4).SetCellValue(sumA);
                        u_row.CreateCell(6).SetCellValue(sumB);
                        u_row.CreateCell(7).SetCellValue(sumC);
                        u_row.GetCell(3).CellStyle = cs_right;
                        u_row.GetCell(4).CellStyle = cs_num;
                        u_row.GetCell(6).CellStyle = cs_num;
                        u_row.GetCell(7).CellStyle = cs_num;
                    }
                }
            }
            //******************* 內容 end *******************//
            workbook.Write(ms);
            string fileName = "("+ file_cityname + ")經費執行第" + stage_cn + "期";
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