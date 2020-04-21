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

public partial class handler_ExportHistorySeasonList : System.Web.UI.Page
{
    ReportCheck_DB rc_db = new ReportCheck_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        Response.Clear();
        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        //******************* 表頭 ******************//
        //string path = "../Template/擴大補助.xlsx";
        string path = Server.MapPath("~/Template/季報歷史資料列表.xlsx");
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

            string str_stage = string.IsNullOrEmpty(Request.QueryString["s"]) ? "" : Request.QueryString["s"].ToString().Trim();
            DataTable dt = new DataTable();
            rc_db._RC_Stage = str_stage;
            dt = rc_db.getHistorySeasonList();

            #region 資料
            int dataSrow = 1, dataScol = 0;//excel從第二列開始塞
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    u_sheet.CreateRow(i + dataSrow);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        u_sheet.GetRow(i + dataSrow).CreateCell(j + dataScol).SetCellValue(dt.Rows[i][j].ToString().Trim());
                        u_sheet.GetRow(i + dataSrow).GetCell(j + dataScol).CellStyle = cs_center;//套用置中style

                    }
                }
            }
            #endregion

            workbook.Write(ms);
            string fileName = "第" + str_stage + "期季報歷史資料列表" + DateTime.Now.ToString("yyyyMMddHHmmss");// DateTime.Now.ToString("yyyyMMddHHmmss")
            string BrowserName = Request.Browser.Browser.ToLower();
            fileName = (BrowserName != "firefox") ? Server.UrlEncode(fileName + ".xlsx") : fileName + ".xlsx"; // firefox 就愛跟別人不一樣
            Response.AddHeader("Content-Disposition", "attachment;filename=\"" + fileName + "\"");//設定utf8 防止中文檔名亂碼
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