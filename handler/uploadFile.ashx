<%@ WebHandler Language="C#" Class="uploadFile" %>

using System;
using System.Web;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Web.SessionState;

public class uploadFile : IHttpHandler,IRequiresSessionState {
    public void ProcessRequest (HttpContext context) {
        try
        {
            if (LogInfo.mGuid == "")
            {
                context.Response.Write("reLogin");
                return;
            }

            if (context.Request.Files.Count > 0)
            {
                string ProjectGuid = (context.Request.Form["pGuid"] == null) ? "" : context.Request.Form["pGuid"].ToString();
                string type = (context.Request.Form["type"] == null) ? "" : context.Request.Form["type"].ToString();

                SqlConnection oConn = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
                oConn.Open();
                SqlCommand oCmd = new SqlCommand();
                oCmd.Connection = oConn;

                oCmd.Parameters.Add("@file_parentid", SqlDbType.NVarChar);
                oCmd.Parameters.Add("@file_type", SqlDbType.NVarChar);
                oCmd.Parameters.Add("@file_orgname", SqlDbType.NVarChar);
                oCmd.Parameters.Add("@file_encryname", SqlDbType.NVarChar);
                oCmd.Parameters.Add("@file_exten", SqlDbType.NVarChar);
                oCmd.Parameters.Add("@file_size", SqlDbType.NVarChar);
                oCmd.Parameters.Add("@file_status", SqlDbType.NVarChar);

                HttpFileCollection files = context.Request.Files;
                HttpPostedFile file = files[0];

                string UpLoadPath = ConfigurationManager.AppSettings["UploadFileRootDir"];

                //副檔名
                string extension = System.IO.Path.GetExtension(file.FileName);
                //原檔名
                string orgName = System.IO.Path.GetFileName(file.FileName).Replace(extension, "");
                //檔案大小
                int file_size = file.ContentLength;
                //取得TIME與GUID
                string timeguid = DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + Guid.NewGuid().ToString("N");
                //儲存的名稱
                string newName = timeguid + extension;

                if (!Directory.Exists(UpLoadPath.Substring(0, UpLoadPath.LastIndexOf("\\"))))//如果上傳路徑中沒有該目錄，則自動新增
                {
                    Directory.CreateDirectory(UpLoadPath.Substring(0, UpLoadPath.LastIndexOf("\\")));
                }

                if (extension.ToLower() == ".doc" || extension.ToLower() == ".docx" || extension.ToLower() == ".xls" || extension.ToLower() == ".xlsx" || extension.ToLower() == ".ppt" ||
                        extension.ToLower() == ".pptx" || extension.ToLower() == ".pdf")
                {
                    file.SaveAs(UpLoadPath + newName);
                    //進資料庫前, 儲存名稱要去除路徑
                    newName = newName.Replace(extension, "");

                    //寫入原本檔案紀錄表
                    oCmd.CommandText = @"INSERT INTO FileTable (
                    file_parentid,
                    file_type,
                    file_orgname,
                    file_encryname,
                    file_exten,
                    file_size,
                    file_status
                    ) VALUES (
                    @file_parentid, 
                    @file_type, 
                    @file_orgname, 
                    @file_encryname,
                    @file_exten,
                    @file_size,
                    @file_status
                    ) ";
                    oCmd.Parameters["@file_parentid"].Value = ProjectGuid;
                    oCmd.Parameters["@file_type"].Value = type;
                    oCmd.Parameters["@file_orgname"].Value = orgName;
                    oCmd.Parameters["@file_encryname"].Value = newName;
                    oCmd.Parameters["@file_exten"].Value = extension;
                    oCmd.Parameters["@file_size"].Value = file_size;
                    oCmd.Parameters["@file_status"].Value = "A";
                    oCmd.ExecuteNonQuery();
                    oConn.Close();
                    oCmd.Connection.Close();
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    public bool IsReusable {
        get {
            return false;
        }
    }
}