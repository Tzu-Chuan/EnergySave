<%@ WebHandler Language="C#" Class="deleteFile" %>

using System;
using System.Web;

public class deleteFile : IHttpHandler {
    File_DB f_db = new File_DB();
    public void ProcessRequest (HttpContext context) {
        try
        {
            string id = (context.Request["id"] != null) ? context.Request["id"].ToString() : "";

            f_db._file_id = id;
            f_db.DeleteFile();

            context.Response.Write("succeed");
        }
        catch (Exception ex) { context.Response.Write("Error：" + ex.Message.Replace("'", "\"")); }
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}