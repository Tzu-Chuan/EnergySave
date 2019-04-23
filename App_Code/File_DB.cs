using System;
using System.Collections.Generic;
using System.Configuration;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;

/// <summary>
/// File_DB 的摘要描述
/// </summary>
public class File_DB
{
    #region file 私用
    string file_id = string.Empty;
    string file_parentid = string.Empty;
    string file_orgname = string.Empty;
    string file_encryname = string.Empty;
    string file_exten = string.Empty;
    string file_type = string.Empty;
    string file_status = string.Empty;
    string file_size = string.Empty;
    string file_desc = string.Empty;
    DateTime file_createdate;
    #endregion
    #region file 公用
    public string _file_id
    {
        set { file_id = value; }
    }
    public string _file_parentid
    {
        set { file_parentid = value; }
    }
    public string _file_orgname
    {
        set { file_orgname = value; }
    }
    public string _file_encryname
    {
        set { file_encryname = value; }
    }
    public string _file_exten
    {
        set { file_exten = value; }
    }
    public string _file_type
    {
        set { file_type = value; }
    }
    public string _file_status
    {
        set { file_status = value; }
    }
    public string _file_size
    {
        set { file_size = value; }
    }
    public string _file_desc
    {
        set { file_desc = value; }
    }
    public DateTime _file_createdate
    {
        set { file_createdate = value; }
    }
    #endregion

    public void InsertFile()
    {
        SqlConnection db = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        SqlCommand thisConnection = new SqlCommand();
        StringBuilder insert_value = new StringBuilder();
        thisConnection.Connection = db;

        try
        {
            thisConnection.Connection.Open();
            insert_value.Append(@"insert into FileTable ");
            insert_value.Append(@"(
file_parentid,
file_orgname,
file_encryname, 
file_exten, 
file_type,
file_size,
file_status
) ");
            insert_value.Append(@"values ");
            insert_value.Append(@"(
@file_parentid,
@file_orgname,
@file_encryname, 
@file_exten, 
@file_type,
@file_size,
@file_status
) ");
            thisConnection.CommandType = CommandType.Text;
            thisConnection.CommandText = insert_value.ToString();

            thisConnection.Parameters.AddWithValue("@file_parentid", file_parentid);
            thisConnection.Parameters.AddWithValue("@file_orgname", file_orgname);
            thisConnection.Parameters.AddWithValue("@file_encryname", file_encryname);
            thisConnection.Parameters.AddWithValue("@file_exten", file_exten);
            thisConnection.Parameters.AddWithValue("@file_type", file_type);
            thisConnection.Parameters.AddWithValue("@file_size", file_size);
            thisConnection.Parameters.AddWithValue("@file_status", file_status);


            thisConnection.ExecuteNonQuery();
            //insert_value.Remove(0, insert_value.Length);

        }
        finally
        {
            thisConnection.Connection.Close();
            thisConnection.Dispose();
        }
    }

    public DataTable SelectFile()
    {
        DataTable dt = new DataTable();
        SqlConnection thisConnection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        SqlCommand thisCommand = thisConnection.CreateCommand();
        SqlDataAdapter oda = new SqlDataAdapter();
        StringBuilder show_value = new StringBuilder();
        try
        {
            thisConnection.Open();
            show_value.Append(@"select * from FileTable where file_status='A' and file_parentid=@file_parentid and file_type=@file_type");
            thisCommand.CommandType = CommandType.Text;
            thisCommand.CommandText = show_value.ToString();
            thisCommand.Parameters.AddWithValue("@file_parentid", file_parentid);
            thisCommand.Parameters.AddWithValue("@file_type", file_type);
            //thisCommand.ExecuteNonQuery();
            oda.SelectCommand = thisCommand;
            oda.Fill(dt);
        }
        finally
        {
            oda.Dispose();
            thisConnection.Close();
            thisConnection.Dispose();
            thisCommand.Dispose();
        }
        return dt;
    }

    public DataTable getFileByID()
    {
        DataTable dt = new DataTable();
        SqlConnection thisConnection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        SqlCommand thisCommand = thisConnection.CreateCommand();
        SqlDataAdapter oda = new SqlDataAdapter();
        StringBuilder show_value = new StringBuilder();
        try
        {
            thisConnection.Open();
            show_value.Append(@"select * from FileTable where file_id=@file_id ");
            thisCommand.CommandType = CommandType.Text;
            thisCommand.CommandText = show_value.ToString();
            thisCommand.Parameters.AddWithValue("@file_id", file_id);
            //thisCommand.Parameters.AddWithValue("@file_type", file_type);
            //thisCommand.ExecuteNonQuery();
            oda.SelectCommand = thisCommand;
            oda.Fill(dt);
        }
        finally
        {
            oda.Dispose();
            thisConnection.Close();
            thisConnection.Dispose();
            thisCommand.Dispose();
        }
        return dt;
    }



    public void DeleteFile()
    {
        SqlConnection db = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        SqlCommand thisConnection = new SqlCommand();
        StringBuilder insert_value = new StringBuilder();
        thisConnection.Connection = db;

        try
        {
            thisConnection.Connection.Open();
            insert_value.Append(@"update FileTable set file_status='D' where file_id=@file_id ");

            thisConnection.CommandType = CommandType.Text;
            thisConnection.CommandText = insert_value.ToString();

            thisConnection.Parameters.AddWithValue("@file_id", file_id);

            thisConnection.ExecuteNonQuery();
        }
        finally
        {
            thisConnection.Connection.Close();
            thisConnection.Dispose();
        }
    }

    public void LogData(string pid,string type)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);

        oCmd.CommandText = @"update FileTable set file_status='D' where file_parentid=@file_parentid and file_type=@file_type and file_status='A' ";

        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@file_parentid", pid);
        oCmd.Parameters.AddWithValue("@file_type", type);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }
}