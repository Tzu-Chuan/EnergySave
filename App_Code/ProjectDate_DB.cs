using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;

/// <summary>
/// ProjectDate_DB 的摘要描述
/// </summary>
public class ProjectDate_DB
{
    string KeyWord = string.Empty;
    public string _KeyWord
    {
        set { KeyWord = value; }
    }
    #region 私用
    string PD_ID = string.Empty;
    string PD_Name = string.Empty;
    string PD_Type = string.Empty;
    string PD_ModId = string.Empty;

    DateTime PD_StartDate;
    DateTime PD_EndDate;
    DateTime PD_ModDate;
    #endregion
    #region 公用
    public string _PD_ID
    {
        set { PD_ID = value; }
    }
    public string _PD_Name
    {
        set { PD_Name = value; }
    }
    public string _PD_Type
    {
        set { PD_Type = value; }
    }
    public string _PD_ModId
    {
        set { PD_ModId = value; }
    }
    public DateTime _PD_StartDate
    {
        set { PD_StartDate = value; }
    }
    public DateTime _PD_EndDate
    {
        set { PD_EndDate = value; }
    }
    public DateTime _PD_ModDate
    {
        set { PD_ModDate = value; }
    }
    #endregion

    public DataTable SelectList()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select * from ProjectDate where 1=1  ");

        if (PD_Type != "")
            sb.Append(@"and PD_Type=@PD_Type ");

        sb.Append(@"order by PD_Type ");
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oCmd.Parameters.AddWithValue("@PD_Type", PD_Type);
        oda.Fill(ds);
        return ds;
    }

    public void setData()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"update ProjectDate set
PD_StartDate=@PD_StartDate,
PD_EndDate=@PD_EndDate,
PD_ModDate=@PD_ModDate,
PD_ModId=@PD_ModId
where PD_Type=@PD_Type
";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@PD_Type", PD_Type);
        if (PD_StartDate.Year != 1)
            oCmd.Parameters.AddWithValue("@PD_StartDate", PD_StartDate);
        else
            oCmd.Parameters.AddWithValue("@PD_StartDate", DBNull.Value);
        if (PD_EndDate.Year != 1)
            oCmd.Parameters.AddWithValue("@PD_EndDate", PD_EndDate);
        else
            oCmd.Parameters.AddWithValue("@PD_EndDate", DBNull.Value);
        oCmd.Parameters.AddWithValue("@PD_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@PD_ModId", PD_ModId);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }
}