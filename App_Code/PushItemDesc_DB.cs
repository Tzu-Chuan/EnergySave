using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;

/// <summary>
/// PushItemDesc_DB 的摘要描述
/// </summary>
public class PushItemDesc_DB
{
    string KeyWord = string.Empty;
    public string _KeyWord
    {
        set { KeyWord = value; }
    }

    #region 公用
    string PD_ID = string.Empty;
    string PD_Guid = string.Empty;
    string PD_PushitemGuid = string.Empty;
    string PD_ProjectGuid = string.Empty;
    string PD_Stage = string.Empty;
    string PD_Year = string.Empty;
    string PD_Season = string.Empty;
    string PD_Summary = string.Empty;
    string PD_BackwardDesc = string.Empty;
    string PD_Status = string.Empty;
    #endregion
    #region 私用
    DateTime PD_CreateDate;
    DateTime PD_ModDate;

    public string _PD_ID
    {
        set { PD_ID = value; }
    }
    public string _PD_Guid
    {
        set { PD_Guid = value; }
    }
    public string _PD_PushitemGuid
    {
        set { PD_PushitemGuid = value; }
    }
    public string _PD_ProjectGuid
    {
        set { PD_ProjectGuid = value; }
    }
    public string _PD_Stage
    {
        set { PD_Stage = value; }
    }
    public string _PD_Year
    {
        set { PD_Year = value; }
    }
    public string _PD_Season
    {
        set { PD_Season = value; }
    }
    public string _PD_Summary
    {
        set { PD_Summary = value; }
    }
    public string _PD_BackwardDesc
    {
        set { PD_BackwardDesc = value; }
    }
    public string _PD_Status
    {
        set { PD_Status = value; }
    }
    public DateTime _PD_CreateDate
    {
        set { PD_CreateDate = value; }
    }
    public DateTime _PD_ModDate
    {
        set { PD_ModDate = value; }
    }
    #endregion

    public DataTable GetPiDesc(string mGuid)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
declare @PerGuid nvarchar(50)=@mGuid
declare @ProjectID nvarchar(50)=(select I_Guid from ProjectInfo where I_People=@PerGuid)

select * from PushItem_Desc
where PD_ProjectGuid=@ProjectID and PD_Year<=@PD_Year and PD_Season<=@PD_Season and PD_Stage<=@PD_Stage
order by PD_PushitemGuid,PD_Year,PD_Season,PD_Stage
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@mGuid", mGuid);
        oCmd.Parameters.AddWithValue("@PD_Year", PD_Year);
        oCmd.Parameters.AddWithValue("@PD_Season", PD_Season);
        oCmd.Parameters.AddWithValue("@PD_Stage", PD_Stage);
        oda.Fill(ds);
        return ds;
    }

    public void setPushitemDesc()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"
declare @rCount int = (select count(*)  from PushItem_Desc where PD_PushitemGuid=@PD_PushitemGuid and PD_Year=@PD_Year and PD_Season=@PD_Season and PD_Stage=@PD_Stage)

if @rCount>0
    begin
        update PushItem_Desc set
        PD_Summary=@PD_Summary,
        PD_BackwardDesc=@PD_BackwardDesc,
        PD_ModDate=@PD_ModDate
        where PD_PushitemGuid=@PD_PushitemGuid and PD_Year=@PD_Year and PD_Season=@PD_Season and PD_Stage=@PD_Stage
    end
else
    begin
	    insert into PushItem_Desc (
        PD_Guid,
        PD_PushitemGuid,
        PD_ProjectGuid,
        PD_Stage,
        PD_Year,
        PD_Season,
        PD_Summary,
        PD_BackwardDesc,
        PD_ModDate,
        PD_Status
        ) values (
        @PD_Guid,
        @PD_PushitemGuid,
        @PD_ProjectGuid,
        @PD_Stage,
        @PD_Year,
        @PD_Season,
        @PD_Summary,
        @PD_BackwardDesc,
        @PD_ModDate,
        @PD_Status
        ) 
    end
";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@PD_Guid", PD_Guid);
        oCmd.Parameters.AddWithValue("@PD_PushitemGuid", PD_PushitemGuid);
        oCmd.Parameters.AddWithValue("@PD_ProjectGuid", PD_ProjectGuid);
        oCmd.Parameters.AddWithValue("@PD_Stage", PD_Stage);
        oCmd.Parameters.AddWithValue("@PD_Year", PD_Year);
        oCmd.Parameters.AddWithValue("@PD_Season", PD_Season);
        oCmd.Parameters.AddWithValue("@PD_Summary", PD_Summary);
        oCmd.Parameters.AddWithValue("@PD_BackwardDesc", PD_BackwardDesc);
        oCmd.Parameters.AddWithValue("@PD_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@PD_Status", "A");

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    public void addPushitemDesc()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"insert into PushItem_Desc (
PD_Guid,
PD_PushitemGuid,
PD_ProjectGuid,
PD_Stage,
PD_Year,
PD_Season,
PD_Summary,
PD_BackwardDesc,
PD_ModDate,
PD_Status
) values (
@PD_Guid,
@PD_PushitemGuid,
@PD_ProjectGuid,
@PD_Stage,
@PD_Year,
@PD_Season,
@PD_Summary,
@PD_BackwardDesc,
@PD_ModDate,
@PD_Status
) ";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@PD_Guid", PD_Guid);
        oCmd.Parameters.AddWithValue("@PD_PushitemGuid", PD_PushitemGuid);
        oCmd.Parameters.AddWithValue("@PD_ProjectGuid", PD_ProjectGuid);
        oCmd.Parameters.AddWithValue("@PD_Stage", PD_Stage);
        oCmd.Parameters.AddWithValue("@PD_Year", PD_Year);
        oCmd.Parameters.AddWithValue("@PD_Season", PD_Season);
        oCmd.Parameters.AddWithValue("@PD_Summary", PD_Summary);
        oCmd.Parameters.AddWithValue("@PD_BackwardDesc", PD_BackwardDesc);
        oCmd.Parameters.AddWithValue("@PD_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@PD_Status", "A");

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    public void UpdatePushitemDesc()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"update PushItem_Desc set
PD_Summary=@PD_Summary,
PD_BackwardDesc=@PD_BackwardDesc,
PD_ModDate=@PD_ModDate
where PD_ID=@PD_ID
 ";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@PD_ID", PD_ID);
        oCmd.Parameters.AddWithValue("@PD_Guid", PD_Guid);
        oCmd.Parameters.AddWithValue("@PD_PushitemGuid", PD_PushitemGuid);
        oCmd.Parameters.AddWithValue("@PD_ProjectGuid", PD_ProjectGuid);
        oCmd.Parameters.AddWithValue("@PD_Stage", PD_Stage);
        oCmd.Parameters.AddWithValue("@PD_Year", PD_Year);
        oCmd.Parameters.AddWithValue("@PD_Season", PD_Season);
        oCmd.Parameters.AddWithValue("@PD_Summary", PD_Summary);
        oCmd.Parameters.AddWithValue("@PD_BackwardDesc", PD_BackwardDesc);
        oCmd.Parameters.AddWithValue("@PD_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@PD_Status", "A");

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }
}