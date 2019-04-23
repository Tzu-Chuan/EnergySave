using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;

/// <summary>
/// ExpandFinish_DB 的摘要描述
/// </summary>
public class ExpandFinish_DB
{
    string KeyWord = string.Empty;
    public string _KeyWord
    {
        set { KeyWord = value; }
    }
    #region 私用
    string EF_ID = string.Empty;
    string EF_ReportId = string.Empty;
    string EF_PushitemId = string.Empty;
    string EF_Finish = string.Empty;
    string EF_Status = string.Empty;

    DateTime EF_CreateDate;
    DateTime EF_ModDate;
    #endregion
    #region 公用
    public string _EF_ID
    {
        set { EF_ID = value; }
    }
    public string _EF_ReportId
    {
        set { EF_ReportId = value; }
    }
    public string _EF_PushitemId
    {
        set { EF_PushitemId = value; }
    }
    public string _EF_Finish
    {
        set { EF_Finish = value; }
    }
    public string _EF_Status
    {
        set { EF_Status = value; }
    }
    public DateTime _EF_CreateDate
    {
        set { EF_CreateDate = value; }
    }
    public DateTime _EF_ModDate
    {
        set { EF_ModDate = value; }
    }
    #endregion

    public DataTable GetDataByRSID(string RS_ID)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
declare @RS_Guid nvarchar(50)=(select RS_Guid from ReportSeason where RS_ID=@RS_ID)

select EF_PushitemId,EF_Finish from ExpandFinish 
where EF_ReportId=@RS_Guid and EF_Status='A'
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@RS_ID", RS_ID);
        oda.Fill(ds);
        return ds;
    }

    public void SaveExFinish()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"
declare @rownum int = (select COUNT(*) from ExpandFinish where EF_ReportId=@EF_ReportId and EF_PushitemId=@EF_PushitemId and EF_Status='A')

if(@rownum>0)
    begin
	    update ExpandFinish set 
        EF_Finish=@EF_Finish,
        EF_ModDate=@EF_ModDate
        where EF_ReportId=@EF_ReportId and EF_PushitemId=@EF_PushitemId
    end
else
    begin
	    insert into ExpandFinish (
		    EF_ReportId,
			EF_PushitemId,
			EF_Finish,
			EF_ModDate,
			EF_Status
		    ) values (
		    @EF_ReportId,
			@EF_PushitemId,
			@EF_Finish,
			@EF_ModDate,
			@EF_Status
		    )
    end ";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@EF_ReportId", EF_ReportId);
        oCmd.Parameters.AddWithValue("@EF_PushitemId", EF_PushitemId);
        oCmd.Parameters.AddWithValue("@EF_Finish", EF_Finish);
        oCmd.Parameters.AddWithValue("@EF_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@EF_Status", "A");

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }
}