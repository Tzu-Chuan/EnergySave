using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;

/// <summary>
/// OtherManage_DB 的摘要描述
/// </summary>
public class OtherManage_DB
{
    public DataSet MonthList(string pStart, string pEnd,string city)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
    select count(*) total from ReportCheck 
    left join ProjectInfo on I_Guid=(select top 1 RM_ProjectGuid from ReportMonth where RC_ReportGuid=RM_ReportGuid) 
    where RC_Status='A' and RC_ReportType='01' ");

        if (city != "")
            sb.Append(@"and I_City=@I_City");

        sb.Append(@"
        select * from (
            select ROW_NUMBER() over (order by RC_CreateDate desc) itemNo,
            City.C_Item_cn as City,
            I_Office,
            M_Name,
            RC_Guid,
            RC_ReportGuid,
	        RC_Stage,
	        RC_Year,
	        RC_Month,
            RC_CreateDate
            from ReportCheck
            left join ProjectInfo on I_Guid=(select top 1 RM_ProjectGuid from ReportMonth where RC_ReportGuid=RM_ReportGuid)
            left join Member on RC_PeopleGuid=M_Guid
            left Join CodeTable as City on I_City=City.C_Item and City.C_Group='02'
            where RC_Status='A' and RC_ReportType='01' 
 ");

        if (city != "")
            sb.Append(@"and I_City=@I_City");

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oCmd.Parameters.AddWithValue("@I_City", city);
        oda.Fill(ds);
        return ds;
    }

    public DataSet SeasonList(string pStart, string pEnd, string city)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
    select count(*) total from ReportCheck 
    left join ProjectInfo on I_Guid=(select top 1 RS_PorjectGuid from ReportSeason where RC_ReportGuid=RS_Guid)
    where RC_Status='A' and RC_ReportType='02' ");

        if (city != "")
            sb.Append(@"and I_City=@I_City");

        sb.Append(@"
        select * from (
            select ROW_NUMBER() over (order by RC_CreateDate desc) itemNo,
            City.C_Item_cn as City,
            I_Office,
            M_Name,
            RC_Guid,
            RC_ReportGuid,
	        RC_Stage,
	        RC_Year,
	        RC_Season,
            RC_CreateDate
            from ReportCheck
            left join ProjectInfo on I_Guid=(select top 1 RS_PorjectGuid from ReportSeason where RC_ReportGuid=RS_Guid)
            left join Member on RC_PeopleGuid=M_Guid
            left Join CodeTable as City on I_City=City.C_Item and City.C_Group='02'
            where RC_Status='A' and RC_ReportType='02' 
 ");

        if (city != "")
            sb.Append(@"and I_City=@I_City");

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oCmd.Parameters.AddWithValue("@I_City", city);
        oda.Fill(ds);
        return ds;
    }

    public void CancelMonthReport(string id)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"update ReportCheck set RC_Status='D' where RC_Guid=@RC_Guid ";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@RC_Guid", id);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }
    public void CancelSeasonReport(string id)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"
declare  @RC_Guid nvarchar(50)=@rcGuid
declare  @ProID nvarchar(50)=(select I_Guid from ProjectInfo where I_People=(select RC_PeopleGuid from ReportCheck where RC_Guid=@RC_Guid))
declare @rcYear nvarchar(5)=(select convert(int,RC_Year)-1911 from ReportCheck where RC_Guid=@RC_Guid)
declare @rcSeason nvarchar(5)=(select RC_Season from ReportCheck where RC_Guid=@RC_Guid)
declare @rcStage nvarchar(5)=(select RC_Stage from ReportCheck where RC_Guid=@RC_Guid)
declare @rsID nvarchar(50)=(select RS_Guid from ReportSeason where RS_PorjectGuid=@ProID and RS_Year=@rcYear and RS_Season=@rcSeason and RS_Stage=@rcStage)

update ReportCheck set RC_Status='D' where RC_Guid=@RC_Guid
update PushItem_Desc set PD_Status='D' where PD_ProjectGuid=@ProID and PD_Year=@rcYear and PD_Season=@rcSeason and PD_Stage=@rcStage 
update ExpandFinish set EF_Status='D' where EF_ReportId=@rsID ";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);

        oCmd.Parameters.AddWithValue("@rcGuid", id);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }
}