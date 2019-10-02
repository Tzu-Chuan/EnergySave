using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;

/// <summary>
/// ReportCheck_DB 的摘要描述
/// </summary>
public class ReportCheck_DB
{
    public ReportCheck_DB()
    {
        //
        // TODO: 在這裡新增建構函式邏輯
        //
    }
    #region 全私用
    string M_Guid = string.Empty;
    string strKeyword = string.Empty;
    string strCheckType = string.Empty;
    string strReportType = string.Empty;
    #endregion
    #region 全公用
    public string _M_Guid
    {
        set { M_Guid = value; }
    }
    public string _strKeyword
    {
        set { strKeyword = value; }
    }
    public string _strCheckType
    {
        set { strCheckType = value; }
    }
    public string _strReportType
    {
        set { strReportType = value; }
    }
    #endregion

    #region 私用
    string RC_ID = string.Empty;
    string RC_Guid = string.Empty;
    string RC_ReportType = string.Empty;
    string RC_OldSeasonGuid = string.Empty;
    string RC_ReportGuid = string.Empty;
    string RC_PeopleGuid = string.Empty;
    string RC_Stage = string.Empty;
    string RC_Year = string.Empty;
    string RC_Month = string.Empty;
    string RC_Season = string.Empty;
    DateTime RC_CreateDate;
    string RC_CreateId = string.Empty;
    string RC_Boss = string.Empty;
    string RC_CheckType = string.Empty;
    DateTime RC_CheckDate;
    string RC_Status = string.Empty;
    #endregion
    #region 公用
    public string _RC_ID
    {
        set { RC_ID = value; }
    }
    public string _RC_Guid
    {
        set { RC_Guid = value; }
    }
    public string _RC_ReportType
    {
        set { RC_ReportType = value; }
    }
    public string _RC_OldSeasonGuid
    {
        set { RC_OldSeasonGuid = value; }
    }
    public string _RC_ReportGuid
    {
        set { RC_ReportGuid = value; }
    }
    public string _RC_PeopleGuid
    {
        set { RC_PeopleGuid = value; }
    }
    public string _RC_Stage
    {
        set { RC_Stage = value; }
    }
    public string _RC_Year
    {
        set { RC_Year = value; }
    }
    public string _RC_Month
    {
        set { RC_Month = value; }
    }
    public string _RC_Season
    {
        set { RC_Season = value; }
    }
    public DateTime _RC_CreateDate
    {
        set { RC_CreateDate = value; }
    }
    public string _RC_CreateId
    {
        set { RC_CreateId = value; }
    }
    public string _RC_Boss
    {
        set { RC_Boss = value; }
    }
    public string _RC_CheckType
    {
        set { RC_CheckType = value; }
    }
    public DateTime _RC_CheckDate
    {
        set { RC_CheckDate = value; }
    }
    public string _RC_Status
    {
        set { RC_Status = value; }
    }
    #endregion

    //月報 送審
    public void addMonth()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        
            oCmd.CommandText = @"
                insert into ReportCheck (RC_Guid,RC_ReportType,RC_ReportGuid,RC_PeopleGuid,RC_Stage,RC_Year,RC_Month,RC_Season,RC_Status,RC_CreateId,RC_CheckType)
                values(newid(),@RC_ReportType,@RC_ReportGuid,@RC_PeopleGuid,@RC_Stage,@RC_Year,@RC_Month,@RC_Season,'A',@RC_PeopleGuid,'')
            ";

        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@RC_ReportType", RC_ReportType);
        oCmd.Parameters.AddWithValue("@RC_ReportGuid", RC_ReportGuid);
        oCmd.Parameters.AddWithValue("@RC_PeopleGuid", RC_PeopleGuid);
        oCmd.Parameters.AddWithValue("@RC_Stage", RC_Stage);
        oCmd.Parameters.AddWithValue("@RC_Year", RC_Year);
        oCmd.Parameters.AddWithValue("@RC_Month", RC_Month);
        oCmd.Parameters.AddWithValue("@RC_Season", RC_Season);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    //季報 送審
    public void addSeason()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);

        oCmd.CommandText = @"
                insert into ReportCheck (
RC_Guid,
RC_ReportType,
RC_ReportGuid,
RC_PeopleGuid,
RC_Stage,
RC_Year,
RC_Season,
RC_CreateId,
RC_Status
)
values(
@RC_Guid,
@RC_ReportType,
@RC_ReportGuid,
@RC_PeopleGuid,
@RC_Stage,
@RC_Year,
@RC_Season,
@RC_CreateId,
@RC_Status
) ";

        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@RC_Guid", RC_Guid);
        oCmd.Parameters.AddWithValue("@RC_ReportType", "02");
        oCmd.Parameters.AddWithValue("@RC_ReportGuid", RC_ReportGuid);
        oCmd.Parameters.AddWithValue("@RC_PeopleGuid", RC_PeopleGuid);
        oCmd.Parameters.AddWithValue("@RC_Stage", RC_Stage);
        oCmd.Parameters.AddWithValue("@RC_Year", RC_Year);
        oCmd.Parameters.AddWithValue("@RC_Season", RC_Season);
        oCmd.Parameters.AddWithValue("@RC_CreateId", RC_CreateId);
        oCmd.Parameters.AddWithValue("@RC_Status", "A");

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    //撈 需要審核的月報&季報
    public DataTable selectCheckList()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
            select a.RC_ID,a.RC_Guid,a.RC_ReportType,a.RC_ReportGuid,a.RC_CheckType,a.RC_CreateDate,a.RC_Status,b.RM_ProjectGuid,b.RM_Stage,b.RM_Year,b.RM_Month,c.I_People,e.M_Name,e.M_Office,f.C_Item_cn
            from ReportCheck a
            left join ReportMonth b on a.RC_ReportGuid = b.RM_ReportGuid and b.RM_Status<>'D'
            left join ProjectInfo c on b.RM_ProjectGuid = c.I_Guid and c.I_Status<>'D'
            left join Member d on c.I_People = d.M_Guid and d.M_Manager_ID=@M_Guid  and d.M_Status<>'D'
            left join Member e on c.I_People = e.M_Guid and e.M_Status<>'D'
            left join CodeTable f on e.M_City = f.C_Item and f.C_Group='02'
            where a.RC_Status<>'D' and a.RC_ReportType=@RC_ReportType
        ");
        if (strKeyword!="") {
            sb.Append(@" and (upper(M_Name) LIKE '%' + upper(@strKeyword) + '%' or upper(M_Office) LIKE '%' + upper(@strKeyword) + '%' )  ");
            oCmd.Parameters.AddWithValue("@strKeyword", strKeyword);
        }
        sb.Append(@"
        
            group by a.RC_ID,a.RC_Guid,a.RC_ReportType,a.RC_ReportGuid,a.RC_CheckType,a.RC_CreateDate,a.RC_Status,b.RM_ProjectGuid,b.RM_Stage,b.RM_Year,b.RM_Month,c.I_People,e.M_Name,e.M_Office,f.C_Item_cn
            order by a.RC_CheckType desc,a.RC_CreateDate desc
        ");
        
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();
        oCmd.Parameters.AddWithValue("@M_Guid", M_Guid); 
        oCmd.Parameters.AddWithValue("@RC_ReportType", RC_ReportType); 
        
        oda.Fill(dt);
        return dt;
    }

    public DataSet getReviewMonth(string pStart, string pEnd, string startDay, string endDay, string bossID, string year, string month)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT COUNT(*) total from ReportCheck 
	left join Member on M_Guid=RC_PeopleGuid and M_Status='A'
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
where (RC_ReportType='01' or RC_ReportType='03') and M_Manager_ID=@M_Manager_ID and RC_Status='A' ");

        if (year != "")
            sb.Append(@"and RC_Year=@RC_Year ");
        if (month != "")
            sb.Append(@"and RC_Month=@RC_Month ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (RC_CreateDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and RC_CreateDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and RC_CreateDate<@endDay ");

        if (strKeyword != "")
        {
            sb.Append(@"and ((upper(M_Name) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        if (strReportType != "")
        {
            sb.Append(@" and RC_ReportType=@RC_ReportType ");
        }

        sb.Append(@"select * from (
	select ROW_NUMBER() over (order by RC_CheckType,RC_CreateDate desc,RC_ID desc) itemNo,ReportCheck.*,Member.*,
	city_type.C_Item_cn City
	from ReportCheck
	left join Member on M_Guid=RC_PeopleGuid and M_Status='A'
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
	where (RC_ReportType='01' or RC_ReportType='03') and M_Manager_ID=@M_Manager_ID and RC_Status='A' ");

        if(year!="")
            sb.Append(@"and RC_Year=@RC_Year ");
        if (month != "")
            sb.Append(@"and RC_Month=@RC_Month ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (RC_CreateDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and RC_CreateDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and RC_CreateDate<@endDay ");

        if (strKeyword != "")
        {
            sb.Append(@"and ((upper(M_Name) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        if (strReportType != "")
        {
            sb.Append(@" and RC_ReportType=@RC_ReportType ");
        }

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@KeyWord", strKeyword);
        oCmd.Parameters.AddWithValue("@M_Manager_ID", bossID);
        if (startDay != "")
            oCmd.Parameters.AddWithValue("@startDay", DateTime.Parse(startDay));
        if (endDay != "")
            oCmd.Parameters.AddWithValue("@endDay", DateTime.Parse(endDay).AddDays(1));
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oCmd.Parameters.AddWithValue("@RC_Year", year);
        oCmd.Parameters.AddWithValue("@RC_Month", month);
        oCmd.Parameters.AddWithValue("@RC_ReportType", strReportType);
        oda.Fill(ds);
        return ds;
    }

    public DataSet getReviewSeason(string pStart, string pEnd, string startDay, string endDay, string city, string bossID, string year, string season)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT COUNT(*) total from ReportCheck 
	left join Member mm on mm.M_Guid=RC_PeopleGuid and mm.M_Status='A'
	left join Member ad on ad.M_Guid=RC_Boss and ad.M_Status='A'
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=mm.M_City
where RC_ReportType='02' and RC_Status='A' ");

        if (city != "")
            sb.Append(@"and mm.M_City=@city ");
        if (year != "")
            sb.Append(@"and RC_Year=@RC_Year ");
        if (season != "")
            sb.Append(@"and RC_Season=@RC_Season ");

        if (startDay != "")
            sb.Append(@"and RC_CheckDate>=@startDay ");
        if (endDay != "")
            sb.Append(@"and RC_CheckDate<=@endDay ");

        if (strKeyword != "")
        {
            sb.Append(@"and ((upper(mm.M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(ad.M_Name) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@"select * from (
	select ROW_NUMBER() over (order by RC_CheckType,RC_CreateDate desc,RC_ID desc) itemNo,
	ReportCheck.*,
	mm.M_City,
	mm.M_Name as MbName,
	ad.M_Name as AdName,
	city_type.C_Item_cn as City,
	(select RS_ID from ReportSeason where RC_ReportGuid=RS_Guid) as RS_ID
	from ReportCheck
	left join Member mm on mm.M_Guid=RC_PeopleGuid and mm.M_Status='A'
	left join Member ad on ad.M_Guid=RC_Boss and ad.M_Status='A'
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=mm.M_City
	where RC_ReportType='02' and RC_Status='A' ");

        if (city != "")
            sb.Append(@"and mm.M_City=@city ");
        if (year != "")
            sb.Append(@"and RC_Year=@RC_Year ");
        if (season != "")
            sb.Append(@"and RC_Season=@RC_Season ");

        if (startDay != "")
            sb.Append(@"and RC_CheckDate>=@startDay ");
        if (endDay != "")
            sb.Append(@"and RC_CheckDate<=@endDay ");

        if (strKeyword != "")
        {
            sb.Append(@"and ((upper(mm.M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(ad.M_Name) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@KeyWord", strKeyword);
        if (startDay != "")
            oCmd.Parameters.AddWithValue("@startDay", DateTime.Parse(startDay));
        if (endDay != "")
            oCmd.Parameters.AddWithValue("@endDay", DateTime.Parse(endDay).AddDays(1));
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oCmd.Parameters.AddWithValue("@city", city);
        oCmd.Parameters.AddWithValue("@RC_Year", year);
        oCmd.Parameters.AddWithValue("@RC_Season", season);
        oda.Fill(ds);
        return ds;
    }

    //月報 季報 主管審核
    public void ReportCheck()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);

        if (strCheckType=="Y") {
            //通過
            oCmd.CommandText = @"
                update ReportCheck 
                set RC_CheckType=@RC_CheckType,RC_CheckDate=@RC_CheckDate,RC_Boss=@RC_Boss
                where RC_ReportGuid=@RC_ReportGuid and RC_Status<>'D'
            ";
            oCmd.Parameters.AddWithValue("@RC_CheckType", strCheckType);
        }
        if (strCheckType == "N")
        {
            //不通過
            oCmd.CommandText = @"
                update ReportCheck 
                set RC_CheckDate=@RC_CheckDate,RC_Status='D',RC_Boss=@RC_Boss
                where RC_ReportGuid=@RC_ReportGuid and RC_Status<>'D'
            ";
        }
        

        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@RC_ReportGuid", RC_ReportGuid);
        oCmd.Parameters.AddWithValue("@RC_CheckDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@RC_Boss", RC_Boss);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    public DataSet getHistoryMonth(string pStart, string pEnd, string startDay, string endDay, string city, string year, string month, string reporttype)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT COUNT(*) total from ReportCheck 
	left join Member mm on mm.M_Guid=RC_PeopleGuid --and mm.M_Status='A'
	left join Member ad on ad.M_Guid=RC_Boss --and ad.M_Status='A'
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=mm.M_City
where  RC_Status='A' and RC_CheckType='Y' and (RC_ReportType='01' or RC_ReportType='03') ");

        if (city != "")
            sb.Append(@"and mm.M_City=@city ");
        if (year != "")
            sb.Append(@"and RC_Year=@RC_Year ");
        if (month != "")
            sb.Append(@"and RC_Month=@RC_Month ");
        if (reporttype != "")
            sb.Append(@"and RC_ReportType=@RC_ReportType ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (RC_CheckDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and RC_CheckDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and RC_CheckDate<@endDay ");

        if (strKeyword != "")
        {
            sb.Append(@"and ((upper(mm.M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(ad.M_Name) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@"select * from (
	select ROW_NUMBER() over (order by RC_CheckType,RC_CreateDate desc,RC_ID desc) itemNo,ReportCheck.*
	,mm.M_City,mm.M_Name MbName,ad.M_Name AdName,
	city_type.C_Item_cn City
	from ReportCheck
	left join Member mm on mm.M_Guid=RC_PeopleGuid --and mm.M_Status='A'
	left join Member ad on ad.M_Guid=RC_Boss --and ad.M_Status='A'
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=mm.M_City
	where  RC_Status='A' and RC_CheckType='Y' and (RC_ReportType='01' or RC_ReportType='03') ");

        if(city!="")
            sb.Append(@"and mm.M_City=@city ");
        if (year != "")
            sb.Append(@"and RC_Year=@RC_Year ");
        if (month != "")
            sb.Append(@"and RC_Month=@RC_Month ");
        if (reporttype != "")
            sb.Append(@"and RC_ReportType=@RC_ReportType ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (RC_CheckDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and RC_CheckDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and RC_CheckDate<@endDay ");

        if (strKeyword != "")
        {
            sb.Append(@"and ((upper(mm.M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(ad.M_Name) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@KeyWord", strKeyword);
        oCmd.Parameters.AddWithValue("@city", city);
        if (startDay != "")
            oCmd.Parameters.AddWithValue("@startDay", DateTime.Parse(startDay));
        if (endDay != "")
            oCmd.Parameters.AddWithValue("@endDay", DateTime.Parse(endDay).AddDays(1));
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oCmd.Parameters.AddWithValue("@RC_Year", year);
        oCmd.Parameters.AddWithValue("@RC_Month", month);
        oCmd.Parameters.AddWithValue("@RC_ReportType", reporttype);
        oda.Fill(ds);
        return ds;
    }

    public DataSet getHistorySeason(string pStart, string pEnd, string startDay, string endDay, string city, string year, string season)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT COUNT(*) total from ReportCheck 
	left join Member mm on mm.M_Guid=RC_PeopleGuid --and mm.M_Status='A'
	left join Member ad on ad.M_Guid=RC_Boss --and ad.M_Status='A'
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=mm.M_City
where RC_ReportType='02' and RC_Status='A' and RC_CheckType='Y' ");

        if (city != "")
            sb.Append(@"and mm.M_City=@city ");
        if (year != "")
            sb.Append(@"and RC_Year=@RC_Year ");
        if (season != "")
            sb.Append(@"and RC_Season=@RC_Season ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (RC_CheckDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and RC_CheckDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and RC_CheckDate<@endDay ");

        if (strKeyword != "")
        {
            sb.Append(@"and ((upper(mm.M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(ad.M_Name) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@"select * from (
	select ROW_NUMBER() over (order by RC_CheckType,RC_CreateDate desc,RC_ID desc) itemNo,
	ReportCheck.*,
	mm.M_City,
	mm.M_Name as MbName,
	ad.M_Name as AdName,
	city_type.C_Item_cn as City,
	(select RS_ID from ReportSeason where RC_ReportGuid=RS_Guid) as RS_ID
	from ReportCheck
	left join Member mm on mm.M_Guid=RC_PeopleGuid
	left join Member ad on ad.M_Guid=RC_Boss
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=mm.M_City
	where RC_ReportType='02' and RC_Status='A' and RC_CheckType='Y' ");

        if (city != "")
            sb.Append(@"and mm.M_City=@city ");
        if (year != "")
            sb.Append(@"and RC_Year=@RC_Year ");
        if (season != "")
            sb.Append(@"and RC_Season=@RC_Season ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (RC_CheckDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and RC_CheckDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and RC_CheckDate<@endDay ");

        if (strKeyword != "")
        {
            sb.Append(@"and ((upper(mm.M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(ad.M_Name) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@KeyWord", strKeyword);
        oCmd.Parameters.AddWithValue("@city", city);
        if (startDay != "")
            oCmd.Parameters.AddWithValue("@startDay", DateTime.Parse(startDay));
        if (endDay != "")
            oCmd.Parameters.AddWithValue("@endDay", DateTime.Parse(endDay).AddDays(1));
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oCmd.Parameters.AddWithValue("@RC_Year", year);
        oCmd.Parameters.AddWithValue("@RC_Season", season);
        oda.Fill(ds);
        return ds;
    }

    public DataTable CheckReportExist()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();
        
        sb.Append(@"select Count(*) as Total from ReportCheck
where RC_ReportGuid=@RC_ReportGuid 
and RC_Stage=@RC_Stage 
and RC_Year=@RC_Year 
and RC_Season=@RC_Season
and RC_Status='A' ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@RC_ReportGuid", RC_ReportGuid);
        oCmd.Parameters.AddWithValue("@RC_Stage", RC_Stage);
        oCmd.Parameters.AddWithValue("@RC_Year", RC_Year);
        oCmd.Parameters.AddWithValue("@RC_Season", RC_Season);
        oda.Fill(ds);
        return ds;
    }

    //20190801新增撈季報歷史資料列表所有資料
    public DataTable getHistoryMonthList()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select ROW_NUMBER() over (order by RC_CheckType,RC_CreateDate desc,RC_ID desc) itemNo,
city_type.C_Item_cn City,
case ReportCheck.RC_ReportType when '01' then '設備汰換' when '03' then '擴大補助' else '' end as ReportType,
ReportCheck.RC_Year,
ReportCheck.RC_Month,
mm.M_Name MbName,
ReportCheck.RC_CheckDate,
ad.M_Name AdName
--ReportCheck.*,
--mm.M_City,
from ReportCheck
left join Member mm on mm.M_Guid=RC_PeopleGuid
left join Member ad on ad.M_Guid=RC_Boss
left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=mm.M_City
where  RC_Status='A' and RC_CheckType='Y' and (RC_ReportType='01' or RC_ReportType='03') and RC_Stage=@RC_Stage
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;

        oCmd.Parameters.AddWithValue("@RC_Stage", RC_Stage);

        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oda.Fill(ds);
        return ds;
    }
    //20190801新增撈季報歷史資料列表所有資料
    public DataTable getHistorySeasonList()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select ROW_NUMBER() over (order by RC_CheckType,RC_CreateDate desc,RC_ID desc) itemNo,
city_type.C_Item_cn as City,
ReportCheck.RC_Year,
ReportCheck.RC_Season,
mm.M_Name as MbName,
ReportCheck.RC_CheckDate,
ad.M_Name as AdName
--ReportCheck.*,
--mm.M_City,
--(select RS_ID from ReportSeason where RC_ReportGuid=RS_Guid) as RS_ID
from ReportCheck
left join Member mm on mm.M_Guid=RC_PeopleGuid
left join Member ad on ad.M_Guid=RC_Boss
left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=mm.M_City
where RC_ReportType='02' and RC_Status='A' and RC_CheckType='Y'  and RC_Stage=@RC_Stage
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        oCmd.Parameters.AddWithValue("@RC_Stage", RC_Stage);

        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oda.Fill(ds);
        return ds;
    }
}