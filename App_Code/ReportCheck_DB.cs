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
    string strCity = string.Empty;
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
    public string _strCity
    {
        set { strCity = value; }
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

    public DataSet getHistoryMonth(string pStart, string pEnd, string startDay, string endDay, string city, string year, string month, string reporttype, string stage)
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
        if (stage != "")
            sb.Append(@"and RC_Stage=@RC_Stage ");

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
        if (stage != "")
            sb.Append(@"and RC_Stage=@RC_Stage ");

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
        oCmd.Parameters.AddWithValue("@RC_Stage", stage);
        oda.Fill(ds);
        return ds;
    }

    public DataSet getHistorySeason(string pStart, string pEnd, string startDay, string endDay, string city, string year, string season,string stage)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT RC_db.*,	
	mm.M_City,
	mm.M_Name as MbName,
	ad.M_Name as AdName,
	city_type.C_Item_cn as City,
	(select RS_ID from ReportSeason where RC_ReportGuid=RS_Guid) as RS_ID
	into #tmpAll 
	from ReportCheck RC_db
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
        if (stage != "")
            sb.Append(@"and RC_Stage=@RC_Stage ");

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

        sb.Append(@"
--總筆數
select count(*) as total from #tmpAll
--分頁資料
select * from (
select ROW_NUMBER() over (order by RC_CheckType,RC_CreateDate desc,RC_ID desc) itemNo,#tmpAll.*
from #tmpAll
)#tmp where itemNo between @pStart and @pEnd

drop table #tmpAll ");

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
        oCmd.Parameters.AddWithValue("@RC_Stage", stage);
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

    //20190801新增撈月報歷史資料列表所有資料
    public DataTable getHistoryMonthList()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select ROW_NUMBER() over (order by mm.M_City asc,RC_Stage asc,RC_Year asc,RC_Month asc) itemNo,
city_type.C_Item_cn City,
case ReportCheck.RC_ReportType when '01' then '設備汰換' when '03' then '擴大補助' else '' end as ReportType,
ReportCheck.RC_Stage,
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
        if (strCity != "")
        {
            sb.Append(@" and mm.M_City=@M_City ");
            oCmd.Parameters.AddWithValue("@M_City", strCity);
        }

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

    //202001新增 設備汰換月報送審後更新當期底下該月份以後的所有月報累計值
    public void updateReportMonthNum()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);

        
        oCmd.CommandText = @"
/*
功能：設備汰換閱報送審後，更新該期該月份以後所有月報的完成數
	  因為就算審核通過了，還是可以抽單重填，但原系統設定每份月報送審後就不能再修改，
	  每份月報都要上一份月報審核通過後才能繼續填下一份月報，且還有人一次填寫多月的月報未送審，
	  現在這樣搞的話月報的累計數字完全不對
	  所以才要在送審後去更新後面月報的數字

傳入參數：RC_ReportGuid
--68aa9ea46b8a4494b19e6858b62dc043--9  9cbfe62eb3da42b3922cb66da822d444--8 
--1438b091a01441e6b7689ecefe9afe58--4  d7d12ec809c947aa9406e0fa15b33386--5
--0feb9f52b0a44e4ba86d41a439fbef3e--6 c634e417c1034200adc714a5cdd9dbec--7
--select * from ReportMonth where RM_ProjectGuid='68aa9ea46b8a4494b19e6858b62dc043' and RM_Year='2019' and RM_Status='A' and RM_ReportType='01'
作者：王晨峻
*/

declare @RC_ReportGuid nvarchar(50)=@RC_ReportGuid
declare @RM_ProjectGuid nvarchar(50)=''--縣市計畫代號
declare @RM_ReportType nvarchar(50)=''--月報類別
declare @RM_Stage nvarchar(50)=''--期
declare @RM_Year nvarchar(50)=''--年
declare @RM_Month nvarchar(50)=''--月
--根據帶過來的RC_ReportGuid 找到ReportMonth裡面該筆月報資料
select top 1  @RM_ProjectGuid=RM_ProjectGuid,@RM_ReportType=RM_ReportType, @RM_Stage=RM_Stage,@RM_Year=RM_Year,@RM_Month=RM_Month
from ReportMonth 
where RM_ReportGuid=@RC_ReportGuid and RM_Status='A'
--select @RM_ProjectGuid
--select @RM_ReportType,@RM_Stage,@RM_Year,@RM_Month

--撈出該份月報以後所有已經審核通過的月報
select RM_ID,RM_ReportGuid,RM_ProjectGuid,RM_CPType,RM_Stage,RM_Year,RM_Month,RM_ReportType,RM_Status,RC_CheckType,RC_Status
into #tmpRM
from ReportMonth
left join ReportCheck on RM_ReportGuid=RC_ReportGuid and RC_CheckType='Y' and RC_Status='A'  and RC_Stage=@RM_Stage and RC_Month=RM_Month and RC_Year=RM_Year
where RM_ProjectGuid=@RM_ProjectGuid and RM_Stage=@RM_Stage and RM_ReportType=@RM_ReportType and RM_Status='A' and RC_CheckType='Y' and RC_Stage=@RM_Stage and RC_Status='A' 
and CONVERT(int,(isnull(RM_Year,'0')+isnull(RM_Month,'0')))>CONVERT(int,(isnull(@RM_Year,'0')+isnull(@RM_Month,'0')))
order by RM_Year,RM_Month

--group by出有幾分月報要更新 進#tmpwhile 跑while
select RM_ReportGuid,RM_ProjectGuid,RM_Stage,RM_Year,RM_Month,RM_ReportType 
into #tmpwhile
from #tmpRM
group by RM_ReportGuid,RM_ProjectGuid,RM_Stage,RM_Year,RM_Month,RM_ReportType

--while 迴圈
declare @introwC int;
declare @whileRM_ReportGuid nvarchar(50);
declare @whileRM_ProjectGuid nvarchar(50);
declare @whileRM_Stage nvarchar(50);
declare @whileRM_Year nvarchar(50);
declare @whileRM_Month nvarchar(50);
declare @whileRM_ReportType nvarchar(50);
declare @numRM_Finish01 decimal(10,1)
declare @numRM_Finish01KW decimal(10,1)
declare @numRM_Finish02 decimal(10,1)
declare @numRM_Finish03 decimal(10,1)
declare @numRM_Finish04 decimal(10,1)
declare @numRM_Finish05 decimal(10,1)
create table #tmpWhileNum(
	RM_CPType nvarchar(50)
	,countFinishKW decimal(10,1)
	,countFinish03 decimal(10,1)
	,countFinish02 decimal(10,1)
	--,countApplyKW decimal(10,1)
	--,countApply01 decimal(10,1)
)
select @introwC = COUNT(*) from #tmpwhile

while @introwC > 0
begin
	select top 1 @whileRM_ReportGuid=RM_ReportGuid,@whileRM_ProjectGuid=RM_ProjectGuid,@whileRM_Stage=RM_Stage
			,@whileRM_Year=RM_Year,@whileRM_Month=RM_Month,@whileRM_ReportType=RM_ReportType
	from #tmpwhile order by RM_Stage asc,RM_Year asc,RM_Month asc

	--select @whileRM_Year,@whileRM_Month
	--找出這一期小於這個月月報的所有資料出來SUM
	insert into #tmpWhileNum(RM_CPType,countFinishKW,countFinish03,countFinish02)
	select RM_CPType--RM_Stage,RM_Year,RM_Month,RM_ReportGuid,RM_ProjectGuid,RM_ReportType,RC_Status,RC_CheckType
	,SUM(isnull(RM_Type4ValueSum,0.0)) as countFinishKW --累計完成數 KW(無風管)
	,SUM(isnull(RM_Type3ValueSum,0.0)) as countFinish03--累計完成數(老舊 停車場 中型 大型)
	,SUM(isnull(RM_Type2ValueSum,0.0)) as countFinish02--累計核定數(老舊 停車場 中型 大型)
	--,SUM(isnull(RM_Type3ValueSum,0.0)) as countApplyKW--累計申請數 KW(無風管)
	--,SUM(isnull(RM_Type1ValueSum,0.0)) as countApply01--累計完成數(老舊 停車場 中型 大型)
	from ReportMonth left join ReportCheck on RM_ReportGuid=RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
	where RM_ProjectGuid=@RM_ProjectGuid
	and RM_Stage=@RM_Stage and RM_Status='A' and RM_ReportType=@RM_ReportType and RC_Status='A' and RC_CheckType='Y'
	and CONVERT(int,(isnull(RM_Year,'0')+isnull(RM_Month,'0')))<CONVERT(int,(isnull(@whileRM_Year,'0')+isnull(@whileRM_Month,'0')))
	group by RM_CPType--RM_Stage,RM_Year,RM_Month,RM_ReportGuid,RM_ProjectGuid,RM_ReportType,RC_Status,RC_CheckType

	select @numRM_Finish01 = countFinish02 from #tmpWhileNum where RM_CPType='01' --update RM_Finish01
	select @numRM_Finish01KW = countFinishKW from #tmpWhileNum where RM_CPType='01'--update RM_Finish
	select @numRM_Finish02 = countFinish02 from #tmpWhileNum where RM_CPType='02'--update RM_Finish
	select @numRM_Finish03 = countFinish02 from #tmpWhileNum where RM_CPType='03'--update RM_Finish
	select @numRM_Finish04 = countFinish03 from #tmpWhileNum where RM_CPType='04'--update RM_Finish
	select @numRM_Finish05 = countFinish03 from #tmpWhileNum where RM_CPType='05'--update RM_Finish
	
	--begin tran
	----驗證用
	----算好新的要update的數字
	--select * from #tmpWhileNum
	--select RM_ID,RM_Stage,RM_Year,RM_Month,RM_CPType,RM_Finish
	--,RM_Type1ValueSum ,RM_Type2ValueSum,RM_Type3ValueSum,RM_Type4ValueSum,RM_Finish01 
	--from ReportMonth where RM_ReportGuid=@whileRM_ReportGuid and RM_ProjectGuid=@whileRM_ProjectGuid and RM_Stage=@whileRM_Stage
	--and RM_Year=@whileRM_Year and RM_Month=@whileRM_Month and RM_ReportType=@whileRM_ReportType and RM_Status='A'
	----驗證用 END

	--update 無風管
	update ReportMonth set RM_Finish=@numRM_Finish01KW,RM_Finish01=@numRM_Finish01
	where RM_ReportGuid=@whileRM_ReportGuid and RM_ProjectGuid=@whileRM_ProjectGuid and RM_Stage=@whileRM_Stage
	and RM_Year=@whileRM_Year and RM_Month=@whileRM_Month and RM_ReportType=@whileRM_ReportType and RM_Status='A' and RM_CPType='01'
	--update 老舊
	update ReportMonth set RM_Finish=@numRM_Finish02
	where RM_ReportGuid=@whileRM_ReportGuid and RM_ProjectGuid=@whileRM_ProjectGuid and RM_Stage=@whileRM_Stage
	and RM_Year=@whileRM_Year and RM_Month=@whileRM_Month and RM_ReportType=@whileRM_ReportType and RM_Status='A' and RM_CPType='02'
	--update 空調
	update ReportMonth set RM_Finish=@numRM_Finish03
	where RM_ReportGuid=@whileRM_ReportGuid and RM_ProjectGuid=@whileRM_ProjectGuid and RM_Stage=@whileRM_Stage
	and RM_Year=@whileRM_Year and RM_Month=@whileRM_Month and RM_ReportType=@whileRM_ReportType and RM_Status='A' and RM_CPType='03'
	--update 中型
	update ReportMonth set RM_Finish=@numRM_Finish04
	where RM_ReportGuid=@whileRM_ReportGuid and RM_ProjectGuid=@whileRM_ProjectGuid and RM_Stage=@whileRM_Stage
	and RM_Year=@whileRM_Year and RM_Month=@whileRM_Month and RM_ReportType=@whileRM_ReportType and RM_Status='A' and RM_CPType='04'
	--update 大型
	update ReportMonth set RM_Finish=@numRM_Finish05
	where RM_ReportGuid=@whileRM_ReportGuid and RM_ProjectGuid=@whileRM_ProjectGuid and RM_Stage=@whileRM_Stage
	and RM_Year=@whileRM_Year and RM_Month=@whileRM_Month and RM_ReportType=@whileRM_ReportType and RM_Status='A' and RM_CPType='05'

	----驗證用
	--select RM_ID,RM_Stage,RM_Year,RM_Month,RM_CPType,RM_Finish
	--,RM_Type1ValueSum ,RM_Type2ValueSum,RM_Type3ValueSum,RM_Type4ValueSum,RM_Finish01 
	--from ReportMonth where RM_ReportGuid=@whileRM_ReportGuid and RM_ProjectGuid=@whileRM_ProjectGuid and RM_Stage=@whileRM_Stage
	--and RM_Year=@whileRM_Year and RM_Month=@whileRM_Month and RM_ReportType=@whileRM_ReportType and RM_Status='A'
	----驗證用 END
	--rollback tran

	--刪除跑迴圈的table
	delete from #tmpwhile
	where  RM_ReportGuid=@whileRM_ReportGuid and RM_ProjectGuid=@whileRM_ProjectGuid and RM_Stage=@whileRM_Stage
			and RM_Year=@whileRM_Year and RM_Month=@whileRM_Month and RM_ReportType=@whileRM_ReportType
	--迴圈
	set @introwC = @introwC-1
	delete from #tmpWhileNum
end

drop table #tmpRM
drop table #tmpwhile
drop table #tmpWhileNum
        ";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@RC_ReportGuid", RC_ReportGuid);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }
}