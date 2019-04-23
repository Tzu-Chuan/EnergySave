using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;

/// <summary>
/// Log_DB 的摘要描述
/// </summary>
public class Log_DB
{
    string KeyWord = string.Empty;
    public string _KeyWord
    {
        set { KeyWord = value; }
    }
    #region 私用
    string L_ID = string.Empty;
    string L_Type = string.Empty;
    string L_Person = string.Empty;
    string L_IP = string.Empty;
    string L_Desc = string.Empty;
    string L_ModItemGuid = string.Empty;

    DateTime L_ModDate;
    #endregion
    #region 公用
    public string _L_ID
    {
        set { L_ID = value; }
    }
    public string _L_Type
    {
        set { L_Type = value; }
    }
    public string _L_Person
    {
        set { L_Person = value; }
    }
    public string _L_IP
    {
        set { L_IP = value; }
    }
    public string _L_Desc
    {
        set { L_Desc = value; }
    }
    public string _L_ModItemGuid
    {
        set { L_ModItemGuid = value; }
    }
    public DateTime _L_ModDate
    {
        set { L_ModDate = value; }
    }
    #endregion

    public DataSet getLogIOList(string pStart, string pEnd,string startDay,string endDay)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT COUNT(*) total from LogTable 
left join Member on M_Guid=L_Person and M_Status='A'
left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
left join CodeTable log_type on log_type.C_Group='06' and log_type.C_Item=L_Type   
where 1=1 ");

        if (L_Type != "")
            sb.Append(@"and L_Type=@L_Type ");
        else
            sb.Append(@"and (L_Type='01' or L_Type='02') ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (L_ModDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and L_ModDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and L_ModDate<@endDay ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_IP) LIKE '%' + upper(@KeyWord) + '%') or 
(upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(comp_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@"select * from (
	select ROW_NUMBER() over (order by L_ModDate desc,L_ID desc) itemNo,log_type.C_Item_cn LType,LogTable.*,Member.*,
	city_type.C_Item_cn City,comp_type.C_Item_cn Comp
	from LogTable
	left join Member on M_Guid=L_Person and M_Status='A'
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
	left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
	left join CodeTable log_type on log_type.C_Group='06' and log_type.C_Item=L_Type
    where 1=1 ");

        if (L_Type != "")
            sb.Append(@"and L_Type=@L_Type ");
        else
            sb.Append(@"and (L_Type='01' or L_Type='02') ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (L_ModDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and L_ModDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and L_ModDate<@endDay ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_IP) LIKE '%' + upper(@KeyWord) + '%') or 
(upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(comp_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@KeyWord", KeyWord);
        oCmd.Parameters.AddWithValue("@L_Type", L_Type);
        if (startDay!="")
            oCmd.Parameters.AddWithValue("@startDay", DateTime.Parse(startDay));
        if (endDay != "")
            oCmd.Parameters.AddWithValue("@endDay", DateTime.Parse(endDay).AddDays(1));
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oda.Fill(ds);
        return ds;
    }

    public DataSet getLogModPwList(string pStart, string pEnd, string startDay, string endDay)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT COUNT(*) total from LogTable 
	left join Member p on p.M_Guid=L_Person and p.M_Status='A'
	left join Member m on m.M_Guid=L_ModItemGuid and m.M_Status='A'
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=m.M_City
where L_Type='03' ");
        
        if (startDay != "" && endDay != "")
            sb.Append(@"and (L_ModDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and L_ModDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and L_ModDate<@endDay ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(L_Person) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_IP) LIKE '%' + upper(@KeyWord) + '%') or 
(upper(L_Desc) LIKE '%' + upper(@KeyWord) + '%') or (upper(m.M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(p.M_Name) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@"select * from (
	select ROW_NUMBER() over (order by L_ModDate desc,L_ID desc) itemNo,LogTable.*,p.M_Name LoginPerson,m.M_Name ModPerson,
	city_type.C_Item_cn City
	from LogTable
	left join Member p on p.M_Guid=L_Person and p.M_Status='A'
	left join Member m on m.M_Guid=L_ModItemGuid and m.M_Status='A'
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=m.M_City
    where L_Type='03' ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (L_ModDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and L_ModDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and L_ModDate<@endDay ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(L_Person) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_IP) LIKE '%' + upper(@KeyWord) + '%') or 
(upper(L_Desc) LIKE '%' + upper(@KeyWord) + '%') or (upper(m.M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(p.M_Name) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@KeyWord", KeyWord);
        oCmd.Parameters.AddWithValue("@L_Type", L_Type);
        if (startDay != "")
            oCmd.Parameters.AddWithValue("@startDay", DateTime.Parse(startDay));
        if (endDay != "")
            oCmd.Parameters.AddWithValue("@endDay", DateTime.Parse(endDay).AddDays(1));
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oda.Fill(ds);
        return ds;
    }

    public DataSet getReportSubList(string pStart, string pEnd, string startDay, string endDay)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT COUNT(*) total from LogTable 
left join Member on M_Guid=L_Person
left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
left join CodeTable log_type on log_type.C_Group='06' and log_type.C_Item=L_Type   
where 1=1 ");

        if (L_Type != "")
            sb.Append(@"and L_Type=@L_Type ");
        else
            sb.Append(@"and (L_Type='04' or L_Type='05') ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (L_ModDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and L_ModDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and L_ModDate<@endDay ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_IP) LIKE '%' + upper(@KeyWord) + '%') or 
(upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(comp_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Office) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@"select * from (
	select ROW_NUMBER() over (order by L_ModDate desc,L_ID desc) itemNo,log_type.C_Item_cn LType,LogTable.*,Member.*,
	city_type.C_Item_cn City,comp_type.C_Item_cn Comp,
	(select top 1 RM_Year from ReportMonth where RM_ReportGuid=L_ModItemGuid and RM_Status='A') MYear,
	(select top 1 RM_Month from ReportMonth where RM_ReportGuid=L_ModItemGuid and RM_Status='A') MMonth,
	(select RS_Year from ReportSeason where RS_Guid=L_ModItemGuid and RS_Status='A') SYear,
	(select RS_Season from ReportSeason where RS_Guid=L_ModItemGuid and RS_Status='A') SSeason
	from LogTable
	left join Member on M_Guid=L_Person
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
	left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
	left join CodeTable log_type on log_type.C_Group='06' and log_type.C_Item=L_Type
    where 1=1 ");

        if (L_Type != "")
            sb.Append(@"and L_Type=@L_Type ");
        else
            sb.Append(@"and (L_Type='04' or L_Type='05') ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (L_ModDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and L_ModDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and L_ModDate<@endDay ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_IP) LIKE '%' + upper(@KeyWord) + '%') or 
(upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(comp_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Office) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@KeyWord", KeyWord);
        oCmd.Parameters.AddWithValue("@L_Type", L_Type);
        if (startDay != "")
            oCmd.Parameters.AddWithValue("@startDay", DateTime.Parse(startDay));
        if (endDay != "")
            oCmd.Parameters.AddWithValue("@endDay", DateTime.Parse(endDay).AddDays(1));
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oda.Fill(ds);
        return ds;
    }

    public DataSet getReviewDateList(string pStart, string pEnd, string startDay, string endDay)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT COUNT(*) total from LogTable 
left join Member on M_Guid=L_Person and M_Status='A'
left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
left join CodeTable log_type on log_type.C_Group='06' and log_type.C_Item=L_Type   
where 1=1 ");

        if (L_Type != "")
            sb.Append(@"and L_Type=@L_Type ");
        else
            sb.Append(@"and (L_Type='06' or L_Type='10') ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (L_ModDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and L_ModDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and L_ModDate<@endDay ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_IP) LIKE '%' + upper(@KeyWord) + '%') or 
(upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(comp_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Office) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@"select * from (
	select ROW_NUMBER() over (order by L_ModDate desc,L_ID desc) itemNo,log_type.C_Item_cn LType,LogTable.*,Member.*,
	city_type.C_Item_cn City,comp_type.C_Item_cn Comp,
	(select top 1 RM_Year from ReportMonth where RM_ReportGuid=L_ModItemGuid and RM_Status='A') MYear,
	(select top 1 RM_Month from ReportMonth where RM_ReportGuid=L_ModItemGuid and RM_Status='A') MMonth,
	(select RS_Year from ReportSeason where RS_Guid=L_ModItemGuid and RS_Status='A') SYear,
	(select RS_Season from ReportSeason where RS_Guid=L_ModItemGuid and RS_Status='A') SSeason
	from LogTable
	left join Member on M_Guid=L_Person and M_Status='A'
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
	left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
	left join CodeTable log_type on log_type.C_Group='06' and log_type.C_Item=L_Type
    where 1=1 ");

        if (L_Type != "")
            sb.Append(@"and L_Type=@L_Type ");
        else
            sb.Append(@"and (L_Type='06' or L_Type='10') ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (L_ModDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and L_ModDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and L_ModDate<@endDay ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_IP) LIKE '%' + upper(@KeyWord) + '%') or 
(upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(comp_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Office) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@KeyWord", KeyWord);
        oCmd.Parameters.AddWithValue("@L_Type", L_Type);
        if (startDay != "")
            oCmd.Parameters.AddWithValue("@startDay", DateTime.Parse(startDay));
        if (endDay != "")
            oCmd.Parameters.AddWithValue("@endDay", DateTime.Parse(endDay).AddDays(1));
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oda.Fill(ds);
         return ds;
    }

    public DataSet getMSModifyList(string pStart, string pEnd, string startDay, string endDay)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT COUNT(*) total from LogTable 
left join Member on M_Guid=L_Person and M_Status='A'
left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
left join CodeTable log_type on log_type.C_Group='06' and log_type.C_Item=L_Type   
where 1=1 ");

        if (L_Type != "")
            sb.Append(@"and L_Type=@L_Type ");
        else
            sb.Append(@"and (L_Type='08' or L_Type='09') ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (L_ModDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and L_ModDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and L_ModDate<@endDay ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_IP) LIKE '%' + upper(@KeyWord) + '%') or 
(upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(comp_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Office) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@"select * from (
	select ROW_NUMBER() over (order by L_ModDate desc,L_ID desc) itemNo,log_type.C_Item_cn LType,LogTable.*,Member.*,
	city_type.C_Item_cn City,comp_type.C_Item_cn Comp,
	(select top 1 RM_Year from ReportMonth where RM_ReportGuid=L_ModItemGuid and RM_Status='A') MYear,
	(select top 1 RM_Month from ReportMonth where RM_ReportGuid=L_ModItemGuid and RM_Status='A') MMonth,
	(select RS_Year from ReportSeason where RS_Guid=L_ModItemGuid and RS_Status='A') SYear,
	(select RS_Season from ReportSeason where RS_Guid=L_ModItemGuid and RS_Status='A') SSeason
	from LogTable
	left join Member on M_Guid=L_Person and M_Status='A'
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
	left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
	left join CodeTable log_type on log_type.C_Group='06' and log_type.C_Item=L_Type
    where 1=1 ");

        if (L_Type != "")
            sb.Append(@"and L_Type=@L_Type ");
        else
            sb.Append(@"and (L_Type='08' or L_Type='09') ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (L_ModDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and L_ModDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and L_ModDate<@endDay ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_IP) LIKE '%' + upper(@KeyWord) + '%') or 
(upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(comp_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Office) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@KeyWord", KeyWord);
        oCmd.Parameters.AddWithValue("@L_Type", L_Type);
        if (startDay != "")
            oCmd.Parameters.AddWithValue("@startDay", DateTime.Parse(startDay));
        if (endDay != "")
            oCmd.Parameters.AddWithValue("@endDay", DateTime.Parse(endDay).AddDays(1));
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oda.Fill(ds);
        return ds;
    }

    public DataSet getChangeContractorList(string pStart, string pEnd, string startDay, string endDay)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT COUNT(*) total from LogTable 
left join Member on M_Guid=L_Person and M_Status='A'
where L_Type='11' ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (L_ModDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and L_ModDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and L_ModDate<@endDay ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_IP) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_Desc) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@"select * from (
	select ROW_NUMBER() over (order by L_ModDate desc,L_ID desc) itemNo,LogTable.*,Member.*
	from LogTable
	left join Member on M_Guid=L_Person and M_Status='A'
	where L_Type='11' ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (L_ModDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and L_ModDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and L_ModDate<@endDay ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_IP) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_Desc) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@KeyWord", KeyWord);
        oCmd.Parameters.AddWithValue("@L_Type", L_Type);
        if (startDay != "")
            oCmd.Parameters.AddWithValue("@startDay", DateTime.Parse(startDay));
        if (endDay != "")
            oCmd.Parameters.AddWithValue("@endDay", DateTime.Parse(endDay).AddDays(1));
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oda.Fill(ds);
        return ds;
    }

    public DataSet getProjectInfo(string pStart, string pEnd, string startDay, string endDay)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT COUNT(*) total from LogTable 
left join Member on M_Guid=L_Person and M_Status='A'
left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
left join CodeTable log_type on log_type.C_Group='06' and log_type.C_Item=L_Type   
where L_Type='07' ");

        if (startDay != "" && endDay != "")
            sb.Append(@"and (L_ModDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and L_ModDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and L_ModDate<@endDay ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_IP) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_Desc) LIKE '%' + upper(@KeyWord) + '%') or 
(upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(comp_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Office) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@"select * from (
	select ROW_NUMBER() over (order by L_ModDate desc,L_ID desc) itemNo,log_type.C_Item_cn LType,LogTable.*,Member.*,
	city_type.C_Item_cn City,comp_type.C_Item_cn Comp
	from LogTable
	left join Member on M_Guid=L_Person and M_Status='A'
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
	left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
	left join CodeTable log_type on log_type.C_Group='06' and log_type.C_Item=L_Type
    where L_Type='07' ");
        
        if (startDay != "" && endDay != "")
            sb.Append(@"and (L_ModDate between @startDay and @endDay) ");
        else if (startDay != "")
            sb.Append(@"and L_ModDate>@startDay ");
        else if (endDay != "")
            sb.Append(@"and L_ModDate<@endDay ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_IP) LIKE '%' + upper(@KeyWord) + '%') or (upper(L_Desc) LIKE '%' + upper(@KeyWord) + '%') or 
(upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(comp_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Office) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@KeyWord", KeyWord);
        oCmd.Parameters.AddWithValue("@L_Type", L_Type);
        if (startDay != "")
            oCmd.Parameters.AddWithValue("@startDay", DateTime.Parse(startDay));
        if (endDay != "")
            oCmd.Parameters.AddWithValue("@endDay", DateTime.Parse(endDay).AddDays(1));
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oda.Fill(ds);
        return ds;
    }

    public DataTable SelectList()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select * from LogTable where L_Type=@L_Type ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oCmd.Parameters.AddWithValue("@L_Type", L_Type);
        oda.Fill(ds);
        return ds;
    }

    public void addLog()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"insert into LogTable (
L_Type,
L_Person,
L_IP,
L_ModItemGuid,
L_Desc,
L_ModDate
) values (
@L_Type,
@L_Person,
@L_IP,
@L_ModItemGuid,
@L_Desc,
@L_ModDate
) ";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@L_Type", L_Type);
        oCmd.Parameters.AddWithValue("@L_Person", L_Person);
        oCmd.Parameters.AddWithValue("@L_ModItemGuid", L_ModItemGuid);
        oCmd.Parameters.AddWithValue("@L_IP", L_IP);
        oCmd.Parameters.AddWithValue("@L_Desc", L_Desc);
        oCmd.Parameters.AddWithValue("@L_ModDate", DateTime.Now);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }
}