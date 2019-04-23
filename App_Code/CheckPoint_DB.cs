using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;

/// <summary>
/// CheckPoint_DB 的摘要描述
/// </summary>
public class CheckPoint_DB
{

    string KeyWord = string.Empty;
    public string _KeyWord
    {
        set { KeyWord = value; }
    }

    #region PushItem

    string P_ID = string.Empty;
    string P_Guid = string.Empty;
    string P_ParentId = string.Empty;
    string P_Period = string.Empty;
    string P_Type = string.Empty;
    string P_ItemName = string.Empty;
    string P_WorkRatio = string.Empty;
    string P_CreateId = string.Empty;
    string P_ModId = string.Empty;
    string P_Status = string.Empty;
    string P_ExFinish = string.Empty;
    string P_ExType = string.Empty;
    string P_ExDeviceType = string.Empty;

    DateTime P_CreateDate;
    DateTime P_ModDate;

    public string _P_ID
    {
        set { P_ID = value; }
    }
    public string _P_Guid
    {
        set { P_Guid = value; }
    }
    public string _P_ParentId
    {
        set { P_ParentId = value; }
    }
    public string _P_Period
    {
        set { P_Period = value; }
    }
    public string _P_Type
    {
        set { P_Type = value; }
    }
    public string _P_ItemName
    {
        set { P_ItemName = value; }
    }
    public string _P_WorkRatio
    {
        set { P_WorkRatio = value; }
    }
    public string _P_CreateId
    {
        set { P_CreateId = value; }
    }
    public string _P_ModId
    {
        set { P_ModId = value; }
    }
    public string _P_Status
    {
        set { P_Status = value; }
    }
    public string _P_ExFinish
    {
        set { P_ExFinish = value; }
    }
    public string _P_ExType
    {
        set { P_ExType = value; }
    }
    public string _P_ExDeviceType
    {
        set { P_ExDeviceType = value; }
    }
    public DateTime _P_CreateDate
    {
        set { P_CreateDate = value; }
    }
    public DateTime _P_ModDate
    {
        set { P_ModDate = value; }
    }
    #endregion
    #region Check_Point

    string CP_ID = string.Empty;
    string CP_Guid = string.Empty;
    string CP_ParentId = string.Empty;
    string CP_ProjectId = string.Empty;
    string CP_Point = string.Empty;
    string CP_ReserveYear = string.Empty;
    string CP_ReserveMonth = string.Empty;
    string CP_Desc = string.Empty;
    string CP_Process = string.Empty;
    string CP_CreateId = string.Empty;
    string CP_ModId = string.Empty;
    string CP_Status = string.Empty;
    string CP_RealProcess = string.Empty;
    string CP_Summary = string.Empty;
    string CP_BackwardDesc = string.Empty;

    DateTime CP_CreateDate;
    DateTime CP_ModDate;

    public string _CP_ID
    {
        set { CP_ID = value; }
    }
    public string _CP_Guid
    {
        set { CP_Guid = value; }
    }
    public string _CP_ParentId
    {
        set { CP_ParentId = value; }
    }
    public string _CP_ProjectId
    {
        set { CP_ProjectId = value; }
    }
    public string _CP_Point
    {
        set { CP_Point = value; }
    }
    public string _CP_ReserveYear
    {
        set { CP_ReserveYear = value; }
    }
    public string _CP_ReserveMonth
    {
        set { CP_ReserveMonth = value; }
    }
    public string _CP_Desc
    {
        set { CP_Desc = value; }
    }
    public string _CP_Process
    {
        set { CP_Process = value; }
    }
    public string _CP_CreateId
    {
        set { CP_CreateId = value; }
    }
    public string _CP_ModId
    {
        set { CP_ModId = value; }
    }
    public string _CP_Status
    {
        set { CP_Status = value; }
    }
    public string _CP_RealProcess
    {
        set { CP_RealProcess = value; }
    }
    public string _CP_Summary
    {
        set { CP_Summary = value; }
    }
    public string _CP_BackwardDesc
    {
        set { CP_BackwardDesc = value; }
    }
    public DateTime _CP_CreateDate
    {
        set { CP_CreateDate = value; }
    }
    public DateTime _CP_ModDate
    {
        set { CP_ModDate = value; }
    }
    #endregion

    public DataTable SelectList()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select * from PushItem
left join Check_Point on CP_ParentId=P_Guid
where CP_ProjectId=@CP_ProjectId and CP_Status='A' and P_Type=@P_Type and P_Period=@P_Period
order by P_ID ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@P_Period", P_Period);
        oCmd.Parameters.AddWithValue("@P_Type", P_Type);
        oCmd.Parameters.AddWithValue("@CP_ProjectId", CP_ProjectId);
        oda.Fill(ds);
        return ds;
    }

    public void addPushItem()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"
declare @rownum int = (select COUNT(*) from PushItem where P_Guid=@P_Guid and P_Status='A')

if(@rownum>0)
    begin
	    update PushItem set 
        P_ItemName=@P_ItemName,
        P_ExFinish=@P_ExFinish,
        P_ExType=@P_ExType,
        P_ExDeviceType=@P_ExDeviceType,
        P_ModId=@P_ModId,
        P_ModDate=@P_ModDate
        where P_Guid=@P_Guid
    end
else
    begin
	    insert into PushItem (
		    P_Guid,
		    P_ParentId,
            P_Period,
		    P_Type,
		    P_ItemName,
            P_ExFinish,
            P_ExType,
            P_ExDeviceType,
		    P_CreateId,
		    P_ModDate,
		    P_ModId,
		    P_Status
		    ) values (
		    @P_Guid,
		    @P_ParentId,
            @P_Period,
		    @P_Type,
		    @P_ItemName,
            @P_ExFinish,
            @P_ExType,
            @P_ExDeviceType,
		    @P_CreateId,
		    @P_ModDate,
		    @P_ModId,
		    @P_Status
		    )
    end ";

        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@P_Guid", P_Guid);
        oCmd.Parameters.AddWithValue("@P_ParentId", P_ParentId);
        oCmd.Parameters.AddWithValue("@P_Period", P_Period);
        oCmd.Parameters.AddWithValue("@P_Type", P_Type);
        oCmd.Parameters.AddWithValue("@P_ItemName", P_ItemName);
        oCmd.Parameters.AddWithValue("@P_CreateId", P_CreateId);
        oCmd.Parameters.AddWithValue("@P_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@P_ModId", P_ModId);
        oCmd.Parameters.AddWithValue("@P_Status", "A");
        oCmd.Parameters.AddWithValue("@P_ExFinish", P_ExFinish);
        oCmd.Parameters.AddWithValue("@P_ExType", P_ExType);
        oCmd.Parameters.AddWithValue("@P_ExDeviceType", P_ExDeviceType);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    public void addCheckPoint()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"
declare @rownum int = (select COUNT(*) from Check_Point where CP_Guid=@CP_Guid and CP_Status='A')

if(@rownum>0)
    begin
	    update Check_Point set 
        CP_Point=@CP_Point,
        CP_ReserveYear=@CP_ReserveYear,
        CP_ReserveMonth=@CP_ReserveMonth,
        CP_Desc=@CP_Desc,
        CP_ModDate=@CP_ModDate,
        CP_ModId=@CP_ModId
        where CP_Guid=@CP_Guid
    end
else
    begin
	    insert into Check_Point (
            CP_Guid,
            CP_ParentId,
            CP_ProjectId,
            CP_Point,
            CP_ReserveYear,
            CP_ReserveMonth,
            CP_Desc,
            CP_CreateId,
            CP_ModDate,
            CP_ModId,
            CP_Status
            ) values (
            @CP_Guid,
            @CP_ParentId,
            @CP_ProjectId,
            @CP_Point,
            @CP_ReserveYear,
            @CP_ReserveMonth,
            @CP_Desc,
            @CP_CreateId,
            @CP_ModDate,
            @CP_ModId,
            @CP_Status
            ) 
    end ";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@CP_Guid", CP_Guid);
        oCmd.Parameters.AddWithValue("@CP_ParentId", CP_ParentId);
        oCmd.Parameters.AddWithValue("@CP_ProjectId", CP_ProjectId);
        oCmd.Parameters.AddWithValue("@CP_Point", CP_Point);
        oCmd.Parameters.AddWithValue("@CP_ReserveYear", CP_ReserveYear);
        oCmd.Parameters.AddWithValue("@CP_ReserveMonth", CP_ReserveMonth);
        oCmd.Parameters.AddWithValue("@CP_Desc", CP_Desc);
        oCmd.Parameters.AddWithValue("@CP_CreateId", CP_CreateId);
        oCmd.Parameters.AddWithValue("@CP_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@CP_ModId", CP_ModId);
        oCmd.Parameters.AddWithValue("@CP_Status", "A");

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    public void deletePushItem()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"update PushItem set 
P_Status=@P_Status,
P_ModDate=@P_ModDate,
P_ModId=@P_ModId
where P_Guid=@P_Guid

update Check_Point set 
CP_Status=@P_Status,
CP_ModDate=@P_ModDate,
CP_ModId=@P_ModId
where CP_ParentId=@P_Guid
";

        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@P_Guid", P_Guid);
        oCmd.Parameters.AddWithValue("@P_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@P_ModId", P_ModId);
        oCmd.Parameters.AddWithValue("@P_Status", "D");

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    public void deleteCheckPoint()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"update Check_Point set 
CP_Status=@CP_Status,
CP_ModDate=@CP_ModDate,
CP_ModId=@CP_ModId
where CP_Guid=@CP_Guid
";

        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@CP_Guid", CP_Guid);
        oCmd.Parameters.AddWithValue("@CP_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@CP_ModId", CP_ModId);
        oCmd.Parameters.AddWithValue("@CP_Status", "D");

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    public DataTable getProgress(string ProjectId, string pType, string period)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
--抓各Type總合
--用於判斷是否為最後一筆row
select * into #tmp
from(
select COUNT(P_Type) TypeTotal,P_Type tp from Check_Point
left join PushItem on CP_ParentId=P_Guid
where CP_ProjectId=@ProjectId and CP_Status='A' and P_Period=@P_Period
group by P_Type
)#t

select P_Type,P_ItemName,P_WorkRatio,CP_Guid,CP_ParentId,CP_Point,CP_ReserveYear,RIGHT(REPLICATE('0', 1) + CP_ReserveMonth, 2) CP_ReserveMonth
,CP_Desc,CP_Process,TypeTotal 
from PushItem
left join Check_Point on CP_ParentId=P_Guid
left join #tmp on P_Type=tp

where CP_ProjectId=@ProjectId and CP_Status='A' and P_Period=@P_Period ");

        if (pType != "")
            sb.Append(@"and P_Type=@P_Type ");

        sb.Append(@"order by P_Type,P_ID,CP_ReserveYear,CP_ReserveMonth 
drop table #tmp
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@ProjectId", ProjectId);
        oCmd.Parameters.AddWithValue("@P_Period", period);
        oCmd.Parameters.AddWithValue("@P_Type", pType);
        oda.Fill(ds);
        return ds;
    }

    public void setWorkRatio()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"update PushItem set  
P_WorkRatio=@P_WorkRatio,
P_ModId=@P_ModId,
P_ModDate=@P_ModDate
where P_Guid=@P_Guid ";

        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@P_Guid", P_Guid);
        oCmd.Parameters.AddWithValue("@P_WorkRatio", P_WorkRatio);
        oCmd.Parameters.AddWithValue("@P_ModId", P_ModId);
        oCmd.Parameters.AddWithValue("@P_ModDate", DateTime.Now);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    public void setCP_Process()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"update Check_Point set  
CP_Process=@CP_Process,
CP_ModId=@CP_ModId,
CP_ModDate=@CP_ModDate
where CP_Guid=@CP_Guid ";

        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@CP_Guid", CP_Guid);
        oCmd.Parameters.AddWithValue("@CP_Process", CP_Process);
        oCmd.Parameters.AddWithValue("@CP_ModId", CP_ModId);
        oCmd.Parameters.AddWithValue("@CP_ModDate", DateTime.Now);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }


    public DataSet getExportInfo(string ProjectID)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select * from ProjectInfo
left join ProjectDate on PD_Type=I_City
where I_Guid=@ProjectID and I_Status='A'


select count(*) ItemCount,CP_ParentId cpid into #tmp
from Check_Point where CP_Status='A'
group by CP_ParentId

select * from PushItem
left join Check_Point on CP_ParentId=P_Guid and CP_Status='A'
left join #tmp on cpid=CP_ParentId
where P_ParentId=@ProjectID and P_Status='A' 
order by P_ID ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@ProjectID", ProjectID);
        oda.Fill(ds);
        return ds;
    }

 public DataSet getExportProgress(string ProjectID,string scriptStr)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select CP_ReserveYear Y,CONVERT(int,CP_ReserveMonth) M,P_Period,P_Type into #tmp
from PushItem 
left join Check_Point on CP_ParentId=P_Guid and CP_Status='A'
where P_ParentId=@ProjectID and P_Status='A' ");
        if (scriptStr != "")
            sb.Append("and " + scriptStr);
        sb.Append(@"
group by CP_ReserveYear,CONVERT(int,CP_ReserveMonth),P_Period,P_Type

select count(*) yc,Y into #tmp2
from #tmp
group by Y

select #tmp.Y,#tmp.M,yc from #tmp
left join #tmp2 on #tmp2.Y=#tmp.Y

select COUNT(*) ItemCount,P_Guid pid into #tmp3 from PushItem
left join Check_Point on CP_ParentId=P_Guid and CP_Status='A'
where P_ParentId=@ProjectID and P_Status='A' ");
        if (scriptStr != "")
            sb.Append("and " + scriptStr);
        sb.Append(@"
group by P_Guid

select P_ID,P_Period,P_Type,P_ItemName,P_WorkRatio,
CP_Point,CP_ReserveYear,CP_ReserveMonth,CP_Desc,CP_Process,P_Guid,P_ParentId,ItemCount from PushItem
left join Check_Point on CP_ParentId=P_Guid and CP_Status='A'
left join #tmp3 on pid=P_Guid
where P_ParentId=@ProjectID and P_Status='A'
order by P_ID

select CP_ReserveYear Y,CONVERT(int,CP_ReserveMonth) M,sum(CONVERT(float,CP_Process)) total
from PushItem
left join Check_Point on CP_ParentId=P_Guid and CP_Status='A'
where P_ParentId=@ProjectID and P_Status='A'");
        if (scriptStr != "")
            sb.Append("and " + scriptStr);
        sb.Append(@"
group by CP_ReserveYear,CONVERT(int,CP_ReserveMonth)

drop table #tmp
drop table #tmp2
drop table #tmp3 ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@ProjectID", ProjectID);
        oda.Fill(ds);
        return ds;
    }

    public DataTable getProgressTotal(string ProjectID)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
SELECT P_ID,P_Type,P_Period,CP_ReserveYear,CP_ReserveMonth,CP_Process FROM Check_Point
left join PushItem on P_Guid=CP_ParentId
where CP_ProjectId=@ProjectID
order by CONVERT(int,CP_ReserveMonth) 
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@ProjectID", ProjectID);
        oda.Fill(ds);
        return ds;
    }


    public DataTable getPeriodDate(string ProjectID)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
SELECT  I_1_Sdate,I_1_Edate,I_2_Sdate,I_2_Edate,I_3_Sdate,I_3_Edate FROM ProjectInfo
  where I_Guid=@ProjectID
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@ProjectID", ProjectID);
        oda.Fill(ds);
        return ds;
    }

    public DataTable getPushItemName(string item)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select C_Item_cn from CodeTable where C_Group='07' and C_Item=@item ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@item", item);
        oda.Fill(ds);
        return ds;
    }

    public DataTable getExPandTypeName(string item)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select C_Item_cn from CodeTable where C_Group='09' and C_Item=@item ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@item", item);
        oda.Fill(ds);
        return ds;
    }

    public void updateSeasonInfo()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"update Check_Point set  
--CP_Summary=@CP_Summary,
--CP_BackwardDesc=@CP_BackwardDesc,
CP_RealProcess=@CP_RealProcess
where CP_Guid=@CP_Guid ";

        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@CP_Guid", CP_Guid);
        oCmd.Parameters.AddWithValue("@CP_RealProcess", CP_RealProcess);
        oCmd.Parameters.AddWithValue("@CP_Summary", CP_Summary);
        oCmd.Parameters.AddWithValue("@CP_BackwardDesc", CP_BackwardDesc);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    public DataTable getExFinish()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select C_Item_cn from CodeTable where C_Group='07' and C_Item=@item ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@P_ParentId", P_ParentId);
        oda.Fill(ds);
        return ds;
    }

    public DataTable getPushitemExFinish(string ProjectID)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select P_Period,P_ItemName,C_Item_cn,P_ExFinish from PushItem
left join CodeTable on C_Group='09' and C_Item=P_ItemName
where P_Type='04' and P_Status='A' and P_ParentId=@ProjectID and P_ItemName<>'99'
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@ProjectID", ProjectID);
        oda.Fill(ds);
        return ds;
    }
}