using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;

/// <summary>
/// ReportSeasonV2_DB 的摘要描述
/// </summary>
public class ReportSeasonV2_DB
{

    #region 私用
    string RS_ID = string.Empty;
    string RS_Guid = string.Empty;
    string RS_PorjectGuid = string.Empty;
    string RS_Stage = string.Empty;
    string RS_Year = string.Empty;
    string RS_Season = string.Empty;
    string RS_StartDay = string.Empty;
    string RS_EndDay = string.Empty;
    string RS_TotalMonth = string.Empty;
    string RS_CostDesc = string.Empty;
    string RS_Type01Money = string.Empty;
    string RS_Type01Real = string.Empty;
    string RS_Type01RealRate = string.Empty;
    string RS_Type02Money = string.Empty;
    string RS_Type02Real = string.Empty;
    string RS_Type02RealRate = string.Empty;
    string RS_Type03Money = string.Empty;
    string RS_Type03Real = string.Empty;
    string RS_Type03RealRate = string.Empty;
    string RS_Type04Money = string.Empty;
    string RS_Type04Real = string.Empty;
    string RS_Type04RealRate = string.Empty;
    string RS_AllSchedule = string.Empty;
    string RS_AllRealSchedule = string.Empty;
    string RS_01Schedule = string.Empty;
    string RS_01RealSchedule = string.Empty;
    string RS_02Schedule = string.Empty;
    string RS_02RealSchedule = string.Empty;
    string RS_03Schedule = string.Empty;
    string RS_03RealSchedule = string.Empty;
    string RS_04Schedule = string.Empty;
    string RS_04RealSchedule = string.Empty;
    string RS_CheckPointData = string.Empty;
    string RS_PushItemDesc = string.Empty;
    string RS_03Type01C = string.Empty;
    string RS_03Type01S = string.Empty;
    string RS_03Type02C = string.Empty;
    string RS_03Type02S = string.Empty;
    string RS_03Type03C = string.Empty;
    string RS_03Type03S = string.Empty;
    string RS_03Type04C = string.Empty;
    string RS_03Type04S = string.Empty;
    string RS_03Type05C = string.Empty;
    string RS_03Type05S = string.Empty;
    DateTime RS_CreateDate;
    string RS_CreateId = string.Empty;
    DateTime RS_ModDate;
    string RS_ModId = string.Empty;
    string RS_Status = string.Empty;
    #endregion
    #region 公用
    public string _RS_ID
    {
        set { RS_ID = value; }
    }
    public string _RS_Guid
    {
        set { RS_Guid = value; }
    }
    public string _RS_PorjectGuid
    {
        set { RS_PorjectGuid = value; }
    }
    public string _RS_Stage
    {
        set { RS_Stage = value; }
    }
    public string _RS_Year
    {
        set { RS_Year = value; }
    }
    public string _RS_Season
    {
        set { RS_Season = value; }
    }
    public string _RS_StartDay
    {
        set { RS_StartDay = value; }
    }
    public string _RS_EndDay
    {
        set { RS_EndDay = value; }
    }
    public string _RS_TotalMonth
    {
        set { RS_TotalMonth = value; }
    }
    public string _RS_CostDesc
    {
        set { RS_CostDesc = value; }
    }
    public string _RS_Type01Money
    {
        set { RS_Type01Money = value; }
    }
    public string _RS_Type01Real
    {
        set { RS_Type01Real = value; }
    }
    public string _RS_Type01RealRate
    {
        set { RS_Type01RealRate = value; }
    }
    public string _RS_Type02Money
    {
        set { RS_Type02Money = value; }
    }
    public string _RS_Type02Real
    {
        set { RS_Type02Real = value; }
    }
    public string _RS_Type02RealRate
    {
        set { RS_Type02RealRate = value; }
    }
    public string _RS_Type03Money
    {
        set { RS_Type03Money = value; }
    }
    public string _RS_Type03Real
    {
        set { RS_Type03Real = value; }
    }
    public string _RS_Type03RealRate
    {
        set { RS_Type03RealRate = value; }
    }
    public string _RS_Type04Money
    {
        set { RS_Type04Money = value; }
    }
    public string _RS_Type04Real
    {
        set { RS_Type04Real = value; }
    }
    public string _RS_Type04RealRate
    {
        set { RS_Type04RealRate = value; }
    }
    public string _RS_AllSchedule
    {
        set { RS_AllSchedule = value; }
    }
    public string _RS_AllRealSchedule
    {
        set { RS_AllRealSchedule = value; }
    }
    public string _RS_01Schedule
    {
        set { RS_01Schedule = value; }
    }
    public string _RS_01RealSchedule
    {
        set { RS_01RealSchedule = value; }
    }
    public string _RS_02Schedule
    {
        set { RS_02Schedule = value; }
    }
    public string _RS_02RealSchedule
    {
        set { RS_02RealSchedule = value; }
    }
    public string _RS_03Schedule
    {
        set { RS_03Schedule = value; }
    }
    public string _RS_03RealSchedule
    {
        set { RS_03RealSchedule = value; }
    }
    public string _RS_04Schedule
    {
        set { RS_04Schedule = value; }
    }
    public string _RS_04RealSchedule
    {
        set { RS_04RealSchedule = value; }
    }
    public string _RS_CheckPointData
    {
        set { RS_CheckPointData = value; }
    }
    public string _RS_PushItemDesc
    {
        set { RS_PushItemDesc = value; }
    }
    public string _RS_03Type01C
    {
        set { RS_03Type01C = value; }
    }
    public string _RS_03Type01S
    {
        set { RS_03Type01S = value; }
    }
    public string _RS_03Type02C
    {
        set { RS_03Type02C = value; }
    }
    public string _RS_03Type02S
    {
        set { RS_03Type02S = value; }
    }
    public string _RS_03Type03C
    {
        set { RS_03Type03C = value; }
    }
    public string _RS_03Type03S
    {
        set { RS_03Type03S = value; }
    }
    public string _RS_03Type04C
    {
        set { RS_03Type04C = value; }
    }
    public string _RS_03Type04S
    {
        set { RS_03Type04S = value; }
    }
    public string _RS_03Type05C
    {
        set { RS_03Type05C = value; }
    }
    public string _RS_03Type05S
    {
        set { RS_03Type05S = value; }
    }
    public DateTime _RS_CreateDate
    {
        set { RS_CreateDate = value; }
    }
    public string _RS_CreateId
    {
        set { RS_CreateId = value; }
    }
    public DateTime _RS_ModDate
    {
        set { RS_ModDate = value; }
    }
    public string _RS_ModId
    {
        set { RS_ModId = value; }
    }
    public string _RS_Status
    {
        set { RS_Status = value; }
    }
    #endregion

    /// -----------------------------------------
    /// 建立者: Nick
    /// 功　能: 查詢承辦人季報列表
    /// -----------------------------------------
    public DataSet getSeasonList(string mGuid, string pStart, string pEnd)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
--傳進來的參數
declare @strYear nvarchar(4)=@RS_Year --yyyy 西元年4碼
declare @strSeason nvarchar(1)=@RS_Season --季 1.2.3.4  1碼
declare @strStage nvarchar(1)=@RS_Stage --期數 1.2.3 1碼
declare @M_Guid nvarchar(50)=@mGuid --承辦人M_ID

create table #tmpAll(
	RS_ID bigint
	,RS_Guid nvarchar(50)
	,RS_Stage nvarchar(1)
	,RS_Year nvarchar(4)
	,RS_Season nvarchar(1)
	,RS_CreateDate datetime
	,RC_CheckType nvarchar(2) NULL
	,RC_Status nvarchar(2) NULL
)

--根據傳進來的承辦人M_ID撈出他的M_GUID跟計劃I_Guid
declare @I_Guid nvarchar(50);
select @I_Guid=I_Guid from ProjectInfo where  I_People=@M_Guid and I_Flag='Y' and I_Status='A'

--先根據條件過濾出ReportSeason的資料
insert into #tmpAll(RS_ID,RS_Guid,RS_Stage,RS_Year,RS_Season,RS_CreateDate)
select RS_ID,RS_Guid,RS_Stage,RS_Year,RS_Season,RS_CreateDate
from ReportSeason
where RS_Status='A' and RS_PorjectGuid=@I_Guid --查尋WHERE條件加這裡

--從過濾完的季報資料表資料去跑迴圈判斷這個季報審核狀態
declare @whilerowcount int;--迴圈數
declare @whileChkRows int;--迴圈內判斷ReportCheck
declare @whileChkGuid nvarchar(50);--迴圈內判斷ReportCheck
select * into #tmpwhile from #tmpAll
select @whilerowcount=count(*) from #tmpwhile
-------------------------迴圈開始-------------------------
set rowcount 1;
while @whilerowcount>0
	begin
		select @whilerowcount=@whilerowcount-1;
		select @whileChkGuid=RS_Guid from #tmpwhile
		--審核通過
		select @whileChkRows=count(*) from ReportCheck where  RC_ReportGuid=@whileChkGuid and RC_CheckType='Y' and RC_Status='A'
		if @whileChkRows>0
			begin
				update #tmpAll set RC_CheckType='Y',RC_Status='A' where RS_Guid=@whileChkGuid
			end
		--審核中
		select @whileChkRows=count(*) from ReportCheck where  RC_ReportGuid=@whileChkGuid and (RC_CheckType='' or RC_CheckType is null) and RC_Status='A'
		if @whileChkRows>0
			begin
				update #tmpAll set RC_Status='A' where RS_Guid=@whileChkGuid
			end
		delete from #tmpwhile;
	end
set rowcount 0;
-------------------------迴圈結束-------------------------

--最後的#tmpAll 你就直接用去做分頁
--RS_ID,RS_Guid,RS_Stage,RS_Year,RS_Season,RC_CheckType,RC_Status
--RC_CheckType	&&	RC_Status
--		NULL			NULL		=>未送審(草稿狀態(儲存而已或者是被退回))
--		NULL			A			=>審核中
--		Y				A			=>審核通過

select count(*) as total from #tmpAll where 1=1 ");
        if (RS_Year != "")
            sb.Append(@"and RS_Year=@RS_Year ");
        if (RS_Season != "")
            sb.Append(@"and RS_Season=@RS_Season ");
        if (RS_Stage != "")
            sb.Append(@"and RS_Stage=@RS_Stage ");

        sb.Append(@"
select * from (
           select ROW_NUMBER() over (order by RS_CreateDate desc) itemNo,#tmpAll.*
		   from #tmpAll
           where 1=1 ");

        if (RS_Year != "")
            sb.Append(@"and RS_Year=@RS_Year ");
        if (RS_Season != "")
            sb.Append(@"and RS_Season=@RS_Season ");
        if (RS_Stage != "")
            sb.Append(@"and RS_Stage=@RS_Stage ");

        sb.Append(@"
)#tmp where itemNo between @pStart and @pEnd

drop table #tmpwhile
drop table #tmpAll ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@mGuid", mGuid);
        oCmd.Parameters.AddWithValue("@RS_Year", RS_Year);
        oCmd.Parameters.AddWithValue("@RS_Season", RS_Season);
        oCmd.Parameters.AddWithValue("@RS_Stage", RS_Stage);
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oda.Fill(ds);
        return ds;
    }

    /// -----------------------------------------
    /// 建立者: Nick
    /// 功　能: 季報詳細資料
    /// -----------------------------------------
    public DataSet getSeasonDetail(string mGuid, string year, string season, string stage)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
declare @PerGuid nvarchar(50)=@mGuid
declare @ProjectID nvarchar(50)=(select I_Guid from ProjectInfo where I_People=@PerGuid)
declare @rsID nvarchar(50)=(select RS_Guid from ReportSeason where RS_PorjectGuid=@ProjectID and RS_Year=@year and RS_Season=@season and RS_Stage=@stage)

--查詢基本資料
select 
city.C_Item_cn as CityName,
Info.* 
from ProjectInfo as Info
left join CodeTable as city on city.C_Group='02' and I_City=city.C_Item
where I_Guid=@ProjectID

--人員基本資料
select 
e.M_Name as UserName,
e.M_Tel as UserTel,
m.M_Name as ManagerName 
from Member as e
left join Member as m on m.M_Guid=e.M_Manager_ID
where e.M_Guid=@PerGuid

--查核點
select P_Guid,P_Type,P_Period,P_ItemName,P_WorkRatio,P_ExFinish,EF_Finish,CP_Guid,CP_Point,CP_ReserveYear,CP_ReserveMonth,CP_Desc,CP_Process,CP_RealProcess 
from PushItem
left join Check_Point on CP_ParentId=P_Guid and CP_Status='A'
left join ExpandFinish on EF_ReportId=@rsID and EF_PushitemId=P_Guid and EF_Status='A'
where P_ParentId=@ProjectID and P_Status='A' and P_Period=@stage
order by P_Type,P_ID,CONVERT(int,CP_ReserveYear),CONVERT(int,CP_ReserveMonth)
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@mGuid", mGuid);
        oCmd.Parameters.AddWithValue("@year", year);
        oCmd.Parameters.AddWithValue("@season", season);
        oCmd.Parameters.AddWithValue("@stage", stage);
        oda.Fill(ds);
        return ds;
    }


    /// -----------------------------------------
    /// 建立者: Nick
    /// 功　能: 查詢預定進度&實際進度
    /// -----------------------------------------
    public DataSet getSeasonProcess(string mGuid, string year, string season, string stage)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
declare @M_Guid nvarchar(50)=@mGuid
declare @RS_Stage nvarchar(50)=@stage
declare @RS_Year nvarchar(50)=@year
declare @RS_Season nvarchar(50)=@season

create table #tmpPro(
	P_Guid nvarchar(50)
    ,P_ItemName nvarchar(100)
	,P_Type nvarchar(50)
	,P_Period nvarchar(50)
	,P_WorkRatio nvarchar(max)
	,CP_Point nvarchar(50)
	,CP_ReserveYear nvarchar(50)
	,CP_ReserveMonth nvarchar(50)
	,CP_Desc nvarchar(max)
	,CP_Process nvarchar(50)
	,CP_RealProcess nvarchar(5)
	,CP_Summary nvarchar(max)
	,CP_BackwardDesc nvarchar(max)
	,CP_showhide nvarchar(2) --Y 顯示input / N 不顯示input
)
create table #tmpProTMP(
    P_ItemName nvarchar(100)
	,P_Type nvarchar(50)
	,P_Period nvarchar(50)
	,CP_ReserveYear nvarchar(50)
	,CP_ReserveMonth nvarchar(50)
	,CP_Desc nvarchar(max)
	,CP_Process nvarchar(5)
	,CP_RealProcess nvarchar(5)
)

create table #tmpNoVal(
	P_Type nvarchar(50)
	,CP_Process nvarchar(50)
	,CP_RealProcess nvarchar(5)
)

--季轉換成月
declare @mth nvarchar(2);
if(@RS_Season='1')
	begin
		set @mth='3';
	end
if(@RS_Season='2')
	begin
		set @mth='6';
	end
if(@RS_Season='3')
	begin
		set @mth='9';
	end
if(@RS_Season='4')
	begin
		set @mth='12';
	end

--根據人員GUID 撈出計畫GUID
declare @RS_PorjectGuid nvarchar(50)
select @RS_PorjectGuid=I_Guid from ProjectInfo where I_People=@M_Guid and I_Status='A' and I_Flag='Y'

--推動項目暫存表
SELECT * into #tmpp from PushItem
where P_ParentId=@RS_PorjectGuid and P_Period=@RS_Stage and P_Status='A'


--撈出所有的推動項目、查核點、預定工作進度
insert into #tmpPro(P_Guid,P_ItemName,P_Type,P_Period,CP_Point,P_WorkRatio,CP_ReserveYear,CP_ReserveMonth,CP_Desc,CP_Process,CP_RealProcess,CP_Summary,CP_BackwardDesc,CP_showhide)
select P_Guid,P_ItemName,P_Type,P_Period,CP_Point,P_WorkRatio,CP_ReserveYear,CP_ReserveMonth,CP_Desc
,case when CP_Process is null then '0' when CP_Process='' then '0' else CP_Process end as CP_Process
,case when CP_RealProcess is null then '0' when CP_RealProcess='' then '0' else CP_RealProcess end as CP_RealProcess
,CP_Summary,CP_BackwardDesc,
case when CP_ReserveMonth=@mth and CP_ReserveYear=@RS_Year and P_Period=@RS_Stage then 'Y' else 'N' end as CP_showhide
from #tmpp left join Check_Point on P_Guid=CP_ParentId and CP_Status='A'
where  P_Period=@RS_Stage and CP_ProjectId=@RS_PorjectGuid and CP_Status='A'  -- and CP_ReserveMonth=@mth and CP_ReserveYear=@RS_Year and


--算出當季
declare @while_rowcount int=0;
declare @w_P_Guid nvarchar(50)=''
declare @w_P_ParentID nvarchar(50)=''
declare @w_P_Period nvarchar(50)=''
declare @w_P_Type nvarchar(50)=''
declare @w_P_ItemName nvarchar(50)=''
select @while_rowcount = count(*) from #tmpp
declare @rcount int=0;
declare @newYM nvarchar(10)='';
declare @minYM nvarchar(10)='';
declare @strYM nvarchar(5)='';
while @while_rowcount>0
	begin
		
		select top 1 @w_P_Guid = P_Guid,@w_P_ParentID = P_ParentID,@w_P_Period=P_Period,@w_P_Type=P_Type,@w_P_ItemName=P_ItemName
		from #tmpp
		
		select @minYM = Min(convert(int,CP_ReserveYear+ case when LEN(CP_ReserveMonth)=1 then '0'+CP_ReserveMonth else CP_ReserveMonth end)) from #tmpPro where P_Guid=@w_P_Guid
		select @strYM = (case when LEN(@mth)=1 then @RS_Year+'0'+@mth else @RS_Year+@mth end)
		--if @mth='3'
		if @minYM=@strYM
			begin
                insert into #tmpProTMP(P_ItemName,P_Type,P_Period,CP_ReserveYear,CP_ReserveMonth,CP_Desc,CP_Process,CP_RealProcess)
				select @w_P_ItemName,@w_P_Type,@w_P_Period,CP_ReserveYear,CP_ReserveMonth,CP_Desc
				,case when CP_Process is null then '0' when CP_Process='' then '0' else CP_Process end as CP_Process
				,case when CP_RealProcess is null then '0' when CP_RealProcess='' then '0' else CP_RealProcess end as CP_RealProcess
				from Check_Point 
				where  CP_ReserveYear=@RS_Year and CP_ReserveMonth=@mth and CP_ProjectId=@RS_PorjectGuid and CP_ParentId =@w_P_Guid and CP_Status='A'
			end
		else
			begin
				select @rcount = count(*)
				from Check_Point 
				where  CP_ReserveYear=@RS_Year and CP_ReserveMonth=@mth and CP_ProjectId=@RS_PorjectGuid and CP_ParentId =@w_P_Guid and CP_Status='A'
				if @rcount>0
					begin
						insert into #tmpProTMP(P_ItemName,P_Type,P_Period,CP_ReserveYear,CP_ReserveMonth,CP_Desc,CP_Process,CP_RealProcess)
						select @w_P_ItemName,@w_P_Type,@w_P_Period,@RS_Year,CP_ReserveMonth=@mth,CP_Desc
						,case 
							when CP_Process='' then '0' 
							else CP_Process 
						end as CP_Process
						,case 
							when CP_RealProcess='' then '0' 
							else CP_RealProcess 
						end as CP_RealProcess
						from Check_Point 
						where  CP_ReserveYear=@RS_Year and CP_ReserveMonth=@mth and CP_ProjectId=@RS_PorjectGuid and CP_ParentId =@w_P_Guid and CP_Status='A'

					end
				else
					--沒有當季查核點 預定進度就要往前加
					begin
						insert into #tmpProTMP(P_ItemName,P_Type,P_Period,CP_ReserveYear,CP_ReserveMonth,CP_Desc,CP_Process,CP_RealProcess)
						select @w_P_ItemName,@w_P_Type,@w_P_Period,@RS_Year,@mth,''--,'20'--CP_Desc
						--,(select top 1 isnull(CP_Process,0) from Check_Point where  CP_ParentId =@w_P_Guid and CP_ProjectId=@w_P_ParentID and CP_ReserveYear=@RS_Year  and CP_ProjectId=@RS_PorjectGuid and convert(int,CP_ReserveMonth)<convert(int,@mth) and CP_Status='A' order by convert(int,isnull(CP_ReserveYear,0))+convert(int,isnull(CP_ReserveMonth,0)) desc  ) as CP_Process 
						--,(select top 1 isnull(CP_RealProcess,0) from Check_Point where  CP_ParentId =@w_P_Guid and CP_ProjectId=@w_P_ParentID and CP_ReserveYear=@RS_Year  and CP_ProjectId=@RS_PorjectGuid and convert(int,CP_ReserveMonth)<convert(int,@mth) and CP_Status='A' order by convert(int,isnull(CP_ReserveYear,0))+convert(int,isnull(CP_ReserveMonth,0)) desc ) as CP_RealProcess 
						,(select top 1 isnull(CP_Process,0) from Check_Point where  CP_ParentId =@w_P_Guid and CP_ProjectId=@w_P_ParentID   and CP_ProjectId=@RS_PorjectGuid and convert(int,CP_ReserveYear + case when LEN(CP_ReserveMonth)=1 then '0'+CP_ReserveMonth else CP_ReserveMonth end )<convert(int,@strYM) and CP_Status='A' order by convert(nvarchar(3),isnull(CP_ReserveYear,'000'))+case when LEN(isnull(CP_ReserveMonth,'00'))=1 then '0'+isnull(CP_ReserveMonth,'00') else isnull(CP_ReserveMonth,'00') end desc  ) as CP_Process 
						,(select top 1 isnull(CP_RealProcess,0) from Check_Point where  CP_ParentId =@w_P_Guid and CP_ProjectId=@w_P_ParentID   and CP_ProjectId=@RS_PorjectGuid and convert(int,CP_ReserveYear + case when LEN(CP_ReserveMonth)=1 then '0'+CP_ReserveMonth else CP_ReserveMonth end )<convert(int,@strYM) and CP_Status='A' order by convert(nvarchar(3),isnull(CP_ReserveYear,'000'))+case when LEN(isnull(CP_ReserveMonth,'00'))=1 then '0'+isnull(CP_ReserveMonth,'00') else isnull(CP_ReserveMonth,'00') end desc ) as CP_RealProcess 
						
						insert into #tmpNoVal(P_Type,CP_Process,CP_RealProcess)
						select @w_P_Type
						--,(select top 1 isnull(CP_Process,0) from Check_Point where  CP_ParentId =@w_P_Guid and CP_ProjectId=@w_P_ParentID and CP_ReserveYear=@RS_Year  and CP_ProjectId=@RS_PorjectGuid and convert(int,CP_ReserveMonth)<convert(int,@mth) and CP_Status='A' order by convert(int,isnull(CP_ReserveYear,0))+convert(int,isnull(CP_ReserveMonth,0)) desc  ) as CP_Process 
						--,(select top 1 isnull(CP_RealProcess,0) from Check_Point where  CP_ParentId =@w_P_Guid and CP_ProjectId=@w_P_ParentID and CP_ReserveYear=@RS_Year  and CP_ProjectId=@RS_PorjectGuid and convert(int,CP_ReserveMonth)<convert(int,@mth) and CP_Status='A' order by convert(int,isnull(CP_ReserveYear,0))+convert(int,isnull(CP_ReserveMonth,0)) desc  ) as CP_RealProcess 
						,(select top 1 isnull(CP_Process,0) from Check_Point where  CP_ParentId =@w_P_Guid and CP_ProjectId=@w_P_ParentID   and CP_ProjectId=@RS_PorjectGuid and convert(int,CP_ReserveYear + case when LEN(CP_ReserveMonth)=1 then '0'+CP_ReserveMonth else CP_ReserveMonth end )<convert(int,@strYM) and CP_Status='A' order by convert(nvarchar(3),isnull(CP_ReserveYear,'000'))+case when LEN(isnull(CP_ReserveMonth,'00'))=1 then '0'+isnull(CP_ReserveMonth,'00') else isnull(CP_ReserveMonth,'00') end desc  ) as CP_Process 
						,(select top 1 isnull(CP_RealProcess,0) from Check_Point where  CP_ParentId =@w_P_Guid and CP_ProjectId=@w_P_ParentID  and CP_ProjectId=@RS_PorjectGuid and convert(int,CP_ReserveYear + case when LEN(CP_ReserveMonth)=1 then '0'+CP_ReserveMonth else CP_ReserveMonth end )<convert(int,@strYM) and CP_Status='A' order by convert(nvarchar(3),isnull(CP_ReserveYear,'000'))+case when LEN(isnull(CP_ReserveMonth,'00'))=1 then '0'+isnull(CP_ReserveMonth,'00') else isnull(CP_ReserveMonth,'00') end desc  ) as CP_RealProcess 
					end
			end

			delete from #tmpp where P_Guid=@w_P_Guid and P_ParentID=@w_P_ParentID and P_Period=@w_P_Period and P_Type=@w_P_Type and P_ItemName=@w_P_ItemName
			select @while_rowcount = @while_rowcount-1
	end
	
--整體(四大項加總後) 平均數
select P_Period,CP_ReserveYear,CP_ReserveMonth,convert(float,sum(convert(float,isnull(CP_Process,0)))/4) as avgVal ,convert(float,sum(convert(float,isnull(CP_RealProcess,0)))/4) as avgRealVal
into #tPAll from #tmpProTMP group by P_Period,CP_ReserveYear,CP_ReserveMonth

--分四大項的預定累計進度
select P_Type,P_Period,CP_ReserveYear,CP_ReserveMonth,sum(convert(float,isnull(CP_Process,0))) as sumVal,sum(convert(float,isnull(CP_RealProcess,0))) as sumRealVal 
into #tP from #tmpProTMP group by P_Type,P_Period,CP_ReserveYear,CP_ReserveMonth

--最後把四大項及整體的預定&實際進度湊成一筆資料回傳
select 
'01' as typ01
,isnull((select c.sumVal from #tP c where c.P_Type='01'),0) as sumVal01
,isnull((select d.sumRealVal from #tP d where d.P_Type='01'),0) as sumRealVal01
,'02' as typ02
,isnull((select f.sumVal from #tP f where f.P_Type='02'),0) as sumVal02
,isnull((select g.sumRealVal from #tP g where g.P_Type='02'),0) as sumRealVal02
,'03' as typ03
,isnull((select i.sumVal from #tP i where i.P_Type='03'),0) as sumVal03
,isnull((select j.sumRealVal from #tP j where j.P_Type='03'),0) as sumRealVal03
,'04' as typ04
,isnull((select m.sumVal from #tP m where m.P_Type='04'),0) as sumVal04
,isnull((select n.sumRealVal from #tP n where n.P_Type='04'),0) as sumRealVal04
,isnull((select k.avgVal from #tPAll k),0) as sumValAll
,isnull((select l.avgRealVal from #tPAll l),0) as sumRealValAll


--這是把三大項當季沒有 但是上一季有的實際值總和抓出來
--如果當季沒有的查核點 就直接加總這個職 合計才會對
select P_Type,SUM(convert(float,CP_RealProcess)) as CP_RealProcess into #tN from #tmpNoVal group by P_Type
select 
'01' as typ01
,isnull((select b.CP_RealProcess from #tN b where b.P_Type='01'),0) as sumRealVal01
,'02' as typ02
,isnull((select d.CP_RealProcess from #tN d where d.P_Type='02'),0) as sumRealVal02
,'03' as typ03
,isnull((select f.CP_RealProcess from #tN f where f.P_Type='03'),0) as sumRealVal03
,'04' as typ04
,isnull((select h.CP_RealProcess from #tN h where h.P_Type='04'),0) as sumRealVal04

drop table #tmpp
drop table #tmpPro
drop table #tmpProTMP
drop table #tP
drop table #tPAll
drop table #tmpNoVal
drop table #tN

");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@mGuid", mGuid);
        oCmd.Parameters.AddWithValue("@year", year);
        oCmd.Parameters.AddWithValue("@season", season);
        oCmd.Parameters.AddWithValue("@stage", stage);
        oda.Fill(ds);
        return ds;
    }

    /// -----------------------------------------
    /// 建立者: Nick
    /// 功　能: 確認計畫是否定稿
    /// -----------------------------------------
    public DataTable checkProjectFlag(string mGuid)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select I_Flag from ProjectInfo where I_People=@mGuid and I_Status='A' ");


        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@mGuid", mGuid);
        oda.Fill(ds);
        return ds;
    }

    /// -----------------------------------------
    /// 建立者: Nick
    /// 功　能: 查詢季報詳細資料
    /// -----------------------------------------
    public DataTable getSeasonInfo(string mGuid, string year, string season, string stage)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
declare @PerGuid nvarchar(50)=@mGuid
declare @ProjectID nvarchar(50)=(select I_Guid from ProjectInfo where I_People=@PerGuid)

SELECT RS_ID,RS_Guid,RS_Year,RS_Season,RS_Stage,RS_CostDesc,RS_Type01Real,RS_Type02Real,RS_Type03Real,RS_Type04Real,
RS_03Type01C,RS_03Type02C,RS_03Type03C,RS_03Type04C,RS_03Type05C
FROM ReportSeason where RS_PorjectGuid=@ProjectID and RS_Status='A' ");

        if (year != "")
            sb.Append(@"and RS_Year=@RS_Year ");
        if (season != "")
            sb.Append(@"and RS_Season=@RS_Season ");
        if (stage != "")
            sb.Append(@"and RS_Stage=@RS_Stage ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@mGuid", mGuid);
        oCmd.Parameters.AddWithValue("@RS_Year", year);
        oCmd.Parameters.AddWithValue("@RS_Season", season);
        oCmd.Parameters.AddWithValue("@RS_Stage", stage);
        oda.Fill(ds);
        return ds;
    }


    /// -----------------------------------------
    /// 建立者: Nick
    /// 功　能: 新增季報
    /// -----------------------------------------
    public void addSeason()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"insert into ReportSeason (
RS_Guid,
RS_PorjectGuid,
RS_Stage,
RS_Year,
RS_Season,
RS_StartDay,
RS_EndDay,
RS_TotalMonth,
RS_CostDesc,
RS_Type01Money,
RS_Type01Real,
RS_Type01RealRate,
RS_Type02Money,
RS_Type02Real,
RS_Type02RealRate,
RS_Type03Money,
RS_Type03Real,
RS_Type03RealRate,
RS_Type04Money,
RS_Type04Real,
RS_Type04RealRate,
RS_AllSchedule,
RS_AllRealSchedule,
RS_01Schedule,
RS_01RealSchedule,
RS_02Schedule,
RS_02RealSchedule,
RS_03Schedule,
RS_03RealSchedule,
RS_04Schedule,
RS_04RealSchedule,
RS_CheckPointData,
RS_PushItemDesc,
RS_03Type01C,
RS_03Type01S,
RS_03Type02C,
RS_03Type02S,
RS_03Type03C,
RS_03Type03S,
RS_03Type04C,
RS_03Type04S,
RS_03Type05C,
RS_03Type05S,
RS_CreateId,
RS_ModDate,
RS_ModId,
RS_Status
) values (
@RS_Guid,
@RS_PorjectGuid,
@RS_Stage,
@RS_Year,
@RS_Season,
@RS_StartDay,
@RS_EndDay,
@RS_TotalMonth,
@RS_CostDesc,
@RS_Type01Money,
@RS_Type01Real,
@RS_Type01RealRate,
@RS_Type02Money,
@RS_Type02Real,
@RS_Type02RealRate,
@RS_Type03Money,
@RS_Type03Real,
@RS_Type03RealRate,
@RS_Type04Money,
@RS_Type04Real,
@RS_Type04RealRate,
@RS_AllSchedule,
@RS_AllRealSchedule,
@RS_01Schedule,
@RS_01RealSchedule,
@RS_02Schedule,
@RS_02RealSchedule,
@RS_03Schedule,
@RS_03RealSchedule,
@RS_04Schedule,
@RS_04RealSchedule,
@RS_CheckPointData,
@RS_PushItemDesc,
@RS_03Type01C,
@RS_03Type01S,
@RS_03Type02C,
@RS_03Type02S,
@RS_03Type03C,
@RS_03Type03S,
@RS_03Type04C,
@RS_03Type04S,
@RS_03Type05C,
@RS_03Type05S,
@RS_CreateId,
@RS_ModDate,
@RS_ModId,
@RS_Status
) ";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@RS_Guid", RS_Guid);
        oCmd.Parameters.AddWithValue("@RS_PorjectGuid", RS_PorjectGuid);
        oCmd.Parameters.AddWithValue("@RS_Stage", RS_Stage);
        oCmd.Parameters.AddWithValue("@RS_Year", RS_Year);
        oCmd.Parameters.AddWithValue("@RS_Season", RS_Season);
        oCmd.Parameters.AddWithValue("@RS_StartDay", RS_StartDay);
        oCmd.Parameters.AddWithValue("@RS_EndDay", RS_EndDay);
        oCmd.Parameters.AddWithValue("@RS_TotalMonth", RS_TotalMonth);
        oCmd.Parameters.AddWithValue("@RS_CostDesc", RS_CostDesc);
        oCmd.Parameters.AddWithValue("@RS_Type01Money", RS_Type01Money);
        oCmd.Parameters.AddWithValue("@RS_Type01Real", RS_Type01Real);
        oCmd.Parameters.AddWithValue("@RS_Type01RealRate", RS_Type01RealRate);
        oCmd.Parameters.AddWithValue("@RS_Type02Money", RS_Type02Money);
        oCmd.Parameters.AddWithValue("@RS_Type02Real", RS_Type02Real);
        oCmd.Parameters.AddWithValue("@RS_Type02RealRate", RS_Type02RealRate);
        oCmd.Parameters.AddWithValue("@RS_Type03Money", RS_Type03Money);
        oCmd.Parameters.AddWithValue("@RS_Type03Real", RS_Type03Real);
        oCmd.Parameters.AddWithValue("@RS_Type03RealRate", RS_Type03RealRate);
        oCmd.Parameters.AddWithValue("@RS_Type04Money", RS_Type04Money);
        oCmd.Parameters.AddWithValue("@RS_Type04Real", RS_Type04Real);
        oCmd.Parameters.AddWithValue("@RS_Type04RealRate", RS_Type04RealRate);
        oCmd.Parameters.AddWithValue("@RS_AllSchedule", RS_AllSchedule);
        oCmd.Parameters.AddWithValue("@RS_AllRealSchedule", RS_AllRealSchedule);
        oCmd.Parameters.AddWithValue("@RS_01Schedule", RS_01Schedule);
        oCmd.Parameters.AddWithValue("@RS_01RealSchedule", RS_01RealSchedule);
        oCmd.Parameters.AddWithValue("@RS_02Schedule", RS_02Schedule);
        oCmd.Parameters.AddWithValue("@RS_02RealSchedule", RS_02RealSchedule);
        oCmd.Parameters.AddWithValue("@RS_03Schedule", RS_03Schedule);
        oCmd.Parameters.AddWithValue("@RS_03RealSchedule", RS_03RealSchedule);
        oCmd.Parameters.AddWithValue("@RS_04Schedule", RS_04Schedule);
        oCmd.Parameters.AddWithValue("@RS_04RealSchedule", RS_04RealSchedule);
        oCmd.Parameters.AddWithValue("@RS_CheckPointData", RS_CheckPointData);
        oCmd.Parameters.AddWithValue("@RS_PushItemDesc", RS_PushItemDesc);
        oCmd.Parameters.AddWithValue("@RS_03Type01C", RS_03Type01C);
        oCmd.Parameters.AddWithValue("@RS_03Type01S", RS_03Type01S);
        oCmd.Parameters.AddWithValue("@RS_03Type02C", RS_03Type02C);
        oCmd.Parameters.AddWithValue("@RS_03Type02S", RS_03Type02S);
        oCmd.Parameters.AddWithValue("@RS_03Type03C", RS_03Type03C);
        oCmd.Parameters.AddWithValue("@RS_03Type03S", RS_03Type03S);
        oCmd.Parameters.AddWithValue("@RS_03Type04C", RS_03Type04C);
        oCmd.Parameters.AddWithValue("@RS_03Type04S", RS_03Type04S);
        oCmd.Parameters.AddWithValue("@RS_03Type05C", RS_03Type05C);
        oCmd.Parameters.AddWithValue("@RS_03Type05S", RS_03Type05S);
        oCmd.Parameters.AddWithValue("@RS_CreateId", RS_CreateId);
        oCmd.Parameters.AddWithValue("@RS_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@RS_ModId", RS_ModId);
        oCmd.Parameters.AddWithValue("@RS_Status", "A");

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }


    /// -----------------------------------------
    /// 建立者: Nick
    /// 功　能: 新增季報
    /// -----------------------------------------
    public void modSeason()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"update ReportSeason set
RS_PorjectGuid=@RS_PorjectGuid,
RS_Stage=@RS_Stage,
RS_Year=@RS_Year,
RS_Season=@RS_Season,
RS_StartDay=@RS_StartDay,
RS_EndDay=@RS_EndDay,
RS_TotalMonth=@RS_TotalMonth,
RS_CostDesc=@RS_CostDesc,
RS_Type01Money=@RS_Type01Money,
RS_Type01Real=@RS_Type01Real,
RS_Type01RealRate=@RS_Type01RealRate,
RS_Type02Money=@RS_Type02Money,
RS_Type02Real=@RS_Type02Real,
RS_Type02RealRate=@RS_Type02RealRate,
RS_Type03Money=@RS_Type03Money,
RS_Type03Real=@RS_Type03Real,
RS_Type03RealRate=@RS_Type03RealRate,
RS_Type04Money=@RS_Type04Money,
RS_Type04Real=@RS_Type04Real,
RS_Type04RealRate=@RS_Type04RealRate,
RS_AllSchedule=@RS_AllSchedule,
RS_AllRealSchedule=@RS_AllRealSchedule,
RS_01Schedule=@RS_01Schedule,
RS_01RealSchedule=@RS_01RealSchedule,
RS_02Schedule=@RS_02Schedule,
RS_02RealSchedule=@RS_02RealSchedule,
RS_03Schedule=@RS_03Schedule,
RS_03RealSchedule=@RS_03RealSchedule,
RS_04Schedule=@RS_04Schedule,
RS_04RealSchedule=@RS_04RealSchedule,
RS_CheckPointData=@RS_CheckPointData,
RS_PushItemDesc=@RS_PushItemDesc,
RS_03Type01C=@RS_03Type01C,
RS_03Type01S=@RS_03Type01S,
RS_03Type02C=@RS_03Type02C,
RS_03Type02S=@RS_03Type02S,
RS_03Type03C=@RS_03Type03C,
RS_03Type03S=@RS_03Type03S,
RS_03Type04C=@RS_03Type04C,
RS_03Type04S=@RS_03Type04S,
RS_03Type05C=@RS_03Type05C,
RS_03Type05S=@RS_03Type05S,
RS_ModDate=@RS_ModDate,
RS_ModId=@RS_ModId,
RS_Status=@RS_Status
where RS_Guid=@RS_Guid
";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@RS_Guid", RS_Guid);
        oCmd.Parameters.AddWithValue("@RS_PorjectGuid", RS_PorjectGuid);
        oCmd.Parameters.AddWithValue("@RS_Stage", RS_Stage);
        oCmd.Parameters.AddWithValue("@RS_Year", RS_Year);
        oCmd.Parameters.AddWithValue("@RS_Season", RS_Season);
        oCmd.Parameters.AddWithValue("@RS_StartDay", RS_StartDay);
        oCmd.Parameters.AddWithValue("@RS_EndDay", RS_EndDay);
        oCmd.Parameters.AddWithValue("@RS_TotalMonth", RS_TotalMonth);
        oCmd.Parameters.AddWithValue("@RS_CostDesc", RS_CostDesc);
        oCmd.Parameters.AddWithValue("@RS_Type01Money", RS_Type01Money);
        oCmd.Parameters.AddWithValue("@RS_Type01Real", RS_Type01Real);
        oCmd.Parameters.AddWithValue("@RS_Type01RealRate", RS_Type01RealRate);
        oCmd.Parameters.AddWithValue("@RS_Type02Money", RS_Type02Money);
        oCmd.Parameters.AddWithValue("@RS_Type02Real", RS_Type02Real);
        oCmd.Parameters.AddWithValue("@RS_Type02RealRate", RS_Type02RealRate);
        oCmd.Parameters.AddWithValue("@RS_Type03Money", RS_Type03Money);
        oCmd.Parameters.AddWithValue("@RS_Type03Real", RS_Type03Real);
        oCmd.Parameters.AddWithValue("@RS_Type03RealRate", RS_Type03RealRate);
        oCmd.Parameters.AddWithValue("@RS_Type04Money", RS_Type04Money);
        oCmd.Parameters.AddWithValue("@RS_Type04Real", RS_Type04Real);
        oCmd.Parameters.AddWithValue("@RS_Type04RealRate", RS_Type04RealRate);
        oCmd.Parameters.AddWithValue("@RS_AllSchedule", RS_AllSchedule);
        oCmd.Parameters.AddWithValue("@RS_AllRealSchedule", RS_AllRealSchedule);
        oCmd.Parameters.AddWithValue("@RS_01Schedule", RS_01Schedule);
        oCmd.Parameters.AddWithValue("@RS_01RealSchedule", RS_01RealSchedule);
        oCmd.Parameters.AddWithValue("@RS_02Schedule", RS_02Schedule);
        oCmd.Parameters.AddWithValue("@RS_02RealSchedule", RS_02RealSchedule);
        oCmd.Parameters.AddWithValue("@RS_03Schedule", RS_03Schedule);
        oCmd.Parameters.AddWithValue("@RS_03RealSchedule", RS_03RealSchedule);
        oCmd.Parameters.AddWithValue("@RS_04Schedule", RS_04Schedule);
        oCmd.Parameters.AddWithValue("@RS_04RealSchedule", RS_04RealSchedule);
        oCmd.Parameters.AddWithValue("@RS_CheckPointData", RS_CheckPointData);
        oCmd.Parameters.AddWithValue("@RS_PushItemDesc", RS_PushItemDesc);
        oCmd.Parameters.AddWithValue("@RS_03Type01C", RS_03Type01C);
        oCmd.Parameters.AddWithValue("@RS_03Type01S", RS_03Type01S);
        oCmd.Parameters.AddWithValue("@RS_03Type02C", RS_03Type02C);
        oCmd.Parameters.AddWithValue("@RS_03Type02S", RS_03Type02S);
        oCmd.Parameters.AddWithValue("@RS_03Type03C", RS_03Type03C);
        oCmd.Parameters.AddWithValue("@RS_03Type03S", RS_03Type03S);
        oCmd.Parameters.AddWithValue("@RS_03Type04C", RS_03Type04C);
        oCmd.Parameters.AddWithValue("@RS_03Type04S", RS_03Type04S);
        oCmd.Parameters.AddWithValue("@RS_03Type05C", RS_03Type05C);
        oCmd.Parameters.AddWithValue("@RS_03Type05S", RS_03Type05S);
        oCmd.Parameters.AddWithValue("@RS_CreateId", RS_CreateId);
        oCmd.Parameters.AddWithValue("@RS_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@RS_ModId", RS_ModId);
        oCmd.Parameters.AddWithValue("@RS_Status", "A");

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    /// -----------------------------------------
    /// 建立者: Nick
    /// 功　能: 查詢季報送審詳細資料
    /// -----------------------------------------
    public DataTable getSeasonReview()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT * FROM ReportCheck where RC_ReportGuid=@RS_Guid and RC_ReportType='02' and RC_Status='A' ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@RS_Guid", RS_Guid);
        oda.Fill(ds);
        return ds;
    }

    /// -----------------------------------------
    /// 建立者: 王晨俊
    /// 功　能: 歷史季報詳細資料
    /// -----------------------------------------
    public DataTable getHistorySeason(string RS_ID)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
declare @RS_Guid nvarchar(50)
select @RS_Guid = RS_Guid from ReportSeason where RS_ID=@RS_ID

SELECT top 1 b.RC_CheckType,b.RC_Status,e.C_Item_cn as CityName,c.M_Office,c.M_Name,c.M_Tel
,case when b.RC_Boss is null then f.M_Name else d.M_Name end as BossName
,b.RC_CheckDate,b.RC_CreateDate
,a.* 
FROM 
ReportSeason a
left join ReportCheck b on a.RS_Guid=b.RC_ReportGuid
left join Member c on b.RC_CreateId=c.M_Guid
left join Member d on b.RC_Boss=d.M_Guid
left join CodeTable e on c.M_City=e.C_Item and e.C_Group='02'
left join Member f on c.M_Manager_ID = f.M_Guid
where RC_ReportGuid=@RS_Guid
order by RC_CreateDate desc
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@RS_ID", RS_ID);
        oda.Fill(ds);
        return ds;
    }

    /// -----------------------------------------
    /// 建立者: Nick
    /// 功　能: 查詢季報匯出資料
    /// -----------------------------------------
    public DataTable getExportSeasonDetail()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select
member.M_Name as Member,
member.M_Tel as Tel,
rc.RC_CreateDate as WriteDate,
rc.RC_CheckDate as CheckDate,
city.C_Item_cn as City,
member.M_Office,
case when ISNULL(rc.RC_Boss,'')='' then MemberManager.M_Name else manager.M_Name end as Manager,
RS.*
from ReportSeason as rs
left join ReportCheck as rc on rc.RC_ReportGuid=rs.RS_Guid and rc.RC_Status!='D'
left join ProjectInfo as p on rs.RS_PorjectGuid=p.I_Guid
left join CodeTable as city on city.C_Group='02' and city.C_Item=p.I_City
left join Member as manager on manager.M_Guid=rc.RC_Boss
left join Member as member on member.M_Guid=rs.RS_CreateId
left join Member as MemberManager on MemberManager.M_Guid=member.M_Manager_ID
where RS_Status='A' and RS_ID=@RS_ID ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@RS_ID", RS_ID);
        oda.Fill(ds);
        return ds;
    }

    /// -----------------------------------------
    /// 建立者: 王晨峻
    /// 功　能: 查詢季報匯出資料
    /// -----------------------------------------
    public DataTable getICityByRSID(string RS_ID)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select top 1 I_City,RC_CheckType,RC_Status from ReportSeason
left join ProjectInfo on RS_PorjectGuid=I_Guid and I_Status='A' and I_Flag='Y'
left join ReportCheck on RS_Guid=RC_ReportGuid
where RS_ID=@RS_ID and RC_Status='A'
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();

        oCmd.Parameters.AddWithValue("@RS_ID", RS_ID);
        oda.Fill(dt);
        return dt;
    }

    //刪除季報草稿   20191225新增   
    public void delReportSeasonNotCheck()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);

        oCmd.CommandText = @"
            update ReportSeason set RS_Status='D' where RS_Guid=@RS_Guid

            update ReportCheck set RC_Status='D' where RC_ReportGuid=@RS_Guid and RC_Status='A' and RC_ReportType='02'
        ";

        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@RS_Guid", RS_Guid);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }
}