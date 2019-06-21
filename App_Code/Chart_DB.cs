using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;

/// <summary>
/// Chart 的摘要描述
/// </summary>
public class Chart_DB
{
    #region 私用
    string strCity = string.Empty;
    string strStage = string.Empty;
    string strLType = string.Empty;
    string strExType = string.Empty;
    #endregion
    #region 公用
    public string _strCity
    {
        set { strCity = value; }
    }
    public string _strStage
    {
        set { strStage = value; }
    }
    public string _strLType
    {
        set { strLType = value; }
    }
    public string _strExType
    {
        set { strExType = value; }
    }
    #endregion

    //地方圖表 設備汰換進度
    public DataSet getReDevice()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        #region 舊code 完成數是撈季報
        //sb.Append(@"
        //declare @City nvarchar(10)=@strCity--縣市代碼 06  all
        //declare @Stage nvarchar(10)=@strStage--期數
        //declare @PJGuid nvarchar(50)--計畫代號
        //declare @Cvalue nvarchar(100)--當期規劃數
        //if @City <>''
        // begin
        //        if @City<>'all'
        //            begin

        //                --根據機關代碼查出計畫GUID
        //          select @PJGuid=I_Guid 
        //                from Member left join ProjectInfo on M_City=@City and M_Guid=I_People and M_Status='A' and I_Status='A' and I_Flag='Y'
        //                where I_Guid is not null
        //                --根據計畫GUID & 期數 月報資料
        //                select RM_ProjectGuid,RM_Stage,RM_Year,RM_CPType,RM_Season,SUM(RM_Type2ValueSum) as RM_SUM,SUM(RM_Type3ValueSum) as RM_SUMKW into #tmpRM 
        //                from
        //                (
        //                 select RM_ProjectGuid,RM_Stage,RM_Year,RM_Month,RM_CPType,RM_Type2ValueSum,RM_Type3ValueSum,
        //                 case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //                 when '04' then '2' when '05' then '2' when '06' then '2'  
        //                 when '07' then '3' when '08' then '3' when '09' then '3'  
        //                 when '10' then '4' when '11' then '4' when '12' then '4'  
        //                 end as RM_Season,RC_CheckType
        //                 --
        //                 from ReportMonth 
        //                    left join ReportCheck on RM_ReportGuid=RC_ReportGuid and RC_Status='A'
        //                    where RM_Stage=@Stage and RM_ProjectGuid=@PJGuid and RM_Status='A'
        //                )#tmp
        //                where RC_CheckType='Y'
        //                group by RM_ProjectGuid,RM_Stage,RM_Year,RM_CPType,RM_Season
        //                --select * from #tmpRM
        //                ----根據計畫GUID & 期數 季報資料 無風管  SUM_01S='當期規劃數量' RM_SUMKW='申請數量' SUM_01C='完成數量'
        //                select RS_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
        //                ,Replace(Convert(Varchar(12),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
        //                from (
        //                 select RS_Stage,RS_Year,RS_Season,isnull(RM_SUMKW,0) as RM_SUM,isnull(RS_03Type01C,0) as SUM_C
        //                  ,case @Stage when '1' then (select I_Finish_item1_1 from ProjectInfo where I_Guid=@PJGuid)
        //                   when '2' then (select I_Finish_item1_2 from ProjectInfo where I_Guid=@PJGuid)
        //                   when '3' then (select I_Finish_item1_3 from ProjectInfo where I_Guid=@PJGuid)
        //                  else '0'
        //                  end as SUM_S
        //                 from ReportSeason
        //                 left join #tmpRM on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='01'
        //                 where RS_PorjectGuid=@PJGuid and RS_Stage=@Stage
        //                )#tmp
        //                ----根據計畫GUID & 期數 季報資料 老舊   '當期規劃數量'  RM_SUM='申請數量'  SUM_02C='完成數量'
        //                select RS_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
        //                ,Replace(Convert(Varchar(12),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
        //                from (
        //                 select RS_Stage,RS_Year,RS_Season,isnull(RM_SUM,0) as RM_SUM,isnull(RS_03Type02C,0) as SUM_C
        //                  ,case @Stage when '1' then (select I_Finish_item2_1 from ProjectInfo where I_Guid=@PJGuid)
        //                    when '2' then (select I_Finish_item2_2 from ProjectInfo where I_Guid=@PJGuid)
        //                    when '3' then (select I_Finish_item2_3 from ProjectInfo where I_Guid=@PJGuid)
        //                   else '0'
        //                   end as SUM_S
        //                 from ReportSeason
        //                 left join #tmpRM on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='02'
        //                 where RS_PorjectGuid=@PJGuid and RS_Stage=@Stage
        //                )#tmp
        //                ----根據計畫GUID & 期數 季報資料 室內停車場
        //                select RS_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
        //                ,Replace(Convert(Varchar(12),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
        //                from (
        //                 select RS_Stage,RS_Year,RS_Season,isnull(RM_SUM,0) as RM_SUM,isnull(RS_03Type03C,0) as SUM_C
        //                 ,case @Stage when '1' then (select I_Finish_item3_1 from ProjectInfo where I_Guid=@PJGuid)
        //                    when '2' then (select I_Finish_item3_2 from ProjectInfo where I_Guid=@PJGuid)
        //                    when '3' then (select I_Finish_item3_3 from ProjectInfo where I_Guid=@PJGuid)
        //                   else '0'
        //                   end as SUM_S
        //                 from ReportSeason
        //                 left join #tmpRM on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='03'
        //                 where RS_PorjectGuid=@PJGuid and RS_Stage=@Stage
        //                )#tmp
        //                ----根據計畫GUID & 期數 季報資料 中型
        //                select RS_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
        //                ,Replace(Convert(Varchar(12),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
        //                from (
        //                 select RS_Stage,RS_Year,RS_Season,isnull(RM_SUM,0) as RM_SUM,isnull(RS_03Type04C,0) as SUM_C
        //                 ,case @Stage when '1' then (select I_Finish_item4_1 from ProjectInfo where I_Guid=@PJGuid)
        //                    when '2' then (select I_Finish_item4_2 from ProjectInfo where I_Guid=@PJGuid)
        //                    when '3' then (select I_Finish_item4_3 from ProjectInfo where I_Guid=@PJGuid)
        //                   else '0'
        //                   end as SUM_S
        //                 from ReportSeason
        //                 left join #tmpRM on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='04'
        //                 where RS_PorjectGuid=@PJGuid and RS_Stage=@Stage
        //                )#tmp
        //                ----根據計畫GUID & 期數 季報資料 大型
        //                select RS_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
        //                ,Replace(Convert(Varchar(12),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
        //                from (
        //                 select RS_Stage,RS_Year,RS_Season,isnull(RM_SUM,0) as RM_SUM,isnull(RS_03Type05C,0) as SUM_C
        //                 ,case @Stage when '1' then (select I_Finish_item5_1 from ProjectInfo where I_Guid=@PJGuid)
        //                    when '2' then (select I_Finish_item5_2 from ProjectInfo where I_Guid=@PJGuid)
        //                    when '3' then (select I_Finish_item5_3 from ProjectInfo where I_Guid=@PJGuid)
        //                   else '0'
        //                   end as SUM_S
        //                 from ReportSeason
        //                 left join #tmpRM on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='05'
        //                 where RS_PorjectGuid=@PJGuid and RS_Stage=@Stage
        //                )#tmp

        //                drop table #tmpRM
        //            end
        //        else
        //            begin
        //                --根據期數 月報資料 指撈有審核過的
        //                select RM_ProjectGuid,RM_Stage,RM_Year,RM_CPType,RM_Season,SUM(RM_Type2ValueSum) as RM_SUM,SUM(RM_Type3ValueSum) as RM_SUMKW into #tmpRMall
        //                from
        //                (
        //                 select RM_ProjectGuid,RM_Stage,RM_Year,RM_Month,RM_CPType,RM_Type2ValueSum,RM_Type3ValueSum,
        //                 case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //                 when '04' then '2' when '05' then '2' when '06' then '2'  
        //                 when '07' then '3' when '08' then '3' when '09' then '3'  
        //                 when '10' then '4' when '11' then '4' when '12' then '4'  
        //                 end as RM_Season,RC_CheckType
        //                 --
        //                 from ReportMonth 
        //                 left join ReportCheck on RM_ReportGuid=RC_ReportGuid and RC_Status='A'
        //                 where RM_Stage=@Stage and RM_Status='A'
        //                )#tmp
        //                where RC_CheckType='Y'
        //                group by RM_ProjectGuid,RM_Stage,RM_Year,RM_CPType,RM_Season

        //                --根據期數 季報資料 無風管  SUM_01S='當期規劃數量' RM_SUMKW='申請數量' SUM_01C='完成數量'
        //                select RS_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
        //                ,Replace(Convert(Varchar(12),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
        //                from (
        //                 select RS_Stage,RS_Year,RS_Season,SUM(RM_SUM) as RM_SUM,SUM(SUM_C) as SUM_C,SUM_S
        //                 from ( 
        //                 select RS_Stage,RS_Year,RS_Season,isnull(RM_SUMKW,0) as RM_SUM,isnull(RS_03Type01C,0) as SUM_C
        //                  ,case @Stage 
        //                   when '1' then (select SUM(isnull(I_Finish_item1_1,0)) as I_Finish_item1_1 from ProjectInfo)
        //                   when '2' then (select SUM(isnull(I_Finish_item1_2,0)) as I_Finish_item1_2 from ProjectInfo)
        //                   when '3' then (select SUM(isnull(I_Finish_item1_3,0)) as I_Finish_item1_3 from ProjectInfo)
        //                  else '0'
        //                  end as SUM_S
        //                 from ReportSeason
        //                 left join #tmpRMall on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='01' and RM_ProjectGuid=RS_PorjectGuid
        //                 where RS_Stage=@Stage
        //                 )#tmp
        //                 group by RS_Stage,RS_Year,RS_Season,SUM_S
        //                )#tmp2

        //                --根據計畫GUID & 期數 季報資料 老舊   '當期規劃數量'  RM_SUM='申請數量'  SUM_02C='完成數量'
        //                select RS_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
        //                ,Replace(Convert(Varchar(12),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
        //                from (
        //                 select RS_Stage,RS_Year,RS_Season,SUM(RM_SUM) as RM_SUM,SUM(SUM_C) as SUM_C
        //                  ,case @Stage 
        //                  when '1' then (select SUM(isnull(I_Finish_item2_1,0)) as I_Finish_item2_1 from ProjectInfo)
        //                  when '2' then (select SUM(isnull(I_Finish_item2_2,0)) as I_Finish_item2_2 from ProjectInfo)
        //                  when '3' then (select SUM(isnull(I_Finish_item2_3,0)) as I_Finish_item2_3 from ProjectInfo)
        //                  else '0'
        //                  end as SUM_S
        //                 from ( 
        //                  select RS_Stage,RS_Year,RS_Season,isnull(RM_SUM,0) as RM_SUM,isnull(RS_03Type02C,0) as SUM_C
        //                  from ReportSeason
        //                  left join #tmpRMall on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='02'
        //                  where RS_Stage=@Stage
        //                 )#tmp
        //                 group by RS_Stage,RS_Year,RS_Season
        //                )#tmp2
        //                --根據計畫GUID & 期數 季報資料 室內停車場
        //                select RS_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
        //                ,Replace(Convert(Varchar(12),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
        //                from (
        //                 select RS_Stage,RS_Year,RS_Season,SUM(RM_SUM) as RM_SUM,SUM(SUM_C) as SUM_C
        //                  ,case @Stage 
        //                   when '1' then (select SUM(isnull(I_Finish_item3_1,0)) as I_Finish_item2_1 from ProjectInfo)
        //                   when '2' then (select SUM(isnull(I_Finish_item3_2,0)) as I_Finish_item2_2 from ProjectInfo)
        //                   when '3' then (select SUM(isnull(I_Finish_item3_3,0)) as I_Finish_item2_3 from ProjectInfo)
        //                   else '0'
        //                  end as SUM_S
        //                 from ( 
        //                  select RS_Stage,RS_Year,RS_Season,isnull(RM_SUM,0) as RM_SUM,isnull(RS_03Type03C,0) as SUM_C
        //                  from ReportSeason
        //                  left join #tmpRMall on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='03'
        //                  where RS_Stage=@Stage
        //                 )#tmp
        //                 group by RS_Stage,RS_Year,RS_Season
        //                )#tmp2
        //                --根據計畫GUID & 期數 季報資料 中型
        //                select RS_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
        //                ,Replace(Convert(Varchar(12),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
        //                from (
        //                 select RS_Stage,RS_Year,RS_Season,SUM(RM_SUM) as RM_SUM,SUM(SUM_C) as SUM_C
        //                  ,case @Stage 
        //                   when '1' then (select SUM(isnull(I_Finish_item4_1,0)) as I_Finish_item4_1 from ProjectInfo)
        //                   when '2' then (select SUM(isnull(I_Finish_item4_2,0)) as I_Finish_item4_2 from ProjectInfo)
        //                   when '3' then (select SUM(isnull(I_Finish_item4_3,0)) as I_Finish_item4_3 from ProjectInfo)
        //                   else '0'
        //                  end as SUM_S
        //                 from ( 
        //                  select RS_Stage,RS_Year,RS_Season,isnull(RM_SUM,0) as RM_SUM,isnull(RS_03Type04C,0) as SUM_C
        //                  from ReportSeason
        //                  left join #tmpRMall on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='04'
        //                  where RS_Stage=@Stage
        //                 )#tmp
        //                 group by RS_Stage,RS_Year,RS_Season
        //                )#tmp2
        //                --根據計畫GUID & 期數 季報資料 大型
        //                select RS_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
        //                ,Replace(Convert(Varchar(12),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
        //                ,Replace(Convert(Varchar(12),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
        //                from (
        //                 select RS_Stage,RS_Year,RS_Season,SUM(RM_SUM) as RM_SUM,SUM(SUM_C) as SUM_C
        //                  ,case @Stage 
        //                   when '1' then (select SUM(isnull(I_Finish_item5_1,0)) as I_Finish_item5_1 from ProjectInfo)
        //                   when '2' then (select SUM(isnull(I_Finish_item5_2,0)) as I_Finish_item5_2 from ProjectInfo)
        //                   when '3' then (select SUM(isnull(I_Finish_item5_3,0)) as I_Finish_item5_3 from ProjectInfo)
        //                  else '0'
        //                  end as SUM_S
        //                 from ( 
        //                  select RS_Stage,RS_Year,RS_Season,isnull(RM_SUM,0) as RM_SUM,isnull(RS_03Type05C,0) as SUM_C
        //                  from ReportSeason
        //                  left join #tmpRMall on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='05'
        //                  where RS_Stage=@Stage
        //                 )#tmp
        //                 group by RS_Stage,RS_Year,RS_Season
        //                )#tmp2
        //                drop table #tmpRMall
        //            end

        // end
        //");
        #endregion

        #region 新code 全部撈月報
        sb.Append(@"
            declare @City nvarchar(10)=@strCity--縣市代碼 06  all
            declare @Stage nvarchar(10)=@strStage--期數
            declare @PJGuid nvarchar(50)--計畫代號
            declare @Cvalue nvarchar(100)--當期規劃數
            if @City <>''
            begin
                if @City<>'all'
                    begin

                        --根據機關代碼查出計畫GUID
		                select @PJGuid=I_Guid 
                        from Member left join ProjectInfo on M_City=@City and M_Guid=I_People and M_Status='A' and I_Status='A' and I_Flag='Y'
                        where I_Guid is not null
                        --根據計畫GUID & 期數 月報資料
                        select RM_ProjectGuid,RM_Stage,RM_Year as RS_Year,RM_CPType,RM_Season as RS_Season,SUM(RM_Type1ValueSum) as RM_SUM1,SUM(RM_Type2ValueSum) as RM_SUM2,SUM(RM_Type3ValueSum) as RM_SUM3,SUM(RM_Type4ValueSum) as RM_SUM4 
                        into #tmpRM 
                        from
                        (
	                        select RM_ProjectGuid,RM_Stage,RM_Year,RM_Month,RM_CPType,RM_Type2ValueSum,RM_Type3ValueSum,RM_Type1ValueSum,RM_Type4ValueSum,
	                        case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
	                        when '04' then '2' when '05' then '2' when '06' then '2'  
	                        when '07' then '3' when '08' then '3' when '09' then '3'  
	                        when '10' then '4' when '11' then '4' when '12' then '4'  
	                        end as RM_Season,RC_CheckType
	                        --
	                        from ReportMonth 
                            left join ReportCheck on RM_ReportGuid=RC_ReportGuid and RC_Status='A'
                            where RM_Stage=@Stage and RM_ProjectGuid=@PJGuid and RM_Status='A' and RM_ReportType='01'
                        )#tmp
                        where RC_CheckType='Y'
                        group by RM_ProjectGuid,RM_Stage,RM_Year,RM_CPType,RM_Season
			            --select * from #tmpRM

                        ----select * from #tmpRM
                        ------根據計畫GUID & 期數 季報資料 無風管  SUM_01S='當期規劃數量' RM_SUMKW='申請數量' SUM_01C='完成數量'
			            select RM_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
			            ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
			            from (
				            select RM_Stage,RS_Year,RS_Season,isnull(RM_SUM3,0) as RM_SUM,isnull(RM_SUM4,0) as SUM_C
				            ,case @Stage when '1' then (select isnull(I_Finish_item1_1,0) from ProjectInfo where I_Guid=@PJGuid)
						            when '2' then (select isnull(I_Finish_item1_2,0) from ProjectInfo where I_Guid=@PJGuid)
						            when '3' then (select isnull(I_Finish_item1_3,0) from ProjectInfo where I_Guid=@PJGuid)
					            else '0'
					            end as SUM_S      
				            from #tmpRM
				            where RM_CPType='01'
			            )#tmp
                        ------根據計畫GUID & 期數 季報資料 老舊   '當期規劃數量'  RM_SUM='申請數量'  SUM_02C='完成數量'
			            select RM_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
			            ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
			            from (
				            select RM_Stage,RS_Year,RS_Season,isnull(RM_SUM1,0) as RM_SUM,isnull(RM_SUM2,0) as SUM_C
				            ,case @Stage when '1' then (select isnull(I_Finish_item2_1,0) from ProjectInfo where I_Guid=@PJGuid)
						            when '2' then (select isnull(I_Finish_item2_2,0) from ProjectInfo where I_Guid=@PJGuid)
						            when '3' then (select isnull(I_Finish_item2_3,0) from ProjectInfo where I_Guid=@PJGuid)
					            else '0'
					            end as SUM_S      
				            from #tmpRM
				            where RM_CPType='02'
			            )#tmp
                        ------根據計畫GUID & 期數 季報資料 室內停車場
			            select RM_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
			            ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
			            from (
				            select RM_Stage,RS_Year,RS_Season,isnull(RM_SUM1,0) as RM_SUM,isnull(RM_SUM2,0) as SUM_C
				            ,case @Stage when '1' then (select isnull(I_Finish_item3_1,0) from ProjectInfo where I_Guid=@PJGuid)
						            when '2' then (select isnull(I_Finish_item3_2,0) from ProjectInfo where I_Guid=@PJGuid)
						            when '3' then (select isnull(I_Finish_item3_3,0) from ProjectInfo where I_Guid=@PJGuid)
					            else '0'
					            end as SUM_S      
				            from #tmpRM
				            where RM_CPType='03'
			            )#tmp
                        ------根據計畫GUID & 期數 季報資料 中型
			            select RM_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
			            ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
			            from (
				            select RM_Stage,RS_Year,RS_Season,isnull(RM_SUM1,0) as RM_SUM,isnull(RM_SUM3,0) as SUM_C
				            ,case @Stage when '1' then (select isnull(I_Finish_item4_1,0) from ProjectInfo where I_Guid=@PJGuid)
						            when '2' then (select isnull(I_Finish_item4_2,0) from ProjectInfo where I_Guid=@PJGuid)
						            when '3' then (select isnull(I_Finish_item4_3,0) from ProjectInfo where I_Guid=@PJGuid)
					            else '0'
					            end as SUM_S      
				            from #tmpRM
				            where RM_CPType='04'
			            )#tmp
                        ------根據計畫GUID & 期數 季報資料 大型
			            select RM_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
			            ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
			            from (
				            select RM_Stage,RS_Year,RS_Season,isnull(RM_SUM1,0) as RM_SUM,isnull(RM_SUM3,0) as SUM_C
				            ,case @Stage when '1' then (select isnull(I_Finish_item5_1,0) from ProjectInfo where I_Guid=@PJGuid)
						            when '2' then (select isnull(I_Finish_item5_2,0) from ProjectInfo where I_Guid=@PJGuid)
						            when '3' then (select isnull(I_Finish_item5_3,0) from ProjectInfo where I_Guid=@PJGuid)
					            else '0'
					            end as SUM_S      
				            from #tmpRM
				            where RM_CPType='05'
			            )#tmp 
                        drop table #tmpRM
                    end
                else
                    begin
                        --根據期數 月報資料 指撈有審核過的
                        select RM_ProjectGuid,RM_Stage,RM_Year as RS_Year,RM_CPType,RM_Season as RS_Season,SUM(RM_Type1ValueSum) as RM_SUM1,SUM(RM_Type2ValueSum) as RM_SUM2,SUM(RM_Type3ValueSum) as RM_SUM3,SUM(RM_Type3ValueSum) as RM_SUM4 
                        into #tmpRMall
                        from
                        (
	                        select RM_ProjectGuid,RM_Stage,RM_Year,RM_Month,RM_CPType,RM_Type1ValueSum,RM_Type2ValueSum,RM_Type3ValueSum,RM_Type4ValueSum,
	                        case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
	                        when '04' then '2' when '05' then '2' when '06' then '2'  
	                        when '07' then '3' when '08' then '3' when '09' then '3'  
	                        when '10' then '4' when '11' then '4' when '12' then '4'  
	                        end as RM_Season,RC_CheckType
	                        --
	                        from ReportMonth 
	                        left join ReportCheck on RM_ReportGuid=RC_ReportGuid and RC_Status='A'
	                        where RM_Stage=@Stage and RM_Status='A' and RM_ReportType='01'
                        )#tmp
                        where RC_CheckType='Y'
                        group by RM_ProjectGuid,RM_Stage,RM_Year,RM_CPType,RM_Season
			            --select * from #tmpRMall
                        ----根據期數 季報資料 無風管  SUM_01S='當期規劃數量' RM_SUMKW='申請數量' SUM_01C='完成數量'
                        select RM_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
                        from (
	                        select RM_Stage,RS_Year,RS_Season,SUM(RM_SUM) as RM_SUM,SUM(SUM_C) as SUM_C,SUM(SUM_S) as SUM_S
	                        from ( 
					            select RM_Stage,RS_Year,RS_Season,isnull(RM_Sum3,0) as RM_SUM,isnull(RM_Sum4,0) as SUM_C
						            ,case @Stage 
							            when '1' then (select SUM(isnull(I_Finish_item1_1,0)) as I_Finish_item1_1 from ProjectInfo)
							            when '2' then (select SUM(isnull(I_Finish_item1_2,0)) as I_Finish_item1_2 from ProjectInfo)
							            when '3' then (select SUM(isnull(I_Finish_item1_3,0)) as I_Finish_item1_3 from ProjectInfo)
						            else '0'
						            end as SUM_S
					            from #tmpRMall
					            where RM_CPType='01'
	                        )#tmp
	                        group by RM_Stage,RS_Year,RS_Season
                        )#tmp2

                        ----根據計畫GUID & 期數 季報資料 老舊   '當期規劃數量'  RM_SUM='申請數量'  SUM_02C='完成數量'
			            select RM_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
                        from (
	                        select RM_Stage,RS_Year,RS_Season,SUM(RM_SUM) as RM_SUM,SUM(SUM_C) as SUM_C,SUM(SUM_S) as SUM_S
	                        from ( 
					            select RM_Stage,RS_Year,RS_Season,isnull(RM_Sum1,0) as RM_SUM,isnull(RM_Sum2,0) as SUM_C
						            ,case @Stage 
							            when '1' then (select SUM(isnull(I_Finish_item2_1,0)) as I_Finish_item1_1 from ProjectInfo)
							            when '2' then (select SUM(isnull(I_Finish_item2_2,0)) as I_Finish_item1_2 from ProjectInfo)
							            when '3' then (select SUM(isnull(I_Finish_item2_3,0)) as I_Finish_item1_3 from ProjectInfo)
						            else '0'
						            end as SUM_S
					            from #tmpRMall
					            where RM_CPType='02'
	                        )#tmp
	                        group by RM_Stage,RS_Year,RS_Season
                        )#tmp2
                        ----根據計畫GUID & 期數 季報資料 室內停車場
			            select RM_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
                        from (
	                        select RM_Stage,RS_Year,RS_Season,SUM(RM_SUM) as RM_SUM,SUM(SUM_C) as SUM_C,SUM(SUM_S) as SUM_S
	                        from ( 
					            select RM_Stage,RS_Year,RS_Season,isnull(RM_Sum1,0) as RM_SUM,isnull(RM_Sum2,0) as SUM_C
						            ,case @Stage 
							            when '1' then (select SUM(isnull(I_Finish_item3_1,0)) as I_Finish_item1_1 from ProjectInfo)
							            when '2' then (select SUM(isnull(I_Finish_item3_2,0)) as I_Finish_item1_2 from ProjectInfo)
							            when '3' then (select SUM(isnull(I_Finish_item3_3,0)) as I_Finish_item1_3 from ProjectInfo)
						            else '0'
						            end as SUM_S
					            from #tmpRMall
					            where RM_CPType='03'
	                        )#tmp
	                        group by RM_Stage,RS_Year,RS_Season
                        )#tmp2
                        ----根據計畫GUID & 期數 季報資料 中型
			            select RM_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
                        from (
	                        select RM_Stage,RS_Year,RS_Season,SUM(RM_SUM) as RM_SUM,SUM(SUM_C) as SUM_C,SUM(SUM_S) as SUM_S
	                        from ( 
					            select RM_Stage,RS_Year,RS_Season,isnull(RM_Sum1,0) as RM_SUM,isnull(RM_Sum3,0) as SUM_C
						            ,case @Stage 
							            when '1' then (select SUM(isnull(I_Finish_item4_1,0)) as I_Finish_item1_1 from ProjectInfo)
							            when '2' then (select SUM(isnull(I_Finish_item4_2,0)) as I_Finish_item1_2 from ProjectInfo)
							            when '3' then (select SUM(isnull(I_Finish_item4_3,0)) as I_Finish_item1_3 from ProjectInfo)
						            else '0'
						            end as SUM_S
					            from #tmpRMall
					            where RM_CPType='04'
	                        )#tmp
	                        group by RM_Stage,RS_Year,RS_Season
                        )#tmp2
                        ----根據計畫GUID & 期數 季報資料 大型
			            select RM_Stage,RS_Year,RS_Season,RM_SUM,SUM_C,SUM_S
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUM),1),'.00','') as RM_SUM_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_C),1),'.00','') as SUM_C_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
                        from (
	                        select RM_Stage,RS_Year,RS_Season,SUM(RM_SUM) as RM_SUM,SUM(SUM_C) as SUM_C,SUM(SUM_S) as SUM_S
	                        from ( 
					            select RM_Stage,RS_Year,RS_Season,isnull(RM_Sum1,0) as RM_SUM,isnull(RM_Sum3,0) as SUM_C
						            ,case @Stage 
							            when '1' then (select SUM(isnull(I_Finish_item5_1,0)) as I_Finish_item1_1 from ProjectInfo)
							            when '2' then (select SUM(isnull(I_Finish_item5_2,0)) as I_Finish_item1_2 from ProjectInfo)
							            when '3' then (select SUM(isnull(I_Finish_item5_3,0)) as I_Finish_item1_3 from ProjectInfo)
						            else '0'
						            end as SUM_S
					            from #tmpRMall
					            where RM_CPType='05'
	                        )#tmp
	                        group by RM_Stage,RS_Year,RS_Season
                        )#tmp2
                        drop table #tmpRMall
                    end
		        
            end
        ");
        #endregion
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@strCity", strCity);
        oCmd.Parameters.AddWithValue("@strStage", strStage);
        oda.Fill(ds);
        return ds;
    }

    //地方圖表 預期節電量
    public DataSet getExSave()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        #region 舊code 有撈季報資料
     //   sb.Append(@"
     //   declare @City nvarchar(10)=@strCity--縣市代碼 06  all
	    //declare @Stage nvarchar(10)=@strStage--期數
	    //declare @PJGuid nvarchar(50)--計畫代號
	    //if @City <>''
	    //begin
     //       if @City <>'all'
     //           begin
     //               --根據機關代碼查出計畫GUID
     //               select @PJGuid=I_Guid 
     //               from Member left join ProjectInfo on M_City=@City and M_Guid=I_People and M_Status='A' and I_Status='A' and I_Flag='Y'
     //               where I_Guid is not null
     //               --根據計畫GUID & 期數 月報資料
     //               select RM_Stage,RM_Year,RM_CPType,RM_Season,SUM(RM_PreVal) as RM_SUMPre,SUM(RM_ChkVal) as RM_SUMFinish into #tmpRM 
     //               from
     //               (
	    //                select RM_Stage,RM_Year,RM_Month,RM_CPType,RM_PreVal,RM_ChkVal,
	    //                case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
	    //                when '04' then '2' when '05' then '2' when '06' then '2'  
	    //                when '07' then '3' when '08' then '3' when '09' then '3'  
	    //                when '10' then '4' when '11' then '4' when '12' then '4'  
	    //                end as RM_Season,RC_CheckType
	    //                --
	    //                from ReportMonth 
     //                   left join ReportCheck on RM_ReportGuid=RC_ReportGuid and RC_Status='A'
     //                   where RM_Stage=@Stage and RM_ProjectGuid=@PJGuid and RM_Status='A'
     //               )#tmp
     //               where RC_CheckType='Y'
     //               group by RM_Stage,RM_Year,RM_CPType,RM_Season

     //               --無風管冷氣節電量：
     //               --=當期預計汰換kW/4kW(每台以4kW計算)*1245(節電度/台/年)
     //               --T8/T9
     //               --=每盞節電率舊燈每具96W，汰換為33W高效率燈具，且每年使用3000小時
     //               --單具節電=(96W-33W)*3000小時/1000=189kWh (度)
     //               --停車場
     //               --=每盞節電率舊燈每盞40W，汰換為高效率燈具，汰換為20W高效率燈具，且365天24小時點燈
     //               --單盞節電=20W*24(小時)*365(天)/1000=175.2度
     //               --中型能管系統
     //               --=每套*40,000(度/年)
     //               --大型能管系統
     //               --=每套*312,000(度/年)
     //               --select * from #tmpRM
     //               ----根據計畫GUID & 期數 季報資料 無風管  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'
     //               select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
     //               ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
     //               from (
	    //                select RS_Stage,RS_Year,RS_Season,convert(numeric(17,1),ROUND (isnull(RS_03Type01S,0)*1245/4,2)) as SUM_S,isnull(RM_SUMPre,0) as RM_SUMPre,isnull(RM_SUMFinish,0) as RM_SUMFinish
	    //                from ReportSeason
	    //                left join #tmpRM on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='01'
	    //                where RS_PorjectGuid=@PJGuid and RS_Stage=@Stage
     //               )#tmp
     //               ----根據計畫GUID & 期數 季報資料 老舊   SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'  63 = 96-33
     //               select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
     //               ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
     //               from (
	    //                select RS_Stage,RS_Year,RS_Season,convert(numeric(17,1),ROUND (isnull(RS_03Type02S,0)*189,2)) as SUM_S,isnull(RM_SUMPre,0) as RM_SUMPre,isnull(RM_SUMFinish,0) as RM_SUMFinish
	    //                from ReportSeason
	    //                left join #tmpRM on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='02'
	    //                where RS_PorjectGuid=@PJGuid and RS_Stage=@Stage
     //               )#tmp

     //               ----根據計畫GUID & 期數 季報資料 室內停車場  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'
     //               select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
     //               ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
     //               from (
	    //                select RS_Stage,RS_Year,RS_Season,convert(numeric(17,1),ROUND (isnull(RS_03Type03S,0)*175.2,2)) as SUM_S,isnull(RM_SUMPre,0) as RM_SUMPre,isnull(RM_SUMFinish,0) as RM_SUMFinish
	    //                from ReportSeason
	    //                left join #tmpRM on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='03'
	    //                where RS_PorjectGuid=@PJGuid and RS_Stage=@Stage
     //               )#tmp

     //               ----根據計畫GUID & 期數 季報資料 中型  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'
     //               select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
     //               ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
     //               from (
	    //                select RS_Stage,RS_Year,RS_Season,convert(bigint,isnull(RS_03Type04S,0))*40000 as SUM_S,isnull(RM_SUMPre,0) as RM_SUMPre,isnull(RM_SUMFinish,0) as RM_SUMFinish
	    //                from ReportSeason
	    //                left join #tmpRM on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='04'
	    //                where RS_PorjectGuid=@PJGuid and RS_Stage=@Stage
     //               )#tmp

     //               ----根據計畫GUID & 期數 季報資料 大型  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'
     //               select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
     //               ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
     //               from (
	    //                select RS_Stage,RS_Year,RS_Season,convert(bigint,isnull(RS_03Type05S,0))*312000 as SUM_S,isnull(RM_SUMPre,0) as RM_SUMPre,isnull(RM_SUMFinish,0) as RM_SUMFinish
	    //                from ReportSeason
	    //                left join #tmpRM on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='05'
	    //                where RS_PorjectGuid=@PJGuid and RS_Stage=@Stage
     //               )#tmp
		   //         drop table #tmpRM
     //           end
     //       else
     //           begin
     //               --根據期數 月報資料 只撈審核過的
     //               select RM_ProjectGuid,RM_Stage,RM_Year,RM_CPType,RM_Season,SUM(RM_PreVal) as RM_SUMPre,SUM(RM_ChkVal) as RM_SUMFinish into #tmpRMall
     //               from
     //               (
	    //                select RM_ProjectGuid,RM_Stage,RM_Year,RM_Month,RM_CPType,RM_PreVal,RM_ChkVal,
	    //                case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
	    //                when '04' then '2' when '05' then '2' when '06' then '2'  
	    //                when '07' then '3' when '08' then '3' when '09' then '3'  
	    //                when '10' then '4' when '11' then '4' when '12' then '4'  
	    //                end as RM_Season,RC_CheckType
	    //                --
	    //                from ReportMonth
	    //                left join ReportCheck on RM_ReportGuid=RC_ReportGuid and RC_Status='A'
     //                   where RM_Stage=@Stage and RM_Status='A'
     //               )#tmp
     //               where RC_CheckType='Y'
     //               group by RM_ProjectGuid,RM_Stage,RM_Year,RM_CPType,RM_Season

     //               --無風管冷氣節電量：
     //               --=當期預計汰換kW/4kW(每台以4kW計算)*1245(節電度/台/年)
     //               --T8/T9
     //               --=每盞節電率舊燈每具96W，汰換為33W高效率燈具，且每年使用3000小時
     //               --單具節電=(96W-33W)*3000小時/1000=189kWh (度)
     //               --停車場
     //               --=每盞節電率舊燈每盞40W，汰換為高效率燈具，汰換為20W高效率燈具，且365天24小時點燈
     //               --單盞節電=20W*24(小時)*365(天)/1000=175.2度
     //               --中型能管系統
     //               --=每套*40,000(度/年)
     //               --大型能管系統
     //               --=每套*312,000(度/年)
     //               --select * from #tmpRM
     //               ----根據計畫GUID & 期數 季報資料 無風管  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'
     //               select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
     //               ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
     //               from (
	    //                select RS_Stage,RS_Year,RS_Season,SUM(RM_SUMPre) as RM_SUMPre,SUM(RM_SUMFinish) as RM_SUMFinish
		   //                 ,case @Stage 
			  //                  when '1' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item1_1,0)*1245/4,2)) ) as I_Finish_item1_1 from ProjectInfo where I_Status='A')
			  //                  when '2' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item1_2,0)*1245/4,2)) ) as I_Finish_item1_2 from ProjectInfo where I_Status='A')
			  //                  when '3' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item1_3,0)*1245/4,2)) ) as I_Finish_item1_3 from ProjectInfo where I_Status='A')
		   //                 else '0'
		   //                 end as SUM_S
	    //                from
	    //                (
		   //                 select RS_Stage,RS_Year,RS_Season,isnull(RM_SUMPre,0) as RM_SUMPre,isnull(RM_SUMFinish,0) as RM_SUMFinish
		   //                 from ReportSeason
		   //                 left join #tmpRMall on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='01'
		   //                 where RS_Stage=@Stage and RS_Status='A'
	    //                )#tmp
	    //                group by RS_Stage,RS_Year,RS_Season
     //               )#tmp2
     //               ----根據期數 季報資料 老舊   SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'  63 = 96-33
     //               select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
     //               ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
     //               from (
	    //                select RS_Stage,RS_Year,RS_Season,SUM(RM_SUMPre) as RM_SUMPre,SUM(RM_SUMFinish) as RM_SUMFinish
		   //                 ,case @Stage 
			  //                  when '1' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item2_1,0)*189,2)) ) as I_Finish_item1_1 from ProjectInfo where I_Status='A')
			  //                  when '2' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item2_2,0)*189,2)) ) as I_Finish_item1_2 from ProjectInfo where I_Status='A')
			  //                  when '3' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item2_3,0)*189,2)) ) as I_Finish_item1_3 from ProjectInfo where I_Status='A')
		   //                 else '0'
		   //                 end as SUM_S
	    //                from
	    //                (
		   //                 select RS_Stage,RS_Year,RS_Season,isnull(RM_SUMPre,0) as RM_SUMPre,isnull(RM_SUMFinish,0) as RM_SUMFinish
		   //                 from ReportSeason
		   //                 left join #tmpRMall on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='02'
		   //                 where RS_Stage=@Stage and RS_Status='A'
	    //                )#tmp
	    //                group by RS_Stage,RS_Year,RS_Season
     //               )#tmp2
     //               ----根據期數 季報資料 室內停車場  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'
     //               select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
     //               ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
     //               from (
	    //                select RS_Stage,RS_Year,RS_Season,SUM(RM_SUMPre) as RM_SUMPre,SUM(RM_SUMFinish) as RM_SUMFinish
		   //                 ,case @Stage 
			  //                  when '1' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item3_1,0)*175.2,2)) ) as I_Finish_item1_1 from ProjectInfo where I_Status='A')
			  //                  when '2' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item3_2,0)*175.2,2)) ) as I_Finish_item1_2 from ProjectInfo where I_Status='A')
			  //                  when '3' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item3_3,0)*175.2,2)) ) as I_Finish_item1_3 from ProjectInfo where I_Status='A')
		   //                 else '0'
		   //                 end as SUM_S
	    //                from
	    //                (
		   //                 select RS_Stage,RS_Year,RS_Season,isnull(RM_SUMPre,0) as RM_SUMPre,isnull(RM_SUMFinish,0) as RM_SUMFinish
		   //                 from ReportSeason
		   //                 left join #tmpRMall on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='03'
		   //                 where RS_Stage=@Stage and RS_Status='A'
	    //                )#tmp
	    //                group by RS_Stage,RS_Year,RS_Season
     //               )#tmp2

     //               ----根據期數 季報資料 中型  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'
     //               select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
     //               ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
     //               from (
	    //                select RS_Stage,RS_Year,RS_Season,SUM(RM_SUMPre) as RM_SUMPre,SUM(RM_SUMFinish) as RM_SUMFinish
		   //                 ,case @Stage 
			  //                  when '1' then (select SUM(convert(bigint,isnull(I_Finish_item4_1,0))*40000 ) as I_Finish_item1_1 from ProjectInfo where I_Status='A')
			  //                  when '2' then (select SUM(convert(bigint,isnull(I_Finish_item4_2,0))*40000 ) as I_Finish_item1_2 from ProjectInfo where I_Status='A')
			  //                  when '3' then (select SUM(convert(bigint,isnull(I_Finish_item4_3,0))*40000 ) as I_Finish_item1_3 from ProjectInfo where I_Status='A')
		   //                 else '0'
		   //                 end as SUM_S
	    //                from
	    //                (
		   //                 select RS_Stage,RS_Year,RS_Season,isnull(RM_SUMPre,0) as RM_SUMPre,isnull(RM_SUMFinish,0) as RM_SUMFinish
		   //                 from ReportSeason
		   //                 left join #tmpRMall on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='04'
		   //                 where RS_Stage=@Stage and RS_Status='A'
	    //                )#tmp
	    //                group by RS_Stage,RS_Year,RS_Season
     //               )#tmp2

     //               ----根據期數 季報資料 大型  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'
     //               select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
     //               ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
     //               ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
     //               from (
	    //                select RS_Stage,RS_Year,RS_Season,SUM(RM_SUMPre) as RM_SUMPre,SUM(RM_SUMFinish) as RM_SUMFinish
		   //                 ,case @Stage 
			  //                  when '1' then (select SUM(convert(bigint,isnull(I_Finish_item5_1,0))*312000 ) as I_Finish_item1_1 from ProjectInfo where I_Status='A')
			  //                  when '2' then (select SUM(convert(bigint,isnull(I_Finish_item5_2,0))*312000 ) as I_Finish_item1_2 from ProjectInfo where I_Status='A')
			  //                  when '3' then (select SUM(convert(bigint,isnull(I_Finish_item5_3,0))*312000 ) as I_Finish_item1_3 from ProjectInfo where I_Status='A')
		   //                 else '0'
		   //                 end as SUM_S
	    //                from
	    //                (
		   //                 select RS_Stage,RS_Year,RS_Season,isnull(RM_SUMPre,0) as RM_SUMPre,isnull(RM_SUMFinish,0) as RM_SUMFinish
		   //                 from ReportSeason
		   //                 left join #tmpRMall on RS_Stage=RM_Stage and RS_Year=RM_Year and RS_Season=RM_Season and RM_CPType='05'
		   //                 where RS_Stage=@Stage and RS_Status='A'
	    //                )#tmp
	    //                group by RS_Stage,RS_Year,RS_Season
     //               )#tmp2

     //               drop table #tmpRMall
     //           end
		    
	    //end
     //   ");
        #endregion

        #region 新code 只撈月報資料
        sb.Append(@"
            declare @City nvarchar(10)=@strCity--縣市代碼 06  all
            declare @Stage nvarchar(10)=@strStage--期數
            declare @PJGuid nvarchar(50)--計畫代號
            if @City <>''
            begin
                if @City <>'all'
                    begin
                        --根據機關代碼查出計畫GUID
                        select @PJGuid=I_Guid 
                        from Member left join ProjectInfo on M_City=@City and M_Guid=I_People and M_Status='A' and I_Status='A' and I_Flag='Y'
                        where I_Guid is not null
                        --根據計畫GUID & 期數 月報資料
                        select RM_Stage as RS_Stage,RM_Year as RS_Year,RM_CPType,RM_Season as RS_Season
			            ,SUM(RM_Type1ValueSum) as RM_SUM1,SUM(RM_Type2ValueSum) as RM_SUM2,SUM(RM_Type3ValueSum) as RM_SUM3,SUM(RM_Type4ValueSum) as RM_SUM4 into #tmpRM 
                        from
                        (
	                        select RM_Stage,RM_Year,RM_Month,RM_CPType,RM_Type1ValueSum,RM_Type2ValueSum,RM_Type3ValueSum,RM_Type4ValueSum,
	                        case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
	                        when '04' then '2' when '05' then '2' when '06' then '2'  
	                        when '07' then '3' when '08' then '3' when '09' then '3'  
	                        when '10' then '4' when '11' then '4' when '12' then '4'  
	                        end as RM_Season,RC_CheckType
	                        --
	                        from ReportMonth 
                            left join ReportCheck on RM_ReportGuid=RC_ReportGuid and RC_Status='A'
                            where RM_Stage=@Stage and RM_ProjectGuid=@PJGuid and RM_Status='A' and RM_ReportType='01'
                        )#tmp
                        where RC_CheckType='Y'
                        group by RM_Stage,RM_Year,RM_CPType,RM_Season

                        --無風管冷氣節電量：
                        --=當期預計汰換kW/4kW(每台以4kW計算)*1245(節電度/台/年)
                        --T8/T9
                        --=每盞節電率舊燈每具96W，汰換為33W高效率燈具，且每年使用3000小時
                        --單具節電=(96W-33W)*3000小時/1000=189kWh (度)
                        --停車場
                        --=每盞節電率舊燈每盞40W，汰換為高效率燈具，汰換為20W高效率燈具，且365天24小時點燈
                        --單盞節電=20W*24(小時)*365(天)/1000=175.2度
                        --中型能管系統
                        --=每套*40,000(度/年)
                        --大型能管系統
                        --=每套*312,000(度/年)
                        --select * from #tmpRM
                        ----根據計畫GUID & 期數 月報資料 無風管  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'
                        select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
                        from (
	                        select RS_Stage,RS_Year,RS_Season
				            ,convert(numeric(17,1),ROUND (isnull(RM_SUM3,0)*1245/4,2)) as RM_SUMPre
				            ,convert(numeric(17,1),ROUND (isnull(RM_SUM4,0)*1245/4,2)) as RM_SUMFinish
				            ,case @Stage 
					            when '1' then (select convert(numeric(17,1),ROUND (isnull(I_Finish_item1_1,0)*1245/4,2)) as I_Finish_item1_1 from ProjectInfo where I_Guid=@PJGuid)
					            when '2' then (select convert(numeric(17,1),ROUND (isnull(I_Finish_item1_2,0)*1245/4,2)) as I_Finish_item1_2 from ProjectInfo where I_Guid=@PJGuid)
					            when '3' then (select convert(numeric(17,1),ROUND (isnull(I_Finish_item1_3,0)*1245/4,2)) as I_Finish_item1_3 from ProjectInfo where I_Guid=@PJGuid)
				            else '0'
				            end as SUM_S
	                        from #tmpRM
	                        where RM_CPType='01'
                        )#tmp
                        ------根據計畫GUID & 期數 月報資料 老舊   SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'  63 = 96-33
			            select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
                        from (
	                        select RS_Stage,RS_Year,RS_Season
				            ,convert(numeric(17,1),ROUND (isnull(RM_SUM1,0)*189,2)) as RM_SUMPre
				            ,convert(numeric(17,1),ROUND (isnull(RM_SUM2,0)*189,2)) as RM_SUMFinish
				            ,case @Stage 
					            when '1' then (select convert(numeric(17,1),ROUND (isnull(I_Finish_item2_1,0)*189,2)) as I_Finish_item1_1 from ProjectInfo where I_Guid=@PJGuid)
					            when '2' then (select convert(numeric(17,1),ROUND (isnull(I_Finish_item2_2,0)*189,2)) as I_Finish_item1_2 from ProjectInfo where I_Guid=@PJGuid)
					            when '3' then (select convert(numeric(17,1),ROUND (isnull(I_Finish_item2_3,0)*189,2)) as I_Finish_item1_3 from ProjectInfo where I_Guid=@PJGuid)
				            else '0'
				            end as SUM_S
	                        from #tmpRM
	                        where RM_CPType='02'
                        )#tmp
                        ------根據計畫GUID & 期數 月報資料 室內停車場  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'
			            select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
                        from (
	                        select RS_Stage,RS_Year,RS_Season
				            ,convert(numeric(17,1),ROUND (isnull(RM_SUM1,0)*175.2,2)) as RM_SUMPre
				            ,convert(numeric(17,1),ROUND (isnull(RM_SUM2,0)*175.2,2)) as RM_SUMFinish
				            ,case @Stage
					            when '1' then (select convert(numeric(17,1),ROUND (isnull(I_Finish_item3_1,0)*175.2,2)) as I_Finish_item1_1 from ProjectInfo where I_Guid=@PJGuid)
					            when '2' then (select convert(numeric(17,1),ROUND (isnull(I_Finish_item3_2,0)*175.2,2)) as I_Finish_item1_2 from ProjectInfo where I_Guid=@PJGuid)
					            when '3' then (select convert(numeric(17,1),ROUND (isnull(I_Finish_item3_3,0)*175.2,2)) as I_Finish_item1_3 from ProjectInfo where I_Guid=@PJGuid)
				            else '0'
				            end as SUM_S
	                        from #tmpRM
	                        where RM_CPType='03'
                        )#tmp
                        ------根據計畫GUID & 期數 月報資料 中型  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'
			            select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
                        from (
	                        select RS_Stage,RS_Year,RS_Season
				            ,convert(bigint,isnull(RM_SUM1,0))*40000 as RM_SUMPre
				            ,convert(bigint,isnull(RM_SUM3,0))*40000 as RM_SUMFinish
				            ,case @Stage
					            when '1' then (select convert(bigint,isnull(I_Finish_item4_1,0))*40000 as I_Finish_item1_1 from ProjectInfo where I_Guid=@PJGuid)
					            when '2' then (select convert(bigint,isnull(I_Finish_item4_2,0))*40000 as I_Finish_item1_2 from ProjectInfo where I_Guid=@PJGuid)
					            when '3' then (select convert(bigint,isnull(I_Finish_item4_3,0))*40000 as I_Finish_item1_3 from ProjectInfo where I_Guid=@PJGuid)
				            else '0'
				            end as SUM_S
	                        from #tmpRM
	                        where RM_CPType='04'
                        )#tmp
                        ------根據計畫GUID & 期數 月報資料 大型  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'
			            select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
                        from (
	                        select RS_Stage,RS_Year,RS_Season
				            ,convert(bigint,isnull(RM_SUM1,0))*312000 as RM_SUMPre
				            ,convert(bigint,isnull(RM_SUM3,0))*312000 as RM_SUMFinish
				            ,case @Stage
					            when '1' then (select convert(bigint,isnull(I_Finish_item5_1,0))*312000 as I_Finish_item1_1 from ProjectInfo where I_Guid=@PJGuid)
					            when '2' then (select convert(bigint,isnull(I_Finish_item5_2,0))*312000 as I_Finish_item1_2 from ProjectInfo where I_Guid=@PJGuid)
					            when '3' then (select convert(bigint,isnull(I_Finish_item5_3,0))*312000 as I_Finish_item1_3 from ProjectInfo where I_Guid=@PJGuid)
				            else '0'
				            end as SUM_S
	                        from #tmpRM
	                        where RM_CPType='05'
                        )#tmp
		                drop table #tmpRM
                    end
                else
                    begin
                        --根據期數 月報資料 只撈審核過的
                        select RM_ProjectGuid,RM_Stage as RS_Stage,RM_Year as RS_Year,RM_CPType,RM_Season as RS_Season
			            ,SUM(RM_Type1ValueSum) as RM_SUM1,SUM(RM_Type2ValueSum) as RM_SUM2,SUM(RM_Type3ValueSum) as RM_SUM3,SUM(RM_Type4ValueSum) as RM_SUM4
			            into #tmpRMall
			            --,SUM(RM_PreVal) as RM_SUMPre,SUM(RM_ChkVal) as RM_SUMFinish 
                        from
                        (
	                        select RM_ProjectGuid,RM_Stage,RM_Year,RM_Month,RM_CPType,RM_PreVal,RM_ChkVal
				            ,RM_Type1ValueSum,RM_Type2ValueSum,RM_Type3ValueSum,RM_Type4ValueSum,
	                        case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
	                        when '04' then '2' when '05' then '2' when '06' then '2'  
	                        when '07' then '3' when '08' then '3' when '09' then '3'  
	                        when '10' then '4' when '11' then '4' when '12' then '4'  
	                        end as RM_Season,RC_CheckType
	                        from ReportMonth
	                        left join ReportCheck on RM_ReportGuid=RC_ReportGuid and RC_Status='A'
                            where RM_Stage=@Stage and RM_Status='A' and RM_ReportType='01'
                        )#tmp
                        where RC_CheckType='Y'
                        group by RM_ProjectGuid,RM_Stage,RM_Year,RM_CPType,RM_Season

                        --無風管冷氣節電量：
                        --=當期預計汰換kW/4kW(每台以4kW計算)*1245(節電度/台/年)
                        --T8/T9
                        --=每盞節電率舊燈每具96W，汰換為33W高效率燈具，且每年使用3000小時
                        --單具節電=(96W-33W)*3000小時/1000=189kWh (度)
                        --停車場
                        --=每盞節電率舊燈每盞40W，汰換為高效率燈具，汰換為20W高效率燈具，且365天24小時點燈
                        --單盞節電=20W*24(小時)*365(天)/1000=175.2度
                        --中型能管系統
                        --=每套*40,000(度/年)
                        --大型能管系統
                        --=每套*312,000(度/年)
                        --select * from #tmpRM
                        ----根據 期數 月報資料 無風管  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量' 3.4 1.2 1.2 1.3 1.3
			            select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
                        from (
	                        select RS_Stage,RS_Year,RS_Season
				            ,SUM(convert(numeric(17,1),ROUND (isnull(RM_SUM3,0)*1245/4,2))) as RM_SUMPre
				            ,SUM(convert(numeric(17,1),ROUND (isnull(RM_SUM4,0)*1245/4,2))) as RM_SUMFinish
				            ,case @Stage 
					            when '1' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item1_1,0)*1245/4,2)) ) as I_Finish_item1_1 from ProjectInfo)
					            when '2' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item1_2,0)*1245/4,2)) ) as I_Finish_item1_2 from ProjectInfo)
					            when '3' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item1_3,0)*1245/4,2)) ) as I_Finish_item1_3 from ProjectInfo)
				            else '0'
				            end as SUM_S
	                        from #tmpRMall
	                        where RM_CPType='01'
				            group by RS_Stage,RS_Year,RS_Season
                        )#tmp

                        ------根據期數 月報資料 老舊   SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'  63 = 96-33
			            select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
                        from (
	                        select RS_Stage,RS_Year,RS_Season
				            ,SUM(convert(numeric(17,1),ROUND (isnull(RM_SUM1,0)*189,2))) as RM_SUMPre
				            ,SUM(convert(numeric(17,1),ROUND (isnull(RM_SUM2,0)*189,2))) as RM_SUMFinish
				            ,case @Stage 
					            when '1' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item2_1,0)*189,2)) ) as I_Finish_item1_1 from ProjectInfo)
					            when '2' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item2_2,0)*189,2)) ) as I_Finish_item1_2 from ProjectInfo)
					            when '3' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item2_3,0)*189,2)) ) as I_Finish_item1_3 from ProjectInfo)
				            else '0'
				            end as SUM_S
	                        from #tmpRMall
	                        where RM_CPType='02'
				            group by RS_Stage,RS_Year,RS_Season
                        )#tmp
            
                        ------根據期數 月報資料 室內停車場  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'
			            select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
                        from (
	                        select RS_Stage,RS_Year,RS_Season
				            ,SUM(convert(numeric(17,1),ROUND (isnull(RM_SUM1,0)*175.2,2))) as RM_SUMPre
				            ,SUM(convert(numeric(17,1),ROUND (isnull(RM_SUM2,0)*175.2,2))) as RM_SUMFinish
				            ,case @Stage 
					            when '1' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item3_1,0)*175.2,2)) ) as I_Finish_item1_1 from ProjectInfo)
					            when '2' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item3_2,0)*175.2,2)) ) as I_Finish_item1_2 from ProjectInfo)
					            when '3' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item3_3,0)*175.2,2)) ) as I_Finish_item1_3 from ProjectInfo)
				            else '0'
				            end as SUM_S
	                        from #tmpRMall
	                        where RM_CPType='03'
				            group by RS_Stage,RS_Year,RS_Season
                        )#tmp

                        ------根據期數 月報資料 中型  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'
			            select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
                        from (
	                        select RS_Stage,RS_Year,RS_Season
				            ,SUM(convert(numeric(17,1),ROUND (isnull(RM_SUM1,0)*40000,2))) as RM_SUMPre
				            ,SUM(convert(numeric(17,1),ROUND (isnull(RM_SUM3,0)*40000,2))) as RM_SUMFinish
				            ,case @Stage 
				            when '1' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item4_1,0)*40000,2)) ) as I_Finish_item1_1 from ProjectInfo)
				            when '2' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item4_2,0)*40000,2)) ) as I_Finish_item1_2 from ProjectInfo)
				            when '3' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item4_3,0)*40000,2)) ) as I_Finish_item1_3 from ProjectInfo)
				            else '0'
				            end as SUM_S
				            from #tmpRMall
	                        where RM_CPType='04'
				            group by RS_Stage,RS_Year,RS_Season
                        )#tmp

                        ------根據期數 月報資料 大型  SUM_S='當期規劃數量' RM_SUMPre='申請數量'  RM_SUMFinish='完成數量'
			            select RS_Stage,RS_Year,RS_Season,SUM_S,RM_SUMPre,RM_SUMFinish
                        ,Replace(Convert(Varchar(20),CONVERT(money,SUM_S),1),'.00','') as SUM_S_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMPre),1),'.00','') as RM_SUMPre_money
                        ,Replace(Convert(Varchar(20),CONVERT(money,RM_SUMFinish),1),'.00','') as RM_SUMFinish_money
                        from (
	                        select RS_Stage,RS_Year,RS_Season
				            ,SUM(convert(numeric(17,1),ROUND (isnull(RM_SUM1,0)*312000,2))) as RM_SUMPre
				            ,SUM(convert(numeric(17,1),ROUND (isnull(RM_SUM3,0)*312000,2))) as RM_SUMFinish
				            ,case @Stage 
					            when '1' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item5_1,0)*312000,2)) ) as I_Finish_item1_1 from ProjectInfo)
					            when '2' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item5_2,0)*312000,2)) ) as I_Finish_item1_2 from ProjectInfo)
					            when '3' then (select SUM(convert(numeric(17,1),ROUND (isnull(I_Finish_item5_3,0)*312000,2)) ) as I_Finish_item1_3 from ProjectInfo)
				            else '0'
				            end as SUM_S
	                        from #tmpRMall
	                        where RM_CPType='05'
				            group by RS_Stage,RS_Year,RS_Season
                        )#tmp

                        drop table #tmpRMall
                    end
		    
            end
        ");
        #endregion

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@strCity", strCity);
        oCmd.Parameters.AddWithValue("@strStage", strStage);
        oda.Fill(ds);
        return ds;
    }

    //地方圖表 申請單位分析
    public DataSet getUnitAnalyze()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
declare @City nvarchar(10)=@strCity--縣市代碼 06  all
declare @Stage nvarchar(10)=@strStage--期數
declare @PJGuid nvarchar(50)--計畫代號
declare @float_1 float;
declare @float_2 float;
declare @float_3 float;
if @City <>''
begin
    if @City <>'all'   
        begin        
		
            --根據機關代碼查出計畫GUID
	        select @PJGuid=I_Guid 
	        from Member left join ProjectInfo on M_City=@City and M_Guid=I_People and M_Status='A' and I_Status='A' and I_Flag='Y'
	        where I_Guid is not null
	        --根據計畫GUID & 期數 月報資料
	        select RM_Stage,RM_CPType,SUM(RM_Type1Value1) as RM_SUM1,SUM(RM_Type1Value2) as RM_SUM2,SUM(RM_Type1Value3) as RM_SUM3 into #tmpRM 
	        from
	        (
		        select RM_Stage,RM_Month,RM_CPType,RM_Type1Value1,RM_Type1Value2,RM_Type1Value3,RC_CheckType
		        --
		        from ReportMonth 
                left join ReportCheck on RM_ReportGuid=RC_ReportGuid and RC_Status='A'
                where RM_Stage=@Stage and RM_ProjectGuid=@PJGuid and RM_Status='A' and RM_ReportType='01'
	        )#tmp
            where RC_CheckType='Y'
	        group by RM_Stage,RM_CPType

	        --select * from #tmpRM
			--無風管
			select @float_1=convert(float,RM_SUM1),@float_2=convert(float,RM_SUM2),@float_3=convert(float,RM_SUM3) from #tmpRM where RM_CPType='01'
			select RM_Stage,RM_CPType,RM_SUM1,RM_SUM2,RM_SUM3,sumAvg,sumAvg2,sumAvg3
			,case RM_SUM1 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM1),1),'.00','') end as RM_SUM1_money
			,case RM_SUM2 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM2),1),'.00','') end as RM_SUM2_money
			,case RM_SUM3 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM3),1),'.00','') end as RM_SUM3_money
			from (
				select *
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_1/(@float_1+@float_2+@float_3),2)*100 end as sumAvg
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_2/(@float_1+@float_2+@float_3),2)*100 end as sumAvg2
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_3/(@float_1+@float_2+@float_3),2)*100 end as sumAvg3 
				from #tmpRM where RM_CPType='01' 
	        )#tmp
			--老舊
			select @float_1=convert(float,RM_SUM1),@float_2=convert(float,RM_SUM2),@float_3=convert(float,RM_SUM3) from #tmpRM where RM_CPType='02'
	        select RM_Stage,RM_CPType,RM_SUM1,RM_SUM2,RM_SUM3,sumAvg,sumAvg2,sumAvg3
			,case RM_SUM1 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM1),1),'.00','') end as RM_SUM1_money
			,case RM_SUM2 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM2),1),'.00','') end as RM_SUM2_money
			,case RM_SUM3 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM3),1),'.00','') end as RM_SUM3_money
			from (
				select *
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_1/(@float_1+@float_2+@float_3),2)*100 end as sumAvg 
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_2/(@float_1+@float_2+@float_3),2)*100 end as sumAvg2
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_3/(@float_1+@float_2+@float_3),2)*100 end as sumAvg3 
				from #tmpRM where RM_CPType='02' 
	        )#tmp

			--停車場
			select @float_1=convert(float,RM_SUM1),@float_2=convert(float,RM_SUM2),@float_3=convert(float,RM_SUM3) from #tmpRM where RM_CPType='03'
	        select RM_Stage,RM_CPType,RM_SUM1,RM_SUM2,RM_SUM3,sumAvg,sumAvg2,sumAvg3
			,case RM_SUM1 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM1),1),'.00','') end as RM_SUM1_money
			,case RM_SUM2 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM2),1),'.00','') end as RM_SUM2_money
			,case RM_SUM3 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM3),1),'.00','') end as RM_SUM3_money
			from (
				select *
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_1/(@float_1+@float_2+@float_3),2)*100 end as sumAvg 
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_2/(@float_1+@float_2+@float_3),2)*100 end as sumAvg2
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_3/(@float_1+@float_2+@float_3),2)*100 end as sumAvg3 
				from #tmpRM where RM_CPType='03' 
	        )#tmp

			--中型
			select @float_1=convert(float,RM_SUM1),@float_2=convert(float,RM_SUM2),@float_3=convert(float,RM_SUM3) from #tmpRM where RM_CPType='04'
	        select RM_Stage,RM_CPType,RM_SUM1,RM_SUM2,RM_SUM3,sumAvg,sumAvg2,sumAvg3
			,case RM_SUM1 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM1),1),'.00','') end as RM_SUM1_money
			,case RM_SUM2 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM2),1),'.00','') end as RM_SUM2_money
			,case RM_SUM3 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM3),1),'.00','') end as RM_SUM3_money
			from (
				select *
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_1/(@float_1+@float_2+@float_3),2)*100 end as sumAvg 
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_2/(@float_1+@float_2+@float_3),2)*100 end as sumAvg2
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_3/(@float_1+@float_2+@float_3),2)*100 end as sumAvg3 
				from #tmpRM where RM_CPType='04' 
	        )#tmp
			--大型
			select @float_1=convert(float,RM_SUM1),@float_2=convert(float,RM_SUM2),@float_3=convert(float,RM_SUM3) from #tmpRM where RM_CPType='05'
	        select RM_Stage,RM_CPType,RM_SUM1,RM_SUM2,RM_SUM3,sumAvg,sumAvg2,sumAvg3
			,case RM_SUM1 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM1),1),'.00','') end as RM_SUM1_money
			,case RM_SUM2 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM2),1),'.00','') end as RM_SUM2_money
			,case RM_SUM3 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM3),1),'.00','') end as RM_SUM3_money
			from (
				select *
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_1/(@float_1+@float_2+@float_3),2)*100 end as sumAvg 
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_2/(@float_1+@float_2+@float_3),2)*100 end as sumAvg2
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_3/(@float_1+@float_2+@float_3),2)*100 end as sumAvg3 
				from #tmpRM where RM_CPType='05' 
			)#tmp
	        drop table #tmpRM
        end
    else
        begin
            --根據期數 月報資料
            select RM_Stage,RM_CPType,SUM(RM_Type1Value1) as RM_SUM1,SUM(RM_Type1Value2) as RM_SUM2,SUM(RM_Type1Value3) as RM_SUM3 into #tmpRMall 
            from
            (
	            select RM_Stage,RM_Month,RM_CPType,RM_Type1Value1,RM_Type1Value2,RM_Type1Value3,RC_CheckType
	            --
	            from ReportMonth 
                left join ReportCheck on RM_ReportGuid=RC_ReportGuid and RC_Status='A'
                where RM_Stage=@Stage and RM_Status='A' and RM_ReportType='01'
            )#tmp
            where RC_CheckType='Y'
            group by RM_Stage,RM_CPType

			--無風管
            select @float_1=convert(float,RM_SUM1),@float_2=convert(float,RM_SUM2),@float_3=convert(float,RM_SUM3) from #tmpRMall where RM_CPType='01'
            select RM_Stage,RM_CPType,RM_SUM1,RM_SUM2,RM_SUM3,sumAvg,sumAvg2,sumAvg3
			,case RM_SUM1 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM1),1),'.00','') end as RM_SUM1_money
			,case RM_SUM2 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM2),1),'.00','') end as RM_SUM2_money
			,case RM_SUM3 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM3),1),'.00','') end as RM_SUM3_money
			from (
				select *
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_1/(@float_1+@float_2+@float_3),2)*100 end as sumAvg
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_2/(@float_1+@float_2+@float_3),2)*100 end as sumAvg2
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_3/(@float_1+@float_2+@float_3),2)*100 end as sumAvg3 
				from #tmpRMall where RM_CPType='01' 
			)#tmp

			--老舊
            select @float_1=convert(float,RM_SUM1),@float_2=convert(float,RM_SUM2),@float_3=convert(float,RM_SUM3) from #tmpRMall where RM_CPType='02'
            select RM_Stage,RM_CPType,RM_SUM1,RM_SUM2,RM_SUM3,sumAvg,sumAvg2,sumAvg3
			,case RM_SUM1 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM1),1),'.00','') end as RM_SUM1_money
			,case RM_SUM2 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM2),1),'.00','') end as RM_SUM2_money
			,case RM_SUM3 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM3),1),'.00','') end as RM_SUM3_money
			from (
				select *
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_1/(@float_1+@float_2+@float_3),2)*100 end as sumAvg 
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_2/(@float_1+@float_2+@float_3),2)*100 end as sumAvg2
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_3/(@float_1+@float_2+@float_3),2)*100 end as sumAvg3 
				from #tmpRMall where RM_CPType='02' 
			)#tmp

			--停車場
            select @float_1=convert(float,RM_SUM1),@float_2=convert(float,RM_SUM2),@float_3=convert(float,RM_SUM3) from #tmpRMall where RM_CPType='03'
            select RM_Stage,RM_CPType,RM_SUM1,RM_SUM2,RM_SUM3,sumAvg,sumAvg2,sumAvg3
			,case RM_SUM1 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM1),1),'.00','')end as RM_SUM1_money
			,case RM_SUM2 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM2),1),'.00','')end as RM_SUM2_money
			,case RM_SUM3 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM3),1),'.00','')end as RM_SUM3_money
			from (
				select *
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_1/(@float_1+@float_2+@float_3),2)*100 end as sumAvg 
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_2/(@float_1+@float_2+@float_3),2)*100 end as sumAvg2
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_3/(@float_1+@float_2+@float_3),2)*100 end as sumAvg3 
				from #tmpRMall where RM_CPType='03' 
			)#tmp

			--中型
            select @float_1=convert(float,RM_SUM1),@float_2=convert(float,RM_SUM2),@float_3=convert(float,RM_SUM3) from #tmpRMall where RM_CPType='04'
            select RM_Stage,RM_CPType,RM_SUM1,RM_SUM2,RM_SUM3,sumAvg,sumAvg2,sumAvg3
			,case RM_SUM1 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM1),1),'.00','') end as RM_SUM1_money
			,case RM_SUM2 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM2),1),'.00','') end as RM_SUM2_money
			,case RM_SUM3 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM3),1),'.00','') end as RM_SUM3_money
			from (
				select *
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_1/(@float_1+@float_2+@float_3),2)*100 end as sumAvg 
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_2/(@float_1+@float_2+@float_3),2)*100 end as sumAvg2
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_3/(@float_1+@float_2+@float_3),2)*100 end as sumAvg3 
				from #tmpRMall where RM_CPType='04' 
			)#tmp

			--大型
            select @float_1=convert(float,RM_SUM1),@float_2=convert(float,RM_SUM2),@float_3=convert(float,RM_SUM3) from #tmpRMall where RM_CPType='05'
            select RM_Stage,RM_CPType,RM_SUM1,RM_SUM2,RM_SUM3,sumAvg,sumAvg2,sumAvg3
			,case RM_SUM1 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM1),1),'.00','') end as RM_SUM1_money
			,case RM_SUM2 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM2),1),'.00','') end as RM_SUM2_money
			,case RM_SUM3 when '0' then '0' else Replace(Convert(Varchar(12),CONVERT(money,RM_SUM3),1),'.00','') end as RM_SUM3_money
			from (
				select *
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_1/(@float_1+@float_2+@float_3),2)*100 end as sumAvg 
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_2/(@float_1+@float_2+@float_3),2)*100 end as sumAvg2
				,case (@float_1+@float_2+@float_3) when 0 then '0' else ROUND(@float_3/(@float_1+@float_2+@float_3),2)*100 end as sumAvg3 
				from #tmpRMall where RM_CPType='05' 
			)#tmp
            drop table #tmpRMall
        end
	                
end
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@strCity", strCity);
        oCmd.Parameters.AddWithValue("@strStage", strStage);
        oda.Fill(ds);
        return ds;
    }

    //節電基礎及因地制宜工作進度摘要 報表
    public DataTable getReportApply()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        #region 舊CODE
        //sb.Append(@"
        //    create table #tmpall(
	       //     C_Item nvarchar(50),
	       //     C_Item_cn nvarchar(50),
	       //     I_Guid nvarchar(50),
	       //     RS_Stage nvarchar(2),
	       //     RS_year nvarchar(4),
	       //     RS_Season nvarchar(2),
	       //     RS_01Why nvarchar(max),
	       //     RS_01Summary nvarchar(max),
        //        RS_02Summary nvarchar(max),
	       //     rowsCount int
        //    )
        //    insert into #tmpall(C_Item,C_Item_cn,I_Guid,RS_Stage,RS_year,RS_Season,RS_01Why,RS_02Why,RS_03Why,rowsCount)
        //    select C_Item,C_Item_cn,I_Guid,RS_Stage,RS_year,RS_Season,RS_01Why,RS_02Why,RS_03Why,0
        //    from  CodeTable
        //    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //    left join ReportSeason on I_Guid=RS_PorjectGuid and RS_Stage=@strStage
        //    where C_Group='02'
        //    order by C_Item

        //    --這一期每個縣市有幾季的資料
        //    select C_Item,C_Item_cn,I_Guid,count(*) as rCount into #tmp
        //    from  CodeTable
        //    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //    left join ReportSeason on I_Guid=RS_PorjectGuid and RS_Stage=@strStage
        //    where C_Group='02'
        //    group by C_Item,C_Item_cn,I_Guid

        //    update #tmpall set #tmpall.rowsCount=#tmp.rCount
        //    from #tmp
        //    where #tmpall.C_Item=#tmp.C_Item

        //    select * from #tmpall

        //    drop table #tmp
        //    drop table #tmpall
        //");
        #endregion

        #region 201808季報改版後新code
        sb.Append(@"
create table #tmpall(
	C_Item nvarchar(50),
    C_Item_cn nvarchar(50),
    I_Guid nvarchar(50),
    RS_Stage nvarchar(2),
    RS_year nvarchar(4),
    RS_Season nvarchar(2),
    RS_Month nvarchar(2),
    RS_01Summary nvarchar(max),
    RS_02Summary nvarchar(max),
    rowsCount int
)
insert into #tmpall(C_Item,C_Item_cn,I_Guid,RS_Stage,RS_year,RS_Season,RS_Month,RS_01Summary,RS_02Summary,rowsCount)
select C_Item, C_Item_cn, I_Guid, RS_Stage, RS_year, RS_Season
,case RS_Season when '1' then '3' when '2' then '6' when '3' then '9' when '4' then '12' else '' end as RS_Month
,'','',0
from CodeTable
left join ProjectInfo on C_Item = I_City and I_Status = 'A' and I_Flag = 'Y'
left join ReportSeason on I_Guid = RS_PorjectGuid and RS_Stage = @strStage
left join ReportCheck on RS_Guid = RC_ReportGuid
where C_Group = '02' and RC_CheckType = 'Y' and RC_Status = 'A' and I_Flag = 'Y'
order by C_Item

--這一期每個縣市有幾季的資料
select C_Item,C_Item_cn,I_Guid,count(*) as rCount into #tmp
from  CodeTable
left join ProjectInfo on C_Item = I_City and I_Status = 'A'
left join ReportSeason on I_Guid = RS_PorjectGuid and RS_Stage = @strStage
where C_Group = '02'
group by C_Item,C_Item_cn,I_Guid

--從查核點撈出所有差異說明
--先撈出各機關定稿的計畫GUID
select I_City,I_Guid into #tmpPJ from CodeTable left join ProjectInfo on C_Item=I_City where C_Group='02'  and I_Status='A' and I_Flag='Y' order by C_Item asc
select* into #tmpwhile from #tmpall
declare @wh_C_Item nvarchar(3);
declare @wh_I_Guid nvarchar(50);
declare @wh_RS_Stage nvarchar(3);
declare @wh_RS_year nvarchar(4);
declare @wh_RS_Season nvarchar(3);
declare @wh_RS_Month nvarchar(3);
declare @wh_rowcount int= 0;
select @wh_rowcount = COUNT(*) from #tmpwhile;

while @wh_rowcount > 0
begin
    select @wh_rowcount = @wh_rowcount - 1;
    select top 1 @wh_C_Item = C_Item,@wh_I_Guid = I_Guid,@wh_RS_Stage = RS_Stage,@wh_RS_year = RS_year,@wh_RS_Season = RS_Season,@wh_RS_Month = RS_Month from #tmpwhile
		

    select P_Type,P_Period,CP_ReserveYear,CP_ReserveMonth,CP_Summary
	,case CP_ReserveMonth when '3' then '1' when '6' then '2' when '9' then '3' when '12' then '4' else '' end as RS_Season

    into #tmpwhinner from PushItem left join Check_Point on P_Guid=CP_ParentId and CP_Status='A'
	where P_ParentId = @wh_I_Guid and P_Period = @wh_RS_Stage and CP_ReserveYear = @wh_RS_year and CP_ReserveMonth = @wh_RS_Month

	SELECT m.P_Type,m.P_Period,m.CP_ReserveYear,m.CP_ReserveMonth,m.RS_Season ,case m.Summarys when '' then '' else  LEFT(m.Summarys, Len(m.Summarys) - 2) end as Summarys into #tmpupdate
	from
        (SELECT P_Type, P_Period, CP_ReserveYear, CP_ReserveMonth, RS_Season, (SELECT CP_Summary +'\n' from #tmpwhinner
	where P_Type = ord.P_Type and P_Period = ord.P_Period

    FOR XML PATH('')) as Summarys

    from #tmpwhinner ord
	) M
    group by m.P_Type,m.P_Period,m.Summarys,m.CP_ReserveYear,m.CP_ReserveMonth,m.RS_Season

    update #tmpall set RS_01Summary=a.Summarys
	from (

        select P_Type,P_Period,CP_ReserveYear,CP_ReserveMonth,RS_Season,Summarys from #tmpupdate where P_Type='01'
	)a
    where #tmpall.RS_Stage=a.P_Period and #tmpall.RS_year=a.CP_ReserveYear and #tmpall.RS_Season=a.RS_Season and #tmpall.C_Item=@wh_C_Item


    update #tmpall set RS_02Summary=a.Summarys
	from(
        select P_Type, P_Period, CP_ReserveYear, CP_ReserveMonth, RS_Season, Summarys from #tmpupdate where P_Type='02'
	)a
    where #tmpall.RS_Stage=a.P_Period and #tmpall.RS_year=a.CP_ReserveYear and #tmpall.RS_Season=a.RS_Season and #tmpall.C_Item=@wh_C_Item


    drop table #tmpwhinner
	drop table #tmpupdate
	delete from #tmpwhile where  C_Item= @wh_C_Item and I_Guid=@wh_I_Guid and RS_Stage=@wh_RS_Stage and RS_year=@wh_RS_year and RS_Season=@wh_RS_Season and RS_Month=@wh_RS_Month
end

update #tmpall set #tmpall.rowsCount=#tmp.rCount
from #tmp
where #tmpall.C_Item=#tmp.C_Item

select *from #tmpall  order by C_Item asc,RS_Year asc,RS_Season asc

drop table #tmp
drop table #tmpall
drop table #tmpPJ
drop table #tmpwhile
        ");
#endregion

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();
        
        oCmd.Parameters.AddWithValue("@strStage", strStage);
        oda.Fill(dt);
        return dt;
    }

    //計畫執行進度遭遇困難 報表
    public DataTable getReportTotalLog()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        #region 舊code
        //sb.Append(@"
        //    create table #tmpall(
        //     C_Item nvarchar(50),
        //     C_Item_cn nvarchar(50),
        //     I_Guid nvarchar(50),
        //     RS_Stage nvarchar(2),
        //     RS_year nvarchar(4),
        //     RS_Season nvarchar(2),
        //     RS_01Summary nvarchar(max),
        //     RS_02Summary nvarchar(max),
        //     rowsCount int
        //    )
        //    insert into #tmpall(C_Item,C_Item_cn,I_Guid,RS_Stage,RS_year,RS_Season,RS_01Summary,RS_02Summary,rowsCount)
        //    select C_Item,C_Item_cn,I_Guid,RS_Stage,RS_year,RS_Season,RS_01Summary,RS_02Summary,0
        //    from  CodeTable
        //    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //    left join ReportSeason on I_Guid=RS_PorjectGuid and RS_Stage='1'
        //    left join ReportCheck on RS_ReportGuid = RC_ReportGuid
        //    where C_Group='02' and RC_CheckType='Y' and RC_Status='A'
        //    order by C_Item

        //    --這一期每個縣市有幾季的資料
        //    select C_Item,C_Item_cn,I_Guid,count(*) as rCount into #tmp
        //    from(
        //    select C_Item,C_Item_cn,I_Guid,RC_CheckType,RC_Status
        //    from  CodeTable
        //    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //    left join ReportSeason on I_Guid=RS_PorjectGuid and RS_Stage='1'
        //    left join ReportCheck on RS_ReportGuid = RC_ReportGuid
        //    where C_Group='02' and RC_CheckType='Y' and RC_Status='A'
        //    )#tmp
        //    group by C_Item,C_Item_cn,I_Guid

        //    update #tmpall set #tmpall.rowsCount=#tmp.rCount
        //    from #tmp
        //    where #tmpall.C_Item=#tmp.C_Item

        //    select * from #tmpall

        //    drop table #tmp
        //    drop table #tmpall
        //");
        #endregion

        #region 201808季報改版後新code
        sb.Append(@"
 create table #tmpall(
	C_Item nvarchar(50),
	C_Item_cn nvarchar(50),
	I_Guid nvarchar(50),
	RS_Stage nvarchar(2),
	RS_year nvarchar(4),
	RS_Season nvarchar(2),
	RS_Month nvarchar(2),
	RS_01Why nvarchar(max),
	RS_02Why nvarchar(max),
	RS_03Why nvarchar(max),
    RS_ExWhy nvarchar(max),
	rowsCount int
)

insert into #tmpall(C_Item,C_Item_cn,I_Guid,RS_Stage,RS_year,RS_Season,RS_Month,RS_01Why,RS_02Why,RS_03Why,RS_ExWhy,rowsCount)
select C_Item,C_Item_cn,I_Guid,RS_Stage,RS_year,RS_Season
,case RS_Season when '1' then '3' when '2' then '6' when '3' then '9' when '4' then '12' else '' end as RS_Month
,'','','','',0
from  CodeTable
left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
left join ReportSeason on I_Guid=RS_PorjectGuid and RS_Stage=@strStage
left join ReportCheck on RS_Guid = RC_ReportGuid
where C_Group='02' and RC_CheckType='Y' and RC_Status='A' and I_Flag='Y'
order by C_Item

--這一期每個縣市有幾季的資料
select C_Item,C_Item_cn,I_Guid,count(*) as rCount into #tmp
from(
select C_Item,C_Item_cn,I_Guid,RC_CheckType,RC_Status
from  CodeTable
left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
left join ReportSeason on I_Guid=RS_PorjectGuid and RS_Stage=@strStage
left join ReportCheck on RS_Guid = RC_ReportGuid
where C_Group='02' and RC_CheckType='Y' and RC_Status='A'
)#tmp
group by C_Item,C_Item_cn,I_Guid

--從查核點撈出所有差異說明
--先撈出各機關定稿的計畫GUID
select I_City,I_Guid into #tmpPJ from CodeTable left join ProjectInfo on C_Item=I_City where C_Group='02'  and I_Status='A' and I_Flag='Y' order by C_Item asc
select * into #tmpwhile from #tmpall
declare @wh_C_Item nvarchar(3);
declare @wh_I_Guid nvarchar(50);
declare @wh_RS_Stage nvarchar(3);
declare @wh_RS_year nvarchar(4);
declare @wh_RS_Season nvarchar(3);
declare @wh_RS_Month nvarchar(3);
declare @wh_rowcount int=0;
select @wh_rowcount=COUNT(*) from #tmpwhile;

while @wh_rowcount>0
	begin
		select @wh_rowcount=@wh_rowcount-1;
		select top 1 @wh_C_Item = C_Item,@wh_I_Guid=I_Guid,@wh_RS_Stage=RS_Stage,@wh_RS_year=RS_year,@wh_RS_Season=RS_Season,@wh_RS_Month=RS_Month from #tmpwhile
		
		select P_Type,P_Period,CP_ReserveYear,CP_ReserveMonth,CP_BackwardDesc 
		,case CP_ReserveMonth when '3' then '1' when '6' then '2' when '9' then '3' when '12' then '4' else '' end as RS_Season
		into #tmpwhinner from PushItem left join Check_Point on P_Guid=CP_ParentId and CP_Status='A'
		where P_ParentId=@wh_I_Guid and P_Period=@wh_RS_Stage and CP_ReserveYear=@wh_RS_year and CP_ReserveMonth=@wh_RS_Month

		SELECT m.P_Type,m.P_Period,m.CP_ReserveYear,m.CP_ReserveMonth,m.RS_Season ,case m.BackwardDescs when '' then '' else  LEFT(m.BackwardDescs,Len(m.BackwardDescs)-2) end as BackwardDescs into #tmpupdate
		from 
		(SELECT P_Type,P_Period,CP_ReserveYear,CP_ReserveMonth,RS_Season,(SELECT CP_BackwardDesc + '\n' from #tmpwhinner
		where P_Type = ord.P_Type and P_Period=ord.P_Period
		FOR XML PATH('')) as BackwardDescs
		from #tmpwhinner ord
		) M 
		group by m.P_Type,m.P_Period,m.BackwardDescs,m.CP_ReserveYear,m.CP_ReserveMonth,m.RS_Season


        --擴大補助
		select I_City,P_ParentId,P_Period,P_Type,P_ItemName,CP_ReserveYear,CP_ReserveMonth,PD_Stage,PD_Season,PD_Summary,PD_BackwardDesc
		into #tmpupdateEx
		from PushItem 
		left join ProjectInfo on P_ParentId=I_Guid
		left join Check_Point on P_Guid=CP_ParentId and P_ParentId=CP_ProjectId
		left join PushItem_Desc
		on P_Guid=PD_PushitemGuid and P_ParentId=PD_ProjectGuid and P_Period=PD_Stage
		where P_ParentId=@wh_I_Guid and P_Type='04' and P_Period=@wh_RS_Stage and P_ItemName<>'99'
				and CP_ReserveYear=@wh_RS_year and CP_ReserveMonth=@wh_RS_Month

		--select * from #tmpupdate

		update #tmpall set RS_01Why=a.BackwardDescs
		from (
			select P_Type,P_Period,CP_ReserveYear,CP_ReserveMonth,RS_Season,BackwardDescs from #tmpupdate where P_Type='01'
		)a
		where #tmpall.RS_Stage=a.P_Period and #tmpall.RS_year=a.CP_ReserveYear and #tmpall.RS_Season=a.RS_Season and #tmpall.C_Item=@wh_C_Item

		update #tmpall set RS_02Why=a.BackwardDescs
		from (
			select P_Type,P_Period,CP_ReserveYear,CP_ReserveMonth,RS_Season,BackwardDescs from #tmpupdate where P_Type='02'
		)a
		where #tmpall.RS_Stage=a.P_Period and #tmpall.RS_year=a.CP_ReserveYear and #tmpall.RS_Season=a.RS_Season and #tmpall.C_Item=@wh_C_Item

		update #tmpall set RS_03Why=a.BackwardDescs
		from (
			select P_Type,P_Period,CP_ReserveYear,CP_ReserveMonth,RS_Season,BackwardDescs from #tmpupdate where P_Type='03'
		)a
		where #tmpall.RS_Stage=a.P_Period and #tmpall.RS_year=a.CP_ReserveYear and #tmpall.RS_Season=a.RS_Season and #tmpall.C_Item=@wh_C_Item

        --擴大補助
		update #tmpall set RS_ExWhy=a.PD_BackwardDesc
		from (
			select P_Type,P_Period,CP_ReserveYear,CP_ReserveMonth,PD_Season,PD_BackwardDesc from #tmpupdateEx where P_Type='04'
		)a
		where #tmpall.RS_Stage=a.P_Period and #tmpall.RS_year=a.CP_ReserveYear and #tmpall.RS_Season=a.PD_Season and #tmpall.C_Item=@wh_C_Item
		

		
		drop table #tmpwhinner
		drop table #tmpupdate
        drop table #tmpupdateEx
		delete from #tmpwhile where  C_Item= @wh_C_Item and I_Guid=@wh_I_Guid and RS_Stage=@wh_RS_Stage and RS_year=@wh_RS_year and RS_Season=@wh_RS_Season and RS_Month=@wh_RS_Month
	end

update #tmpall set #tmpall.rowsCount=#tmp.rCount
from #tmp
where #tmpall.C_Item=#tmp.C_Item

select * from #tmpall  order by C_Item asc,RS_Year asc,RS_Season asc

drop table #tmp
drop table #tmpall
drop table #tmpPJ
drop table #tmpwhile
        ");
        #endregion

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();

        oCmd.Parameters.AddWithValue("@strStage", strStage);
        oda.Fill(dt);
        return dt;
    }

    //各縣市申請數 報表(BY 季統計)
    public DataTable getReportTotalBehindg()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        #region 舊CODE
        //        sb.Append(@"
        //            create table #tmpcity(
        //	city_Item nvarchar(50),
        //	city_Item_cn nvarchar(50),
        //	city_I_Guid nvarchar(50),
        //	city_Stage nvarchar(5),
        //	city_Year nvarchar(5),
        //	city_Season nvarchar(5)
        //)
        //create table #tmpA(
        //	cityno nvarchar(50),
        //	citycn nvarchar(50),
        //	I_Guid nvarchar(50),
        //	RM_Stage nvarchar(5),
        //	RM_Year nvarchar(5),
        //	cpcn nvarchar(50),
        //	RM_Season nvarchar(5),
        //	RM_Sum decimal(10, 1),
        //	RM_SumC decimal(10, 1),
        //	RM_SumF decimal(10, 1)
        //)
        //create table #tmpB(
        //	cityno nvarchar(50),
        //	citycn nvarchar(50),
        //	I_Guid nvarchar(50),
        //	RM_Stage nvarchar(5),
        //	RM_Year nvarchar(5),
        //	cpcn nvarchar(50),
        //	RM_Season nvarchar(5),
        //	RM_Sum decimal(10, 1),
        //	RM_SumC decimal(10, 1),
        //	RM_SumF decimal(10, 1)
        //)
        //create table #tmpC(
        //	cityno nvarchar(50),
        //	citycn nvarchar(50),
        //	I_Guid nvarchar(50),
        //	RM_Stage nvarchar(5),
        //	RM_Year nvarchar(5),
        //	cpcn nvarchar(50),
        //	RM_Season nvarchar(5),
        //	RM_Sum decimal(10, 1),
        //	RM_SumC decimal(10, 1),
        //	RM_SumF decimal(10, 1)
        //)
        //create table #tmpD(
        //	cityno nvarchar(50),
        //	citycn nvarchar(50),
        //	I_Guid nvarchar(50),
        //	RM_Stage nvarchar(5),
        //	RM_Year nvarchar(5),
        //	cpcn nvarchar(50),
        //	RM_Season nvarchar(5),
        //	RM_Sum decimal(10, 1),
        //	RM_SumC decimal(10, 1),
        //	RM_SumF decimal(10, 1)
        //)
        //create table #tmpE(
        //	cityno nvarchar(50),
        //	citycn nvarchar(50),
        //	I_Guid nvarchar(50),
        //	RM_Stage nvarchar(5),
        //	RM_Year nvarchar(5),
        //	cpcn nvarchar(50),
        //	RM_Season nvarchar(5),
        //	RM_Sum decimal(10, 1),
        //	RM_SumC decimal(10, 1),
        //	RM_SumF decimal(10, 1)
        //)
        //insert into #tmpcity(city_Item,city_Item_cn,city_I_Guid,city_Stage,city_Year,city_Season)
        //select C_Item,C_Item_cn,I_Guid,RM_Stage,RM_Year,RM_Season
        //from(
        //select C_Item,C_Item_cn,I_Guid,RM_Stage,RM_Year
        //,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //	when '04' then '2' when '05' then '2' when '06' then '2'  
        //	when '07' then '3' when '08' then '3' when '09' then '3'  
        //	when '10' then '4' when '11' then '4' when '12' then '4'  
        //	end as RM_Season
        //from CodeTable 
        //left join ProjectInfo on C_Item=I_City and I_Status='A'
        //left join ReportMonth on I_Guid=RM_ProjectGuid
        //where C_Group='02'
        //)#tmp
        //group by C_Item,C_Item_cn,I_Guid,RM_Stage,RM_Year,RM_Season

        //if @strStage='1'
        //	begin
        //		--無風管
        //		insert into #tmpA(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        //        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season 
        //            ,(select I_Finish_item1_1 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A') as RM_Sum,RM_SumC,RM_SumF
        //        from (
        //		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //		    ,(select I_Finish_item1_1 from ProjectInfo c where c.I_Guid=I_Guid and c.I_Status='A') as RM_Sum
        //		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
        //		    from(
        //			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
        //			    ,isnull(RM_Type3ValueSum,0) as RM_Type3Value3,isnull(RM_Type4ValueSum,0) as RM_Type4Value3
        //			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //			    when '04' then '2' when '05' then '2' when '06' then '2'  
        //			    when '07' then '3' when '08' then '3' when '09' then '3'  
        //			    when '10' then '4' when '11' then '4' when '12' then '4'  
        //			    end as RM_Season,RC_CheckType,RC_Status
        //			    from  CodeTable a
        //			    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //			    left join ReportMonth on I_Guid=RM_ProjectGuid
        //			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
        //				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
        //			    where a.C_Group='02' and b.C_Item='01'
        //		    )#tmp1
        //			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
        //		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //        )#tmpttt
        //		--老舊
        //		insert into #tmpB(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        //		select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //            ,(select I_Finish_item2_1 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A') as RM_Sum,RM_SumC,RM_SumF
        //        from
        //        (
        //            select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
        //		    from(
        //			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
        //			    ,isnull(RM_Type3ValueSum,0) as RM_Type3Value3,isnull(RM_Type4ValueSum,0) as RM_Type4Value3
        //			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //			    when '04' then '2' when '05' then '2' when '06' then '2'  
        //			    when '07' then '3' when '08' then '3' when '09' then '3'  
        //			    when '10' then '4' when '11' then '4' when '12' then '4'  
        //			    end as RM_Season,RC_CheckType,RC_Status
        //			    from  CodeTable a
        //			    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //			    left join ReportMonth on I_Guid=RM_ProjectGuid
        //			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
        //				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
        //			    where a.C_Group='02' and b.C_Item='02'
        //		    )#tmp1
        //			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
        //		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //        )#tmpttt
        //		--停車場
        //		insert into #tmpC(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        //        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //            ,(select I_Finish_item3_1 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A') as RM_Sum,RM_SumC,RM_SumF
        //        from
        //        (
        //		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
        //		    from(
        //			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
        //			    ,isnull(RM_Type3ValueSum,0) as RM_Type3Value3,isnull(RM_Type4ValueSum,0) as RM_Type4Value3
        //			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //			    when '04' then '2' when '05' then '2' when '06' then '2'  
        //			    when '07' then '3' when '08' then '3' when '09' then '3'  
        //			    when '10' then '4' when '11' then '4' when '12' then '4'  
        //			    end as RM_Season,RC_CheckType,RC_Status
        //			    from  CodeTable a
        //			    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //			    left join ReportMonth on I_Guid=RM_ProjectGuid
        //			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
        //				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
        //			    where a.C_Group='02' and b.C_Item='03'
        //		    )#tmp1
        //			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
        //		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //        )#tmpttt
        //		--中型
        //		insert into #tmpD(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        //        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //            ,(select I_Finish_item4_1 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A') as RM_Sum,RM_SumC,RM_SumF
        //        from
        //        (
        //		select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //		,SUM(isnull(RM_Type1Value3,0)) as RM_SumC,SUM(isnull(RM_Type3Value3,0)) as RM_SumF
        //		from(
        //			select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
        //			,isnull(RM_Type1ValueSum,0) as RM_Type1Value3,isnull(RM_Type3ValueSum,0) as RM_Type3Value3
        //			,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //			when '04' then '2' when '05' then '2' when '06' then '2'  
        //			when '07' then '3' when '08' then '3' when '09' then '3'  
        //			when '10' then '4' when '11' then '4' when '12' then '4'  
        //			end as RM_Season,RC_CheckType,RC_Status
        //			from  CodeTable a
        //			left join ProjectInfo on C_Item=I_City and I_Status='A'
        //			left join ReportMonth on I_Guid=RM_ProjectGuid
        //			left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
        //			left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
        //			where a.C_Group='02' and b.C_Item='04'
        //		)#tmp1
        //		where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
        //		group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //        )#tmpttt
        //		--大型
        //		insert into #tmpE(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        //        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //            ,(select I_Finish_item5_1 from ProjectInfo c where I_Guid=#tmpttt.I_Guid and I_Status='A') as RM_Sum,RM_SumC,RM_SumF
        //        from
        //        (
        //		select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //		,SUM(isnull(RM_Type1Value3,0)) as RM_SumC,SUM(isnull(RM_Type3Value3,0)) as RM_SumF
        //		from(
        //			select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
        //			,isnull(RM_Type1Value3,0) as RM_Type1Value3,isnull(RM_Type3Value3,0) as RM_Type3Value3
        //			,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //			when '04' then '2' when '05' then '2' when '06' then '2'  
        //			when '07' then '3' when '08' then '3' when '09' then '3'  
        //			when '10' then '4' when '11' then '4' when '12' then '4'  
        //			end as RM_Season,RC_CheckType,RC_Status
        //			from  CodeTable a
        //			left join ProjectInfo on C_Item=I_City and I_Status='A'
        //			left join ReportMonth on I_Guid=RM_ProjectGuid
        //			left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
        //			left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
        //			where a.C_Group='02' and b.C_Item='05'
        //		)#tmp1
        //		where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
        //		group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //        )#tmpttt
        //end

        //if @strStage='2'
        //	begin
        //		--無風管
        //		insert into #tmpA(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        //        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //            ,(select I_Finish_item1_2 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A') as RM_Sum,RM_SumC,RM_SumF
        //        from (
        //		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
        //		    from(
        //			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
        //			    ,isnull(RM_Type3Value3,0) as RM_Type3Value3,isnull(RM_Type4Value3,0) as RM_Type4Value3
        //			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //			    when '04' then '2' when '05' then '2' when '06' then '2'  
        //			    when '07' then '3' when '08' then '3' when '09' then '3'  
        //			    when '10' then '4' when '11' then '4' when '12' then '4'  
        //			    end as RM_Season,RC_CheckType,RC_Status
        //			    from  CodeTable a
        //			    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //			    left join ReportMonth on I_Guid=RM_ProjectGuid
        //			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
        //				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
        //			    where a.C_Group='02' and b.C_Item='01'
        //		    )#tmp1
        //			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
        //		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //        )#tmpttt
        //		--老舊
        //		insert into #tmpB(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        //        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //            ,(select isnull(I_Finish_item2_2,0)*63 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A') as RM_Sum,RM_SumC,RM_SumF
        //        from (
        //		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
        //		    from(
        //			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
        //			    ,isnull(RM_Type3Value3,0) as RM_Type3Value3,isnull(RM_Type4Value3,0) as RM_Type4Value3
        //			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //			    when '04' then '2' when '05' then '2' when '06' then '2'  
        //			    when '07' then '3' when '08' then '3' when '09' then '3'  
        //			    when '10' then '4' when '11' then '4' when '12' then '4'  
        //			    end as RM_Season,RC_CheckType,RC_Status
        //			    from  CodeTable a
        //			    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //			    left join ReportMonth on I_Guid=RM_ProjectGuid
        //			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
        //				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
        //			    where a.C_Group='02' and b.C_Item='02'
        //		    )#tmp1
        //			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
        //		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //        )#tmpttt
        //		--停車場
        //		insert into #tmpC(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        //        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //            ,(select isnull(I_Finish_item3_2,0)*63 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A') as RM_Sum,RM_SumC,RM_SumF
        //        from (
        //		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
        //		    from(
        //			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
        //			    ,isnull(RM_Type3Value3,0) as RM_Type3Value3,isnull(RM_Type4Value3,0) as RM_Type4Value3
        //			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //			    when '04' then '2' when '05' then '2' when '06' then '2'  
        //			    when '07' then '3' when '08' then '3' when '09' then '3'  
        //			    when '10' then '4' when '11' then '4' when '12' then '4'  
        //			    end as RM_Season,RC_CheckType,RC_Status
        //			    from  CodeTable a
        //			    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //			    left join ReportMonth on I_Guid=RM_ProjectGuid
        //			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
        //				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
        //			    where a.C_Group='02' and b.C_Item='03'
        //		    )#tmp1
        //			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
        //		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //        )#tmpttt
        //		--中型
        //		insert into #tmpD(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        //        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //            ,(select I_Finish_item4_2 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A') as RM_Sum,RM_SumC,RM_SumF
        //        from (
        //		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //		    ,SUM(isnull(RM_Type1Value3,0)) as RM_SumC,SUM(isnull(RM_Type3Value3,0)) as RM_SumF
        //		    from(
        //			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
        //			    ,isnull(RM_Type1Value3,0) as RM_Type1Value3,isnull(RM_Type3Value3,0) as RM_Type3Value3
        //			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //			    when '04' then '2' when '05' then '2' when '06' then '2'  
        //			    when '07' then '3' when '08' then '3' when '09' then '3'  
        //			    when '10' then '4' when '11' then '4' when '12' then '4'  
        //			    end as RM_Season,RC_CheckType,RC_Status
        //			    from  CodeTable a
        //			    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //			    left join ReportMonth on I_Guid=RM_ProjectGuid
        //			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
        //				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
        //			    where a.C_Group='02' and b.C_Item='04'
        //		    )#tmp1
        //			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
        //		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //        )#tmpttt
        //		--大型
        //		insert into #tmpE(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        //        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //            ,(select I_Finish_item5_2 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A') as RM_Sum,RM_SumC,RM_SumF
        //        from (
        //		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //		    ,SUM(isnull(RM_Type1Value3,0)) as RM_SumC,SUM(isnull(RM_Type3Value3,0)) as RM_SumF
        //		    from(
        //			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
        //			    ,isnull(RM_Type1Value3,0) as RM_Type1Value3,isnull(RM_Type3Value3,0) as RM_Type3Value3
        //			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //			    when '04' then '2' when '05' then '2' when '06' then '2'  
        //			    when '07' then '3' when '08' then '3' when '09' then '3'  
        //			    when '10' then '4' when '11' then '4' when '12' then '4'  
        //			    end as RM_Season,RC_CheckType,RC_Status
        //			    from  CodeTable a
        //			    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //			    left join ReportMonth on I_Guid=RM_ProjectGuid
        //			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
        //				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
        //			    where a.C_Group='02' and b.C_Item='05'
        //		    )#tmp1
        //			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
        //		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //        )#tmpttt
        //end

        //if @strStage='3'
        //	begin
        //		--無風管
        //		insert into #tmpA(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        //        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //            ,(select I_Finish_item1_3 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A') as RM_Sum,RM_SumC,RM_SumF
        //        from (
        //		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
        //		    from(
        //			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
        //			    ,isnull(RM_Type3Value3,0) as RM_Type3Value3,isnull(RM_Type4Value3,0) as RM_Type4Value3
        //			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //			    when '04' then '2' when '05' then '2' when '06' then '2'  
        //			    when '07' then '3' when '08' then '3' when '09' then '3'  
        //			    when '10' then '4' when '11' then '4' when '12' then '4'  
        //			    end as RM_Season,RC_CheckType,RC_Status
        //			    from  CodeTable a
        //			    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //			    left join ReportMonth on I_Guid=RM_ProjectGuid
        //			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
        //				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
        //			    where a.C_Group='02' and b.C_Item='01'
        //		    )#tmp1
        //			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
        //		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //        )#tmpttt
        //		--老舊
        //		insert into #tmpB(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        //        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //            ,(select isnull(I_Finish_item2_3,0)*63 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A') as RM_Sum,RM_SumC,RM_SumF
        //        from (
        //		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
        //		    from(
        //			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
        //			    ,isnull(RM_Type3Value3,0) as RM_Type3Value3,isnull(RM_Type4Value3,0) as RM_Type4Value3
        //			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //			    when '04' then '2' when '05' then '2' when '06' then '2'  
        //			    when '07' then '3' when '08' then '3' when '09' then '3'  
        //			    when '10' then '4' when '11' then '4' when '12' then '4'  
        //			    end as RM_Season,RC_CheckType,RC_Status
        //			    from  CodeTable a
        //			    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //			    left join ReportMonth on I_Guid=RM_ProjectGuid
        //			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
        //				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
        //			    where a.C_Group='02' and b.C_Item='02'
        //		    )#tmp1
        //			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
        //		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //        )#tmpttt
        //		--停車場
        //		insert into #tmpC(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        //        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //            ,(select isnull(I_Finish_item3_3,0)*63 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A') as RM_Sum,RM_SumC,RM_SumF
        //        from (
        //		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season

        //		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
        //		    from(
        //			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
        //			    ,isnull(RM_Type3Value3,0) as RM_Type3Value3,isnull(RM_Type4Value3,0) as RM_Type4Value3
        //			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //			    when '04' then '2' when '05' then '2' when '06' then '2'  
        //			    when '07' then '3' when '08' then '3' when '09' then '3'  
        //			    when '10' then '4' when '11' then '4' when '12' then '4'  
        //			    end as RM_Season,RC_CheckType,RC_Status
        //			    from  CodeTable a
        //			    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //			    left join ReportMonth on I_Guid=RM_ProjectGuid
        //			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
        //				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
        //			    where a.C_Group='02' and b.C_Item='03'
        //		    )#tmp1
        //			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
        //		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //        )#tmpttt
        //		--中型
        //		insert into #tmpD(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        //        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //            ,(select I_Finish_item4_3 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A') as RM_Sum,RM_SumC,RM_SumF
        //        from (
        //		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //		    ,SUM(isnull(RM_Type1Value3,0)) as RM_SumC,SUM(isnull(RM_Type3Value3,0)) as RM_SumF
        //		    from(
        //			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
        //			    ,isnull(RM_Type1Value3,0) as RM_Type1Value3,isnull(RM_Type3Value3,0) as RM_Type3Value3
        //			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //			    when '04' then '2' when '05' then '2' when '06' then '2'  
        //			    when '07' then '3' when '08' then '3' when '09' then '3'  
        //			    when '10' then '4' when '11' then '4' when '12' then '4'  
        //			    end as RM_Season,RC_CheckType,RC_Status
        //			    from  CodeTable a
        //			    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //			    left join ReportMonth on I_Guid=RM_ProjectGuid
        //			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
        //				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
        //			    where a.C_Group='02' and b.C_Item='04'
        //		    )#tmp1
        //			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
        //		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //        )#tmpttt
        //		--大型
        //		insert into #tmpE(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        //        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //            ,(select I_Finish_item5_3 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A') as RM_Sum,RM_SumC,RM_SumF
        //        from (
        //		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //		    ,SUM(isnull(RM_Type1Value3,0)) as RM_SumC,SUM(isnull(RM_Type3Value3,0)) as RM_SumF
        //		    from(
        //			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
        //			    ,isnull(RM_Type1Value3,0) as RM_Type1Value3,isnull(RM_Type3Value3,0) as RM_Type3Value3
        //			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
        //			    when '04' then '2' when '05' then '2' when '06' then '2'  
        //			    when '07' then '3' when '08' then '3' when '09' then '3'  
        //			    when '10' then '4' when '11' then '4' when '12' then '4'  
        //			    end as RM_Season,RC_CheckType,RC_Status
        //			    from  CodeTable a
        //			    left join ProjectInfo on C_Item=I_City and I_Status='A'
        //			    left join ReportMonth on I_Guid=RM_ProjectGuid
        //			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
        //				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
        //			    where a.C_Group='02' and b.C_Item='05'
        //		    )#tmp1
        //			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
        //		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        //        )#tmpttt
        //end

        //select ct.*
        //,a.cpcn,isnull(a.RM_Sum,0) as RM_Sum01,isnull(a.RM_SumC,0) as RM_SumC01,isnull(a.RM_SumF,0) as RM_SumF01
        //,b.cpcn,isnull(b.RM_Sum,0) as RM_Sum02,isnull(b.RM_SumC,0) as RM_SumC02,isnull(b.RM_SumF,0) as RM_SumF02
        //,c.cpcn,isnull(c.RM_Sum,0) as RM_Sum03,isnull(c.RM_SumC,0) as RM_SumC03,isnull(c.RM_SumF,0) as RM_SumF03
        //,d.cpcn,isnull(d.RM_Sum,0) as RM_Sum04,isnull(d.RM_SumC,0) as RM_SumC04,isnull(d.RM_SumF,0) as RM_SumF04
        //,e.cpcn,isnull(e.RM_Sum,0) as RM_Sum05,isnull(e.RM_SumC,0) as RM_SumC05,isnull(e.RM_SumF,0) as RM_SumF05
        //from #tmpcity ct
        //left join #tmpA a on ct.city_Item=a.cityno and ct.city_I_Guid=a.I_Guid and ct.city_Stage=a.RM_Stage and ct.city_Year=a.RM_Year and ct.city_Season=a.RM_Season
        //left join #tmpB b on ct.city_Item=b.cityno and ct.city_I_Guid=b.I_Guid and ct.city_Stage=b.RM_Stage and ct.city_Year=b.RM_Year and ct.city_Season=b.RM_Season
        //left join #tmpC c on ct.city_Item=c.cityno and ct.city_I_Guid=c.I_Guid and ct.city_Stage=c.RM_Stage and ct.city_Year=c.RM_Year and ct.city_Season=c.RM_Season
        //left join #tmpD d on ct.city_Item=d.cityno and ct.city_I_Guid=d.I_Guid and ct.city_Stage=d.RM_Stage and ct.city_Year=d.RM_Year and ct.city_Season=d.RM_Season
        //left join #tmpE e on ct.city_Item=e.cityno and ct.city_I_Guid=e.I_Guid and ct.city_Stage=e.RM_Stage and ct.city_Year=e.RM_Year and ct.city_Season=e.RM_Season


        //drop table #tmpA
        //drop table #tmpB
        //drop table #tmpC
        //drop table #tmpD
        //drop table #tmpE
        //drop table #tmpcity


        //        ");
        #endregion

        #region 201808新code
        sb.Append(@"
 create table #tmpcity(
	city_Item nvarchar(50),
	city_Item_cn nvarchar(50),
	city_I_Guid nvarchar(50),
	city_Stage nvarchar(5),
	city_Year nvarchar(5),
	city_Season nvarchar(5)
)
create table #tmpA(
	cityno nvarchar(50),
	citycn nvarchar(50),
	I_Guid nvarchar(50),
	RM_Stage nvarchar(5),
	RM_Year nvarchar(5),
	cpcn nvarchar(50),
	RM_Season nvarchar(5),
	RM_Sum decimal(10, 1),
	RM_SumC decimal(10, 1),
	RM_SumF decimal(10, 1)
)
create table #tmpB(
	cityno nvarchar(50),
	citycn nvarchar(50),
	I_Guid nvarchar(50),
	RM_Stage nvarchar(5),
	RM_Year nvarchar(5),
	cpcn nvarchar(50),
	RM_Season nvarchar(5),
	RM_Sum decimal(10, 1),
	RM_SumC decimal(10, 1),
	RM_SumF decimal(10, 1)
)
create table #tmpC(
	cityno nvarchar(50),
	citycn nvarchar(50),
	I_Guid nvarchar(50),
	RM_Stage nvarchar(5),
	RM_Year nvarchar(5),
	cpcn nvarchar(50),
	RM_Season nvarchar(5),
	RM_Sum decimal(10, 1),
	RM_SumC decimal(10, 1),
	RM_SumF decimal(10, 1)
)
create table #tmpD(
	cityno nvarchar(50),
	citycn nvarchar(50),
	I_Guid nvarchar(50),
	RM_Stage nvarchar(5),
	RM_Year nvarchar(5),
	cpcn nvarchar(50),
	RM_Season nvarchar(5),
	RM_Sum decimal(10, 1),
	RM_SumC decimal(10, 1),
	RM_SumF decimal(10, 1)
)
create table #tmpE(
	cityno nvarchar(50),
	citycn nvarchar(50),
	I_Guid nvarchar(50),
	RM_Stage nvarchar(5),
	RM_Year nvarchar(5),
	cpcn nvarchar(50),
	RM_Season nvarchar(5),
	RM_Sum decimal(10, 1),
	RM_SumC decimal(10, 1),
	RM_SumF decimal(10, 1)
)
insert into #tmpcity(city_Item,city_Item_cn,city_I_Guid,city_Stage,city_Year,city_Season)
select C_Item,C_Item_cn,I_Guid,RM_Stage,RM_Year,RM_Season
from(
select C_Item,C_Item_cn,I_Guid,@strStage as RM_Stage,RM_Year
,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
	when '04' then '2' when '05' then '2' when '06' then '2'  
	when '07' then '3' when '08' then '3' when '09' then '3'  
	when '10' then '4' when '11' then '4' when '12' then '4'  
	end as RM_Season
from CodeTable 
left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
where C_Group='02' and RC_Status='A' and RC_CheckType='Y'
)#tmp
group by C_Item,C_Item_cn,I_Guid,RM_Stage,RM_Year,RM_Season

if @strStage='1'
	begin
		--無風管
		insert into #tmpA(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season 
            ,(select I_Finish_item1_1 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A' and I_Flag='Y') as RM_Sum,RM_SumC,RM_SumF
        from (
		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
		    ,(select I_Finish_item1_1 from ProjectInfo c where c.I_Guid=I_Guid and c.I_Status='A' and c.I_Flag='Y') as RM_Sum
		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
		    from(
			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
			    ,isnull(RM_Type3ValueSum,0) as RM_Type3Value3,isnull(RM_Type4ValueSum,0) as RM_Type4Value3
			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
			    when '04' then '2' when '05' then '2' when '06' then '2'  
			    when '07' then '3' when '08' then '3' when '09' then '3'  
			    when '10' then '4' when '11' then '4' when '12' then '4'  
			    end as RM_Season,RC_CheckType,RC_Status
			    from  CodeTable a
			    left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
			    left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
			    where a.C_Group='02' and b.C_Item='01'
		    )#tmp1
			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        )#tmpttt
		--老舊
		insert into #tmpB(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
		select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
            ,(select I_Finish_item2_1 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A' and I_Flag='Y') as RM_Sum,RM_SumC,RM_SumF
        from
        (
            select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
		    from(
			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
			    ,isnull(RM_Type1ValueSum,0) as RM_Type3Value3,isnull(RM_Type2ValueSum,0) as RM_Type4Value3
			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
			    when '04' then '2' when '05' then '2' when '06' then '2'  
			    when '07' then '3' when '08' then '3' when '09' then '3'  
			    when '10' then '4' when '11' then '4' when '12' then '4'  
			    end as RM_Season,RC_CheckType,RC_Status
			    from  CodeTable a
			    left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
			    left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
			    where a.C_Group='02' and b.C_Item='02'
		    )#tmp1
			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        )#tmpttt
		--停車場
		insert into #tmpC(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
            ,(select I_Finish_item3_1 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A' and I_Flag='Y') as RM_Sum,RM_SumC,RM_SumF
        from
        (
		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
		    from(
			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
			    ,isnull(RM_Type1ValueSum,0) as RM_Type3Value3,isnull(RM_Type2ValueSum,0) as RM_Type4Value3
			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
			    when '04' then '2' when '05' then '2' when '06' then '2'  
			    when '07' then '3' when '08' then '3' when '09' then '3'  
			    when '10' then '4' when '11' then '4' when '12' then '4'  
			    end as RM_Season,RC_CheckType,RC_Status
			    from  CodeTable a
			    left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
			    left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
			    where a.C_Group='02' and b.C_Item='03'
		    )#tmp1
			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        )#tmpttt
		--中型
		insert into #tmpD(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
            ,(select I_Finish_item4_1 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A' and I_Flag='Y') as RM_Sum,RM_SumC,RM_SumF
        from
        (
		select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
		,SUM(isnull(RM_Type1Value3,0)) as RM_SumC,SUM(isnull(RM_Type3Value3,0)) as RM_SumF
		from(
			select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
			,isnull(RM_Type1ValueSum,0) as RM_Type1Value3,isnull(RM_Type3ValueSum,0) as RM_Type3Value3
			,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
			when '04' then '2' when '05' then '2' when '06' then '2'  
			when '07' then '3' when '08' then '3' when '09' then '3'  
			when '10' then '4' when '11' then '4' when '12' then '4'  
			end as RM_Season,RC_CheckType,RC_Status
			from  CodeTable a
			left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
			left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
			left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
			left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
			where a.C_Group='02' and b.C_Item='04'
		)#tmp1
		where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
		group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        )#tmpttt
		--大型
		insert into #tmpE(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
            ,(select I_Finish_item5_1 from ProjectInfo c where I_Guid=#tmpttt.I_Guid and I_Status='A' and I_Flag='Y') as RM_Sum,RM_SumC,RM_SumF
        from
        (
		select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
		,SUM(isnull(RM_Type1Value3,0)) as RM_SumC,SUM(isnull(RM_Type3Value3,0)) as RM_SumF
		from(
			select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
			,isnull(RM_Type1ValueSum,0) as RM_Type1Value3,isnull(RM_Type3ValueSum,0) as RM_Type3Value3
			,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
			when '04' then '2' when '05' then '2' when '06' then '2'  
			when '07' then '3' when '08' then '3' when '09' then '3'  
			when '10' then '4' when '11' then '4' when '12' then '4'  
			end as RM_Season,RC_CheckType,RC_Status
			from  CodeTable a
			left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
			left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
			left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
			left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
			where a.C_Group='02' and b.C_Item='05'
		)#tmp1
		where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
		group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        )#tmpttt
end

if @strStage='2'
	begin
		--無風管
		insert into #tmpA(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
            ,(select I_Finish_item1_2 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A' and I_Flag='Y') as RM_Sum,RM_SumC,RM_SumF
        from (
		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
		    from(
			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,@strStage as RM_Stage,RM_Year,b.C_Item_cn as cpcn
			    ,isnull(RM_Type3ValueSum,0) as RM_Type3Value3,isnull(RM_Type4ValueSum,0) as RM_Type4Value3
			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
			    when '04' then '2' when '05' then '2' when '06' then '2'  
			    when '07' then '3' when '08' then '3' when '09' then '3'  
			    when '10' then '4' when '11' then '4' when '12' then '4'  
			    end as RM_Season,RC_CheckType,RC_Status
			    from  CodeTable a
			    left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
			    left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
			    where a.C_Group='02' and b.C_Item='01' and RM_Stage=@strStage
		    )#tmp1
			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        )#tmpttt
		--老舊
		insert into #tmpB(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
            ,(select isnull(I_Finish_item2_2,0)*63 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A' and I_Flag='Y') as RM_Sum,RM_SumC,RM_SumF
        from (
		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
		    from(
			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,@strStage as RM_Stage,RM_Year,b.C_Item_cn as cpcn
			    ,isnull(RM_Type1ValueSum,0) as RM_Type3Value3,isnull(RM_Type2ValueSum,0) as RM_Type4Value3
			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
			    when '04' then '2' when '05' then '2' when '06' then '2'  
			    when '07' then '3' when '08' then '3' when '09' then '3'  
			    when '10' then '4' when '11' then '4' when '12' then '4'  
			    end as RM_Season,RC_CheckType,RC_Status
			    from  CodeTable a
			    left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
			    left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
			    where a.C_Group='02' and b.C_Item='02'
		    )#tmp1
			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        )#tmpttt
		--停車場
		insert into #tmpC(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
            ,(select isnull(I_Finish_item3_2,0)*63 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A' and I_Flag='Y') as RM_Sum,RM_SumC,RM_SumF
        from (
		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
		    from(
			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,@strStage as RM_Stage,RM_Year,b.C_Item_cn as cpcn
			    ,isnull(RM_Type1ValueSum,0) as RM_Type3Value3,isnull(RM_Type2ValueSum,0) as RM_Type4Value3
			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
			    when '04' then '2' when '05' then '2' when '06' then '2'  
			    when '07' then '3' when '08' then '3' when '09' then '3'  
			    when '10' then '4' when '11' then '4' when '12' then '4'  
			    end as RM_Season,RC_CheckType,RC_Status
			    from  CodeTable a
			    left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
			    left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
			    where a.C_Group='02' and b.C_Item='03'
		    )#tmp1
			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        )#tmpttt
		--中型
		insert into #tmpD(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
            ,(select I_Finish_item4_2 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A' and I_Flag='Y') as RM_Sum,RM_SumC,RM_SumF
        from (
		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
		    ,SUM(isnull(RM_Type1Value3,0)) as RM_SumC,SUM(isnull(RM_Type3Value3,0)) as RM_SumF
		    from(
			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,@strStage as RM_Stage,RM_Year,b.C_Item_cn as cpcn
			    ,isnull(RM_Type1ValueSum,0) as RM_Type1Value3,isnull(RM_Type3ValueSum,0) as RM_Type3Value3
			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
			    when '04' then '2' when '05' then '2' when '06' then '2'  
			    when '07' then '3' when '08' then '3' when '09' then '3'  
			    when '10' then '4' when '11' then '4' when '12' then '4'  
			    end as RM_Season,RC_CheckType,RC_Status
			    from  CodeTable a
			    left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
			    left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
			    where a.C_Group='02' and b.C_Item='04'
		    )#tmp1
			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        )#tmpttt
		--大型
		insert into #tmpE(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
            ,(select I_Finish_item5_2 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A' and I_Flag='Y') as RM_Sum,RM_SumC,RM_SumF
        from (
		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
		    ,SUM(isnull(RM_Type1Value3,0)) as RM_SumC,SUM(isnull(RM_Type3Value3,0)) as RM_SumF
		    from(
			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,@strStage as RM_Stage,RM_Year,b.C_Item_cn as cpcn
			    ,isnull(RM_Type1ValueSum,0) as RM_Type1Value3,isnull(RM_Type3ValueSum,0) as RM_Type3Value3
			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
			    when '04' then '2' when '05' then '2' when '06' then '2'  
			    when '07' then '3' when '08' then '3' when '09' then '3'  
			    when '10' then '4' when '11' then '4' when '12' then '4'  
			    end as RM_Season,RC_CheckType,RC_Status
			    from  CodeTable a
			    left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
			    left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
			    where a.C_Group='02' and b.C_Item='05'
		    )#tmp1
			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        )#tmpttt
end

if @strStage='3'
	begin
		--無風管
		insert into #tmpA(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
            ,(select I_Finish_item1_3 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A' and I_Flag='Y') as RM_Sum,RM_SumC,RM_SumF
        from (
		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
		    from(
			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
			    ,isnull(RM_Type3ValueSum,0) as RM_Type3Value3,isnull(RM_Type4ValueSum,0) as RM_Type4Value3
			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
			    when '04' then '2' when '05' then '2' when '06' then '2'  
			    when '07' then '3' when '08' then '3' when '09' then '3'  
			    when '10' then '4' when '11' then '4' when '12' then '4'  
			    end as RM_Season,RC_CheckType,RC_Status
			    from  CodeTable a
			    left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
			    left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
			    where a.C_Group='02' and b.C_Item='01'
		    )#tmp1
			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        )#tmpttt
		--老舊
		insert into #tmpB(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
            ,(select isnull(I_Finish_item2_3,0)*63 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A' and I_Flag='Y') as RM_Sum,RM_SumC,RM_SumF
        from (
		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
		    from(
			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
			    ,isnull(RM_Type1ValueSum,0) as RM_Type3Value3,isnull(RM_Type2ValueSum,0) as RM_Type4Value3
			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
			    when '04' then '2' when '05' then '2' when '06' then '2'  
			    when '07' then '3' when '08' then '3' when '09' then '3'  
			    when '10' then '4' when '11' then '4' when '12' then '4'  
			    end as RM_Season,RC_CheckType,RC_Status
			    from  CodeTable a
			    left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
			    left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
			    where a.C_Group='02' and b.C_Item='02'
		    )#tmp1
			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        )#tmpttt
		--停車場
		insert into #tmpC(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
            ,(select isnull(I_Finish_item3_3,0)*63 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A' and I_Flag='Y') as RM_Sum,RM_SumC,RM_SumF
        from (
		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
		                
		    ,SUM(isnull(RM_Type3Value3,0)) as RM_SumC,SUM(isnull(RM_Type4Value3,0)) as RM_SumF
		    from(
			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
			    ,isnull(RM_Type1ValueSum,0) as RM_Type3Value3,isnull(RM_Type2ValueSum,0) as RM_Type4Value3
			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
			    when '04' then '2' when '05' then '2' when '06' then '2'  
			    when '07' then '3' when '08' then '3' when '09' then '3'  
			    when '10' then '4' when '11' then '4' when '12' then '4'  
			    end as RM_Season,RC_CheckType,RC_Status
			    from  CodeTable a
			    left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
			    left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
			    where a.C_Group='02' and b.C_Item='03'
		    )#tmp1
			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        )#tmpttt
		--中型
		insert into #tmpD(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
            ,(select I_Finish_item4_3 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A' and I_Flag='Y') as RM_Sum,RM_SumC,RM_SumF
        from (
		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
		    ,SUM(isnull(RM_Type1Value3,0)) as RM_SumC,SUM(isnull(RM_Type3Value3,0)) as RM_SumF
		    from(
			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
			    ,isnull(RM_Type1ValueSum,0) as RM_Type1Value3,isnull(RM_Type3ValueSum,0) as RM_Type3Value3
			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
			    when '04' then '2' when '05' then '2' when '06' then '2'  
			    when '07' then '3' when '08' then '3' when '09' then '3'  
			    when '10' then '4' when '11' then '4' when '12' then '4'  
			    end as RM_Season,RC_CheckType,RC_Status
			    from  CodeTable a
			    left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
			    left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
			    where a.C_Group='02' and b.C_Item='04'
		    )#tmp1
			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        )#tmpttt
		--大型
		insert into #tmpE(cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season,RM_Sum,RM_SumC,RM_SumF)
        select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
            ,(select I_Finish_item5_3 from ProjectInfo c where #tmpttt.I_Guid=I_Guid and I_Status='A' and I_Flag='Y') as RM_Sum,RM_SumC,RM_SumF
        from (
		    select cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
		    ,SUM(isnull(RM_Type1Value3,0)) as RM_SumC,SUM(isnull(RM_Type3Value3,0)) as RM_SumF
		    from(
			    select a.C_Item as cityno,a.C_Item_cn as citycn,I_Guid,RM_Stage,RM_Year,b.C_Item_cn as cpcn
			    ,isnull(RM_Type1ValueSum,0) as RM_Type1Value3,isnull(RM_Type3ValueSum,0) as RM_Type3Value3
			    ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
			    when '04' then '2' when '05' then '2' when '06' then '2'  
			    when '07' then '3' when '08' then '3' when '09' then '3'  
			    when '10' then '4' when '11' then '4' when '12' then '4'  
			    end as RM_Season,RC_CheckType,RC_Status
			    from  CodeTable a
			    left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
			    left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='01'
			    left join CodeTable b on b.C_Group='07' and b.C_Item = RM_CPType
				left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
			    where a.C_Group='02' and b.C_Item='05'
		    )#tmp1
			where #tmp1.RC_CheckType='Y' and #tmp1.RC_Status='A'
		    group by cityno,citycn,I_Guid,RM_Stage,RM_Year,cpcn,RM_Season
        )#tmpttt
end



select ct.*
,a.cpcn,isnull(a.RM_Sum,0) as RM_Sum01,isnull(a.RM_SumC,0) as RM_SumC01,isnull(a.RM_SumF,0) as RM_SumF01
,b.cpcn,isnull(b.RM_Sum,0) as RM_Sum02,isnull(b.RM_SumC,0) as RM_SumC02,isnull(b.RM_SumF,0) as RM_SumF02
,c.cpcn,isnull(c.RM_Sum,0) as RM_Sum03,isnull(c.RM_SumC,0) as RM_SumC03,isnull(c.RM_SumF,0) as RM_SumF03
,d.cpcn,isnull(d.RM_Sum,0) as RM_Sum04,isnull(d.RM_SumC,0) as RM_SumC04,isnull(d.RM_SumF,0) as RM_SumF04
,e.cpcn,isnull(e.RM_Sum,0) as RM_Sum05,isnull(e.RM_SumC,0) as RM_SumC05,isnull(e.RM_SumF,0) as RM_SumF05
from #tmpcity ct
left join #tmpA a on ct.city_Item=a.cityno and ct.city_I_Guid=a.I_Guid and ct.city_Stage=a.RM_Stage and ct.city_Year=a.RM_Year and ct.city_Season=a.RM_Season
left join #tmpB b on ct.city_Item=b.cityno and ct.city_I_Guid=b.I_Guid and ct.city_Stage=b.RM_Stage and ct.city_Year=b.RM_Year and ct.city_Season=b.RM_Season
left join #tmpC c on ct.city_Item=c.cityno and ct.city_I_Guid=c.I_Guid and ct.city_Stage=c.RM_Stage and ct.city_Year=c.RM_Year and ct.city_Season=c.RM_Season
left join #tmpD d on ct.city_Item=d.cityno and ct.city_I_Guid=d.I_Guid and ct.city_Stage=d.RM_Stage and ct.city_Year=d.RM_Year and ct.city_Season=d.RM_Season
left join #tmpE e on ct.city_Item=e.cityno and ct.city_I_Guid=e.I_Guid and ct.city_Stage=e.RM_Stage and ct.city_Year=e.RM_Year and ct.city_Season=e.RM_Season
where a.cpcn is not null or b.cpcn is not null or c.cpcn is not null or d.cpcn is not null or e.cpcn is not null

drop table #tmpA
drop table #tmpB
drop table #tmpC
drop table #tmpD
drop table #tmpE
drop table #tmpcity



        ");
        #endregion
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();

        oCmd.Parameters.AddWithValue("@strStage", strStage);
        oda.Fill(dt);
        return dt;
    }

    //各縣市申請數 報表(BY 月累計)
    public DataTable getReportTotalBehindgByMonth()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        #region 舊code
   //     sb.Append(@"
   //         create table #tmpcity(
	  //          city_Item nvarchar(50),
	  //          city_Item_cn nvarchar(50),
	  //          city_I_Guid nvarchar(50),
	  //          city_Stage nvarchar(5),
	  //          city_Year nvarchar(5),
	  //          city_Month nvarchar(5),
	  //          city_maxmonth nvarchar(5)
   //         );

   //         --先撈出各縣市代碼 名稱 計畫GUID
   //         insert into #tmpcity(city_Item,city_Item_cn,city_I_Guid,city_Stage,city_Year,city_Month,city_maxmonth)
   //         select C_Item,C_Item_cn,I_Guid,'','','',''
   //         from CodeTable 
   //         left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
   //         left join ReportMonth on I_Guid=RM_ProjectGuid
   //         where C_Group='02'
   //         group by C_Item,C_Item_cn,I_Guid,RM_Stage
   //         --select * from #tmpcity

   //         --根據計畫GUID撈出該期底下已審核通過的月報最大的年月份
   //         select #tmp.RM_ProjectGuid,b.RM_ReportGuid,#tmp.maxmonth,b.RM_Year,b.RM_Month into #tmptmp
   //         from(
	  //          select RM_ProjectGuid,MAX(CONVERT(int,isnull(RM_Year,'0')+isnull(RM_Month,'0'))) maxmonth 
	  //          from ReportMonth 
	  //          left join ReportCheck on RM_ReportGuid=RC_ReportGuid
	  //          where RM_Stage=@str_stage and RC_CheckType ='Y'
	  //          group by RM_ProjectGuid
   //         )#tmp
   //         left join ReportMonth b on 
   //         #tmp.RM_ProjectGuid=b.RM_ProjectGuid and b.RM_Stage=@str_stage
   //         and #tmp.maxmonth=CONVERT(int,(isnull(b.RM_Year,'0')+isnull(b.RM_Month,'0')))
   //         group by #tmp.RM_ProjectGuid,b.RM_ReportGuid,#tmp.maxmonth,b.RM_Year,b.RM_Month

   //         --select * from #tmptmp

   //         -- 將已審核最大年月資料update 回#tmpcity
   //         update #tmpcity set city_Year=b.RM_Year,city_Month=b.RM_Month
   //         from
   //         (
	  //          select RM_ProjectGuid,RM_Year,RM_Month from #tmptmp
   //         )b
   //         where city_I_Guid = b.RM_ProjectGuid

   //         --過濾出有審核通過的月報資料
			//select * into #tmpRM from(
			//select a.*,b.RC_ID from ReportMonth a left join ReportCheck b on a.RM_ReportGuid=b.RC_ReportGuid and b.RC_Status='A' and RC_CheckType='Y'
			//)#tmp
			//where #tmp.RC_ID is not null

   //         --用#tmpcity去撈出累計值
   //         declare @stage_item1_col nvarchar(50)='I_Finish_item1_'+@str_stage --無風館冷氣第X期規劃數
   //         declare @stage_item2_col nvarchar(50)='I_Finish_item2_'+@str_stage --老舊辦公室照明第X期規劃數
   //         declare @stage_item3_col nvarchar(50)='I_Finish_item3_'+@str_stage --室內停車場智慧照明第X期規劃數
   //         declare @stage_item4_col nvarchar(50)='I_Finish_item4_'+@str_stage --中型能管系統第X期規劃數
   //         declare @stage_item5_col nvarchar(50)='I_Finish_item5_'+@str_stage --大型能管系統第X期規劃數
   //         declare @sql nvarchar(max)='';

   //         set @sql='
   //         select a.city_Item,a.city_Item_cn,a.city_I_Guid,a.city_Stage,a.city_Year,a.city_Month,a.city_maxmonth
   //         ,(select '+@stage_item1_col+' from ProjectInfo where I_Guid=a.city_I_Guid and a.city_I_Guid is not null) as I_Finish_item1
   //         ,(select SUM(isnull(b.RM_Type3ValueSum,0)) from #tmpRM b where a.city_I_Guid=b.RM_ProjectGuid and b.RM_Stage='+@str_stage+' and b.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(b.RM_Year,''0'')+isnull(b.RM_Month,''0''))) and RM_CPType=''01'' ) sumC_item1
   //         ,(select SUM(isnull(c.RM_Type4ValueSum,0)) from #tmpRM c where a.city_I_Guid=c.RM_ProjectGuid and c.RM_Stage='+@str_stage+' and c.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(c.RM_Year,''0'')+isnull(c.RM_Month,''0''))) and RM_CPType=''01'' ) sumF_item1
   //         ,(select '+@stage_item2_col+' from ProjectInfo where I_Guid=a.city_I_Guid and a.city_I_Guid is not null) as I_Finish_item2
   //         ,(select SUM(isnull(b.RM_Type3ValueSum,0)) from #tmpRM b where a.city_I_Guid=b.RM_ProjectGuid and b.RM_Stage='+@str_stage+' and b.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(b.RM_Year,''0'')+isnull(b.RM_Month,''0''))) and RM_CPType=''02'' ) sumC_item2
   //         ,(select SUM(isnull(c.RM_Type4ValueSum,0)) from #tmpRM c where a.city_I_Guid=c.RM_ProjectGuid and c.RM_Stage='+@str_stage+' and c.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(c.RM_Year,''0'')+isnull(c.RM_Month,''0''))) and RM_CPType=''02'' ) sumF_item2
   //         ,(select '+@stage_item3_col+' from ProjectInfo where I_Guid=a.city_I_Guid and a.city_I_Guid is not null) as I_Finish_item3
   //         ,(select SUM(isnull(b.RM_Type3ValueSum,0)) from #tmpRM b where a.city_I_Guid=b.RM_ProjectGuid and b.RM_Stage='+@str_stage+' and b.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(b.RM_Year,''0'')+isnull(b.RM_Month,''0''))) and RM_CPType=''03'' ) sumC_item3
   //         ,(select SUM(isnull(c.RM_Type4ValueSum,0)) from #tmpRM c where a.city_I_Guid=c.RM_ProjectGuid and c.RM_Stage='+@str_stage+' and c.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(c.RM_Year,''0'')+isnull(c.RM_Month,''0''))) and RM_CPType=''03'' ) sumF_item3
   //         ,(select '+@stage_item4_col+' from ProjectInfo where I_Guid=a.city_I_Guid and a.city_I_Guid is not null) as I_Finish_item4
   //         ,(select SUM(isnull(b.RM_Type1ValueSum,0)) from #tmpRM b where a.city_I_Guid=b.RM_ProjectGuid and b.RM_Stage='+@str_stage+' and b.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(b.RM_Year,''0'')+isnull(b.RM_Month,''0''))) and RM_CPType=''04'' ) sumC_item4
   //         ,(select SUM(isnull(c.RM_Type3ValueSum,0)) from #tmpRM c where a.city_I_Guid=c.RM_ProjectGuid and c.RM_Stage='+@str_stage+' and c.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(c.RM_Year,''0'')+isnull(c.RM_Month,''0''))) and RM_CPType=''04'' ) sumF_item4
   //         ,(select '+@stage_item5_col+' from ProjectInfo where I_Guid=a.city_I_Guid and a.city_I_Guid is not null) as I_Finish_item5
   //         ,(select SUM(isnull(b.RM_Type1ValueSum,0)) from #tmpRM b where a.city_I_Guid=b.RM_ProjectGuid and b.RM_Stage='+@str_stage+' and b.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(b.RM_Year,''0'')+isnull(b.RM_Month,''0''))) and RM_CPType=''05'' ) sumC_item5
   //         ,(select SUM(isnull(c.RM_Type3ValueSum,0)) from #tmpRM c where a.city_I_Guid=c.RM_ProjectGuid and c.RM_Stage='+@str_stage+' and c.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(c.RM_Year,''0'')+isnull(c.RM_Month,''0''))) and RM_CPType=''05'' ) sumF_item5
   //         into #tmpAll
   //         from #tmpcity a

   //         select * from 
			//(
			//select a.*,b.I_ID
			//    from #tmpAll a left join ProjectInfo b on a.city_I_Guid = b.I_Guid and b.I_Flag=''Y''
			//)#tmpR
			//where #tmpR.I_ID is not null
   //         ';

   //         EXECUTE sp_executesql @sql;
            
   //         drop table #tmptmp
   //         drop table #tmpcity

   //     ");
        #endregion

        #region 201808新code
        sb.Append(@"
create table #tmpcity(
	city_Item nvarchar(50),
	city_Item_cn nvarchar(50),
	city_I_Guid nvarchar(50),
	city_Stage nvarchar(5),
	city_Year nvarchar(5),
	city_Month nvarchar(5),
	city_maxmonth nvarchar(5)
);

--先撈出各縣市代碼 名稱 計畫GUID
insert into #tmpcity(city_Item,city_Item_cn,city_I_Guid,city_Stage,city_Year,city_Month,city_maxmonth)
select C_Item,C_Item_cn,I_Guid,'','','',''
from CodeTable 
left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
--left join ReportMonth on I_Guid=RM_ProjectGuid and RM_ReportType='01'
where C_Group='02'
--group by C_Item,C_Item_cn,I_Guid--,RM_Stage
--select * from #tmpcity

--根據計畫GUID撈出該期底下已審核通過的月報最大的年月份
select #tmp.RM_ProjectGuid,b.RM_ReportGuid,#tmp.maxmonth,b.RM_Stage,b.RM_Year,b.RM_Month into #tmptmp
from(
	select RM_ProjectGuid,MAX(CONVERT(int,isnull(RM_Year,'0')+isnull(RM_Month,'0'))) maxmonth 
	from ReportMonth 
	left join ReportCheck on RM_ReportGuid=RC_ReportGuid
	where RM_Stage=@str_stage and RC_CheckType ='Y' and RC_Status='A' and RM_ReportType='01'
	group by RM_ProjectGuid
)#tmp
left join ReportMonth b on 
#tmp.RM_ProjectGuid=b.RM_ProjectGuid and b.RM_Stage=@str_stage
and #tmp.maxmonth=CONVERT(int,(isnull(b.RM_Year,'0')+isnull(b.RM_Month,'0')))
group by #tmp.RM_ProjectGuid,b.RM_ReportGuid,#tmp.maxmonth,b.RM_Year,b.RM_Month,b.RM_Stage

--select * from #tmptmp

-- 將已審核最大年月資料update 回#tmpcity
update #tmpcity set city_Year=b.RM_Year,city_Month=b.RM_Month,city_Stage=b.RM_Stage
from
(
	select RM_ProjectGuid,RM_Year,RM_Month,RM_Stage from #tmptmp
)b
where city_I_Guid = b.RM_ProjectGuid

--過濾出有審核通過的月報資料
select * into #tmpRM from(
select a.*,b.RC_ID from ReportMonth a left join ReportCheck b on a.RM_ReportGuid=b.RC_ReportGuid and b.RC_Status='A' and RC_CheckType='Y' and a.RM_ReportType='01'
)#tmp
where #tmp.RC_ID is not null

--用#tmpcity去撈出累計值
declare @stage_item1_col nvarchar(50)='I_Finish_item1_'+@str_stage --無風館冷氣第X期規劃數
declare @stage_item2_col nvarchar(50)='I_Finish_item2_'+@str_stage --老舊辦公室照明第X期規劃數
declare @stage_item3_col nvarchar(50)='I_Finish_item3_'+@str_stage --室內停車場智慧照明第X期規劃數
declare @stage_item4_col nvarchar(50)='I_Finish_item4_'+@str_stage --中型能管系統第X期規劃數
declare @stage_item5_col nvarchar(50)='I_Finish_item5_'+@str_stage --大型能管系統第X期規劃數
declare @sql nvarchar(max)='';

set @sql='
select a.city_Item,a.city_Item_cn,a.city_I_Guid,a.city_Stage,a.city_Year,a.city_Month,a.city_maxmonth
,(select '+@stage_item1_col+' from ProjectInfo where I_Guid=a.city_I_Guid and a.city_I_Guid is not null) as I_Finish_item1
,(select SUM(isnull(b.RM_Type3ValueSum,0)) from #tmpRM b where a.city_I_Guid=b.RM_ProjectGuid and b.RM_Stage='+@str_stage+' and b.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(b.RM_Year,''0'')+isnull(b.RM_Month,''0''))) and RM_CPType=''01'' and RM_ReportType=''01'' ) sumC_item1
,(select SUM(isnull(c.RM_Type4ValueSum,0)) from #tmpRM c where a.city_I_Guid=c.RM_ProjectGuid and c.RM_Stage='+@str_stage+' and c.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(c.RM_Year,''0'')+isnull(c.RM_Month,''0''))) and RM_CPType=''01'' and RM_ReportType=''01'' ) sumF_item1
,(select '+@stage_item2_col+' from ProjectInfo where I_Guid=a.city_I_Guid and a.city_I_Guid is not null) as I_Finish_item2
,(select SUM(isnull(b.RM_Type1ValueSum,0)) from #tmpRM b where a.city_I_Guid=b.RM_ProjectGuid and b.RM_Stage='+@str_stage+' and b.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(b.RM_Year,''0'')+isnull(b.RM_Month,''0''))) and RM_CPType=''02'' and RM_ReportType=''01'' ) sumC_item2
,(select SUM(isnull(c.RM_Type2ValueSum,0)) from #tmpRM c where a.city_I_Guid=c.RM_ProjectGuid and c.RM_Stage='+@str_stage+' and c.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(c.RM_Year,''0'')+isnull(c.RM_Month,''0''))) and RM_CPType=''02'' and RM_ReportType=''01'' ) sumF_item2
,(select '+@stage_item3_col+' from ProjectInfo where I_Guid=a.city_I_Guid and a.city_I_Guid is not null) as I_Finish_item3
,(select SUM(isnull(b.RM_Type1ValueSum,0)) from #tmpRM b where a.city_I_Guid=b.RM_ProjectGuid and b.RM_Stage='+@str_stage+' and b.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(b.RM_Year,''0'')+isnull(b.RM_Month,''0''))) and RM_CPType=''03'' and RM_ReportType=''01'' ) sumC_item3
,(select SUM(isnull(c.RM_Type2ValueSum,0)) from #tmpRM c where a.city_I_Guid=c.RM_ProjectGuid and c.RM_Stage='+@str_stage+' and c.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(c.RM_Year,''0'')+isnull(c.RM_Month,''0''))) and RM_CPType=''03'' and RM_ReportType=''01'' ) sumF_item3
,(select '+@stage_item4_col+' from ProjectInfo where I_Guid=a.city_I_Guid and a.city_I_Guid is not null) as I_Finish_item4
,(select SUM(isnull(b.RM_Type1ValueSum,0)) from #tmpRM b where a.city_I_Guid=b.RM_ProjectGuid and b.RM_Stage='+@str_stage+' and b.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(b.RM_Year,''0'')+isnull(b.RM_Month,''0''))) and RM_CPType=''04'' and RM_ReportType=''01'' ) sumC_item4
,(select SUM(isnull(c.RM_Type3ValueSum,0)) from #tmpRM c where a.city_I_Guid=c.RM_ProjectGuid and c.RM_Stage='+@str_stage+' and c.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(c.RM_Year,''0'')+isnull(c.RM_Month,''0''))) and RM_CPType=''04'' and RM_ReportType=''01'' ) sumF_item4
,(select '+@stage_item5_col+' from ProjectInfo where I_Guid=a.city_I_Guid and a.city_I_Guid is not null) as I_Finish_item5
,(select SUM(isnull(b.RM_Type1ValueSum,0)) from #tmpRM b where a.city_I_Guid=b.RM_ProjectGuid and b.RM_Stage='+@str_stage+' and b.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(b.RM_Year,''0'')+isnull(b.RM_Month,''0''))) and RM_CPType=''05'' and RM_ReportType=''01'' ) sumC_item5
,(select SUM(isnull(c.RM_Type3ValueSum,0)) from #tmpRM c where a.city_I_Guid=c.RM_ProjectGuid and c.RM_Stage='+@str_stage+' and c.RM_Status=''A'' and a.city_maxmonth is not null and a.city_maxmonth <=CONVERT(int,(isnull(c.RM_Year,''0'')+isnull(c.RM_Month,''0''))) and RM_CPType=''05'' and RM_ReportType=''01'' ) sumF_item5
into #tmpAll
from #tmpcity a

select * from 
(
select a.*,b.I_ID
	from #tmpAll a left join ProjectInfo b on a.city_I_Guid = b.I_Guid and b.I_Flag=''Y'' and a.city_Stage<>''''
)#tmpR
where #tmpR.I_ID is not null
';

EXECUTE sp_executesql @sql;
            
drop table #tmptmp
drop table #tmpRM
drop table #tmpcity

        ");

        #endregion
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();

        oCmd.Parameters.AddWithValue("@str_stage", strStage);
        oda.Fill(dt);
        return dt;
        //city_Item 各縣市代碼
        //city_Item_cn 各縣市中文名稱
        //city_I_Guid 各縣市計畫GUID
        //city_Stage 查詢期數
        //city_Year 審核過月報最大年分(ex:2018)
        //city_Month 審核過月報最大月分(ex:12)
        //city_maxmonth 審核過月報最大年分+審核過月報最大月分(ex: 201812)
        //I_Finish_item1 無風館冷氣第X期規劃數
        //sumC_item1 無風館冷氣累計申請數
        //sumF_item1 無風館冷氣累計完成數
        //I_Finish_item2 老舊辦公室照明第X期規劃數
        //sumC_item2 老舊辦公室照明累計申請數
        //sumF_item2 老舊辦公室照明累計完成數
        //I_Finish_item3 室內停車場智慧照明第X期規劃數
        //sumC_item3 室內停車場智慧照明累計申請數
        //sumF_item3 室內停車場智慧照明累計完成數
        //I_Finish_item4 中型能管系統第X期規劃數
        //sumC_item4 中型能管系統累計申請數
        //sumF_item4 中型能管系統累計完成數
        //I_Finish_item5 大型能管系統第X期規劃數
        //sumC_item5 大型能管系統累計申請數
        //sumF_item5 大型能管系統累計完成數
    }

    //撈LOG 計畫變更紀錄
    public DataTable getReportLog()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
            select L_Person,L_IP,convert(nvarchar,L_ModDate,120) as L_ModDate,M_Name
            from LogTable
            left join Member on L_Person=M_Guid
            where L_Type=@strLType
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();

        oCmd.Parameters.AddWithValue("@strLType", strLType);
        oda.Fill(dt);
        return dt;
    }

    //撈當期累計執行進度 報表
    public DataTable getReportProcess()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
create table #tmpcity(
	city_Item nvarchar(50),
	city_Item_cn nvarchar(50),
	city_I_Guid nvarchar(50),
	RS_Guid nvarchar(50),
	city_Stage nvarchar(5),
	city_Year nvarchar(5),
	city_Season nvarchar(5)
)
create table #tmpA(
	cityno nvarchar(50),
	citycn nvarchar(50),
	I_Guid nvarchar(50),
	RS_Guid nvarchar(50),
	RS_Stage nvarchar(5),
	RS_Year nvarchar(5),
	RS_Season nvarchar(5),
	RS_Sum01S nvarchar(50),--預定
	RS_Sum01F nvarchar(50),--實際
	RS_Sum01S_F nvarchar(50),--預定-實際
	RS_Sum02S nvarchar(50),--預定
	RS_Sum02F nvarchar(50),--實際
	RS_Sum02S_F nvarchar(50),--預定-實際
	RS_Sum03S nvarchar(50),--預定
	RS_Sum03F nvarchar(50),--實際
	RS_Sum03S_F nvarchar(50),--預定-實際
	RS_Sum04S nvarchar(50),--預定
	RS_Sum04F nvarchar(50),--實際
	RS_Sum04S_F nvarchar(50)--預定-實際
)

--RS_01Schedule RS_01RealSchedule
--declare @strStage nvarchar(2)='1'
insert into #tmpcity(city_Item,city_Item_cn,city_I_Guid,RS_Guid,city_Stage,city_Year,city_Season)
select C_Item,C_Item_cn,I_Guid,RS_Guid,RS_Stage,RS_Year,RS_Season
from(
    select C_Item,C_Item_cn,I_Guid,RS_Guid,RS_Stage,RS_Year,RS_Season
    from CodeTable 
    left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
    left join ReportSeason on I_Guid=RS_PorjectGuid
left join ReportCheck on RC_ReportGuid=RS_Guid and RC_ReportType='02' and RC_CheckType='Y' and RC_Status='A'
    where C_Group='02' and RS_Stage=@strStage and RC_CheckType='Y' and RC_Status='A'
)#tmp
group by C_Item,C_Item_cn,I_Guid,RS_Guid,RS_Stage,RS_Year,RS_Season

--節電基礎工作
insert into #tmpA(cityno,citycn,I_Guid,RS_Guid,RS_Stage,RS_Year,RS_Season,RS_Sum01S,RS_Sum01F,RS_Sum01S_F,RS_Sum02S,RS_Sum02F,RS_Sum02S_F,RS_Sum03S,RS_Sum03F,RS_Sum03S_F,RS_Sum04S,RS_Sum04F,RS_Sum04S_F)
select C_Item,C_Item_cn,I_Guid,RS_Guid,RS_Stage,RS_Year,RS_Season
,isnull(RS_01Schedule,'0') RS_01Schedule
,isnull(RS_01RealSchedule,'0') RS_01RealSchedule
,convert(float,isnull(RS_01Schedule,0))-convert(float,isnull(RS_01RealSchedule,0)) RS01_S_F
,isnull(RS_02Schedule,'0') RS_02Schedule
,isnull(RS_02RealSchedule,'0') RS_01RealSchedule
,convert(float,isnull(RS_02Schedule,0))-convert(float,isnull(RS_02RealSchedule,0)) RS02_S_F
,isnull(RS_03Schedule,'0') RS_03Schedule
,isnull(RS_03RealSchedule,'0') RS_03RealSchedule
,convert(float,isnull(RS_03Schedule,'0'))-convert(float,isnull(RS_03RealSchedule,'0')) RS03_S_F
,isnull(RS_04Schedule,'0') RS_03Schedule
,isnull(RS_04RealSchedule,'0') RS_03RealSchedule
,convert(float,isnull(RS_04Schedule,'0'))-convert(float,isnull(RS_04RealSchedule,'0')) RS04_S_F
from  CodeTable a
left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
left join ReportSeason on I_Guid=RS_PorjectGuid and RS_Stage=@strStage
left join ReportCheck on RC_ReportGuid=RS_Guid and RC_ReportType='02' and RC_CheckType='Y' and RC_Status='A'
where C_Group='02' and RC_CheckType='Y' and RC_Status='A'

select * from #tmpA where RS_Guid is not null

drop table #tmpA
drop table #tmpcity
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();

        oCmd.Parameters.AddWithValue("@strStage", strStage);
        oda.Fill(dt);
        return dt;
    }

    //撈當期累計執行進度 報表
    public DataTable getReportReal()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
create table #tmpcity(
	city_Item nvarchar(50),
	city_Item_cn nvarchar(50),
	city_I_Guid nvarchar(50),
	city_RS_ReportGuid nvarchar(50),
	city_Stage nvarchar(5),
	city_Year nvarchar(5),
	city_Season nvarchar(5)
)
create table #tmpA(
	cityno nvarchar(50),
	citycn nvarchar(50),
	I_Guid nvarchar(50),
	RS_Guid nvarchar(50),
	RS_Stage nvarchar(5),
	RS_Year nvarchar(5),
	RS_Season nvarchar(5),
    RS_CostDesc nvarchar(max),
	RS_Type01Money nvarchar(50),--經費
	RS_Type01Real nvarchar(50),--實支數
	RS_Type01RealRate nvarchar(50),--實支率
	RS_Type02Money nvarchar(50),--經費
	RS_Type02Real nvarchar(50),--實支數
	RS_Type02RealRate nvarchar(50),--實支率
	RS_Type03Money nvarchar(50),--經費
	RS_Type03Real nvarchar(50),--實支率
	RS_Type03RealRate nvarchar(50),--實支率
	RS_Type04Money nvarchar(50),--經費
	RS_Type04Real nvarchar(50),--實支數
	RS_Type04RealRate nvarchar(50),--實支率
	RS_sumMoney nvarchar(50),--
	RS_sumReal nvarchar(50)--
)

--RS_01Schedule RS_01RealSchedule
--declare @strStage nvarchar(2)='1'
insert into #tmpcity(city_Item,city_Item_cn,city_I_Guid,city_RS_ReportGuid,city_Stage,city_Year,city_Season)
select C_Item,C_Item_cn,I_Guid,RS_Guid,RS_Stage,RS_Year,RS_Season
from(
	select C_Item,C_Item_cn,I_Guid,RS_Guid,RS_Stage,RS_Year,RS_Season
	from CodeTable 
	left join ProjectInfo on C_Item=I_City and I_Status='A'
	left join ReportSeason on I_Guid=RS_PorjectGuid
	where C_Group='02' and RS_Stage=@strStage
)#tmp
group by C_Item,C_Item_cn,I_Guid,RS_Guid,RS_Stage,RS_Year,RS_Season

--節電基礎工作
insert into #tmpA(cityno,citycn,I_Guid,RS_Guid,RS_Stage,RS_Year,RS_Season,RS_CostDesc,RS_Type01Money,RS_Type01Real,RS_Type01RealRate,RS_Type02Money,RS_Type02Real,RS_Type02RealRate,RS_Type03Money,RS_Type03Real,RS_Type03RealRate,RS_Type04Money,RS_Type04Real,RS_Type04RealRate,RS_sumMoney,RS_sumReal)
select C_Item,C_Item_cn,I_Guid,RS_Guid,RS_Stage,RS_Year,RS_Season,RS_CostDesc
,isnull(RS_Type01Money,0),isnull(RS_Type01Real,0),isnull(RS_Type01RealRate,0)
,isnull(RS_Type02Money,0),isnull(RS_Type02Real,0),isnull(RS_Type02RealRate,0)
,isnull(RS_Type03Money,0),isnull(RS_Type03Real,0),isnull(RS_Type03RealRate,0)
,isnull(RS_Type04Money,0),isnull(RS_Type04Real,0),isnull(RS_Type04RealRate,0)
,isnull(RS_Type01Money,0)+isnull(RS_Type02Money,0)+isnull(RS_Type03Money,0)+isnull(RS_Type04Money,0) as RS_SumMoney
,isnull(RS_Type01Real,0)+isnull(RS_Type02Real,0)+isnull(RS_Type03Real,0)+isnull(RS_Type04Real,0) as RS_SumReal
from  CodeTable a
left join ProjectInfo on C_Item=I_City and I_Status='A'
left join ReportSeason on I_Guid=RS_PorjectGuid and RS_Stage=@strStage
where C_Group='02'

select * from #tmpA where RS_Guid is not null

drop table #tmpA
drop table #tmpcity
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();

        oCmd.Parameters.AddWithValue("@strStage", strStage);
        oda.Fill(dt);
        return dt;
    }

    //管理員總表 - 各縣市申請數(擴大補助)
    public DataSet getReportTotalBehindgForEx()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();
        
        #region 201808新code
        sb.Append(@"
create table #tmpcity(
	city_Item nvarchar(50),
	city_Item_cn nvarchar(50),
	city_I_Guid nvarchar(50),
	city_Stage nvarchar(5),
	city_Year nvarchar(5),
	city_Season nvarchar(5)
)

---------------- Table 1 撈出codeTable裡面的擴大補助項目----------------

if @strExType =''--全部
    begin
        select C_Group,C_Item_cn,C_Item from CodeTable where C_Group='09' and C_Item<>'99' order by C_Item
    end
if @strExType ='1'--服務業(機關、學校)
    begin
        select C_Group,C_Item_cn,C_Item from CodeTable 
        where C_Group='09' and C_Item<>'99' 
            and C_Item in ('01','02','03','04','05','07','11','18','20','21','22','23','24','25','26','27','28','29','30','31','32')
        order by C_Item
    end
if @strExType ='2'--住宅
    begin
        select C_Group,C_Item_cn,C_Item from CodeTable 
        where C_Group='09' and C_Item<>'99' 
            and C_Item in ('03','04','05','06','08','09','10','12','13','14','15','16','17','19','22','25')
        order by C_Item
    end


-------------------------END---------------------------------------


------------------撈出當期底下有季報紀錄的縣市 insert到#tmpcity----------------
if @strExType =''--全部
    begin
        insert into #tmpcity(city_Item,city_Item_cn,city_I_Guid,city_Stage,city_Year,city_Season)
        select C_Item,C_Item_cn,I_Guid,RM_Stage,RM_Year,RM_Season
        from(
            select C_Item,C_Item_cn,I_Guid,@strStage as RM_Stage,RM_Year
            ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
	            when '04' then '2' when '05' then '2' when '06' then '2'  
	            when '07' then '3' when '08' then '3' when '09' then '3'  
	            when '10' then '4' when '11' then '4' when '12' then '4'  
	            end as RM_Season
            from CodeTable 
            left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
            left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='02'
            left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
            where C_Group='02' --and RC_Status='A' and RC_CheckType='Y'
        )#tmp
        group by C_Item,C_Item_cn,I_Guid,RM_Stage,RM_Year,RM_Season
    end
if @strExType ='1'--服務業(機關、學校)
    begin
        insert into #tmpcity(city_Item,city_Item_cn,city_I_Guid,city_Stage,city_Year,city_Season)
        select C_Item,C_Item_cn,I_Guid,RM_Stage,RM_Year,RM_Season
        from(
            select C_Item,C_Item_cn,I_Guid,@strStage as RM_Stage,RM_Year
            ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
	            when '04' then '2' when '05' then '2' when '06' then '2'  
	            when '07' then '3' when '08' then '3' when '09' then '3'  
	            when '10' then '4' when '11' then '4' when '12' then '4'  
	            end as RM_Season
            from CodeTable 
            left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
            left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='02'
            left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
            where C_Group='02' --and RC_Status='A' and RC_CheckType='Y'
                    and C_Item in ('01','02','03','04','05','07','11','18','20','21','22','23','24','25','26','27','28','29','30','31','32')
        )#tmp
        group by C_Item,C_Item_cn,I_Guid,RM_Stage,RM_Year,RM_Season
    end
if @strExType ='2'--住宅
    begin
        insert into #tmpcity(city_Item,city_Item_cn,city_I_Guid,city_Stage,city_Year,city_Season)
        select C_Item,C_Item_cn,I_Guid,RM_Stage,RM_Year,RM_Season
        from(
            select C_Item,C_Item_cn,I_Guid,@strStage as RM_Stage,RM_Year
            ,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
	            when '04' then '2' when '05' then '2' when '06' then '2'  
	            when '07' then '3' when '08' then '3' when '09' then '3'  
	            when '10' then '4' when '11' then '4' when '12' then '4'  
	            end as RM_Season
            from CodeTable 
            left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
            left join ReportMonth on I_Guid=RM_ProjectGuid and RM_Stage=@strStage and RM_ReportType='02'
            left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RC_Status='A' and RC_CheckType='Y'
            where C_Group='02' --and RC_Status='A' and RC_CheckType='Y'
                and C_Item in ('03','04','05','06','08','09','10','12','13','14','15','16','17','19','22','25')
        )#tmp
        group by C_Item,C_Item_cn,I_Guid,RM_Stage,RM_Year,RM_Season
    end
-------------------------END---------------------------------------



---------------- 撈出月報裡面當期各季的資料 insert 到#tmpVal----------------
select RM_ProjectGuid,RM_Stage,RM_Year,RM_Season,RM_Planning,SUM(isnull(RM_Finish,0)) sumFinsh,SUM(isnull(RM_Finish01,0)) sumFinsh01,RM_CPType
,SUM(isnull(RM_Type1ValueSum,0)) sum1,SUM(isnull(RM_Type2ValueSum,0)) sum2,SUM(isnull(RM_Type3ValueSum,0)) sum3,SUM(isnull(RM_Type4ValueSum,0)) sum4
into #tmpVal
from (
	select RM_ProjectGuid,RM_Stage,RM_Year,RM_Planning,RM_Finish,RM_Finish01,RM_CPType
	,RM_Type1ValueSum,RM_Type2ValueSum,RM_Type3ValueSum,RM_Type4ValueSum
	,case RM_Month when '01' then '1' when '02' then '1' when '03' then '1'
		when '04' then '2' when '05' then '2' when '06' then '2'  
		when '07' then '3' when '08' then '3' when '09' then '3'  
		when '10' then '4' when '11' then '4' when '12' then '4'  
		end as RM_Season
		 from ReportMonth 
	where RM_ReportType='02' and RM_Stage=@strStage
)tmp
group by RM_ProjectGuid,RM_Stage,RM_Year,RM_Planning,RM_Season,RM_CPType
-------------------------END---------------------------------------

-------------------------Table 2 把兩張表join起來---------------------------------------
select a.*,b.* 
from #tmpcity a left join #tmpVal b
on a.city_I_Guid=b.RM_ProjectGuid and a.city_Stage=b.RM_Stage and a.city_Year=b.RM_Year and a.city_Season=b.RM_Season

drop table #tmpcity
drop table #tmpVal


        ");
        #endregion
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@strStage", strStage);
        oCmd.Parameters.AddWithValue("@strExType", strExType);
        oda.Fill(ds);
        return ds;
    }

    //管理員總表 - 各縣市申請數 月累計(擴大補助)
    public DataSet getReportTotalBehindByMForEx()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        #region 201808新code
        sb.Append(@"
create table #tmpcity(
	city_Item nvarchar(50),
	city_Item_cn nvarchar(50),
	city_I_Guid nvarchar(50),
	city_Stage nvarchar(5),
	city_Year nvarchar(5),
	city_Month nvarchar(5),
	city_maxmonth nvarchar(6)
);

---------------- Table 1 撈出codeTable裡面的擴大補助項目----------------

if @strExType =''--全部
    begin
        select C_Group,C_Item_cn,C_Item from CodeTable where C_Group='09' and C_Item<>'99' order by C_Item
    end
if @strExType ='1'--服務業(機關、學校)
    begin
        select C_Group,C_Item_cn,C_Item from CodeTable 
        where C_Group='09' and C_Item<>'99' 
            and C_Item in ('01','02','03','04','05','07','11','18','20','21','22','23','24','25','26','27','28','29','30','31','32')
        order by C_Item
    end
if @strExType ='2'--住宅
    begin
        select C_Group,C_Item_cn,C_Item from CodeTable 
        where C_Group='09' and C_Item<>'99' 
            and C_Item in ('03','04','05','06','08','09','10','12','13','14','15','16','17','19','22','25')
        order by C_Item
    end

-------------------------END---------------------------------------


------------------撈出當期底下有季報紀錄的縣市 insert到#tmpcity----------------

--先撈出各縣市代碼 名稱 計畫GUID
insert into #tmpcity(city_Item,city_Item_cn,city_I_Guid,city_Stage,city_Year,city_Month,city_maxmonth)
select C_Item,C_Item_cn,I_Guid,'','','',''
from CodeTable 
left join ProjectInfo on C_Item=I_City and I_Status='A' and I_Flag='Y'
left join ReportMonth on I_Guid=RM_ProjectGuid and RM_ReportType='02'
where C_Group='02'  --and RM_ProjectGuid is not null
group by C_Item,C_Item_cn,I_Guid,RM_Stage


--根據計畫GUID撈出該期底下已審核通過的月報最大的年月份
select #tmp.RM_ProjectGuid,b.RM_ReportGuid,#tmp.maxmonth,b.RM_Stage,b.RM_Year,b.RM_Month into #tmptmp
from(
	select RM_ProjectGuid,MAX(CONVERT(int,isnull(RM_Year,'0')+isnull(RM_Month,'0'))) maxmonth 
	from ReportMonth 
	left join ReportCheck on RM_ReportGuid=RC_ReportGuid
	where RM_Stage=@strStage and RC_CheckType ='Y' and RC_Status='A' and RM_ReportType='02'
	group by RM_ProjectGuid
)#tmp
left join ReportMonth b on 
#tmp.RM_ProjectGuid=b.RM_ProjectGuid and b.RM_Stage=@strStage and b.RM_ReportType='02'
and #tmp.maxmonth=CONVERT(int,(isnull(b.RM_Year,'0')+isnull(b.RM_Month,'0')))
where  b.RM_ReportType='02'
group by #tmp.RM_ProjectGuid,b.RM_ReportGuid,#tmp.maxmonth,b.RM_Year,b.RM_Month,b.RM_Stage

-- 將已審核最大年月資料update 回#tmpcity
update #tmpcity set city_Year=b.RM_Year,city_Month=b.RM_Month,city_Stage=b.RM_Stage,city_maxmonth=b.maxmonth
from
(
	select RM_ProjectGuid,RM_Year,RM_Month,RM_Stage,maxmonth from #tmptmp
)b
where city_I_Guid = b.RM_ProjectGuid
-------------------------END---------------------------------------


-------------------------撈出月報裡面當期各月的資料 insert 到#tmpVal-------------------------
--1.先找出小於當期最大年月份的所有擴大補助月報
--2.group by 算出總和
--3.再根據當其最大年月抓出該最大年月份的個項目的規畫數(因為有可能隨時去更胎計畫資料 所以以最新的為主)

if @strExType =''--全部
    begin
        select maxmonth,RM_ProjectGuid,RM_CPType,RM_Stage
	        ,sum(isnull(RM_Type1ValueSum,0)) as sum1,sum(isnull(RM_Type2ValueSum,0)) as sum2
	        ,sum(isnull(RM_Type3ValueSum,0)) as sum3,sum(isnull(RM_Type4ValueSum,0)) as sum4
	        ,sum(isnull(RM_Finish,0)) as sumFinsh, sum(isnull(RM_Finish01,0)) as sumFinsh01
	        ,(select TOP 1 c.RM_Planning from ReportMonth c where c.RM_Status='A' and c.RM_ReportType='02' and c.RM_CPType=tmp.RM_CPType and CONVERT(int,(isnull(c.RM_Year,'0')+isnull(c.RM_Month,'0')))=CONVERT(int,tmp.maxmonth) and c.RM_ProjectGuid=tmp.RM_ProjectGuid) as RM_Planning
        into #tmpVal
        from 
        (
	        select a.maxmonth,b.RM_ProjectGuid,b.RM_ReportGuid,b.RM_CPType,b.RM_Stage,b.RM_Year,b.RM_Month
		        ,b.RM_Type1ValueSum,b.RM_Type2ValueSum,b.RM_Type3ValueSum,b.RM_Type4ValueSum
		        ,b.RM_Finish,b.RM_Finish01
	        from #tmptmp a
	        left join ReportMonth b on a.RM_ProjectGuid=b.RM_ProjectGuid and a.RM_Stage=b.RM_Stage  and a.RM_ReportGuid=b.RM_ReportGuid
	        where CONVERT(int,(isnull(b.RM_Year,'0')+isnull(b.RM_Month,'0')))>=CONVERT(int,a.maxmonth) and b.RM_Stage=@strStage and b.RM_Status='A'
        )tmp
        group by maxmonth,RM_ProjectGuid,RM_CPType,RM_CPType,RM_Stage

		-------------------------Table 2 把兩張表join起來---------------------------------------
		select a.*,b.RM_CPType,b.sum1,b.sum2,b.sum3,b.sum4,b.sumFinsh,b.sumFinsh01,b.RM_Planning from #tmpcity a
		left join #tmpVal b on a.city_I_Guid=b.RM_ProjectGuid and a.city_Stage=b.RM_Stage and a.city_maxmonth=b.maxmonth
		order by city_Item asc

		drop table #tmpVal
	end

if @strExType ='1'--服務業(機關、學校)
    begin
        select maxmonth,RM_ProjectGuid,RM_CPType,RM_Stage
	        ,sum(isnull(RM_Type1ValueSum,0)) as sum1,sum(isnull(RM_Type2ValueSum,0)) as sum2
	        ,sum(isnull(RM_Type3ValueSum,0)) as sum3,sum(isnull(RM_Type4ValueSum,0)) as sum4
	        ,sum(isnull(RM_Finish,0)) as sumFinsh, sum(isnull(RM_Finish01,0)) as sumFinsh01
	        ,(select TOP 1 c.RM_Planning from ReportMonth c where c.RM_Status='A' and c.RM_ReportType='02' and c.RM_CPType=tmp.RM_CPType and CONVERT(int,(isnull(c.RM_Year,'0')+isnull(c.RM_Month,'0')))=CONVERT(int,tmp.maxmonth) and c.RM_ProjectGuid=tmp.RM_ProjectGuid) as RM_Planning
        into #tmpVal_1
        from 
        (
	        select a.maxmonth,b.RM_ProjectGuid,b.RM_ReportGuid,b.RM_CPType,b.RM_Stage,b.RM_Year,b.RM_Month
		        ,b.RM_Type1ValueSum,b.RM_Type2ValueSum,b.RM_Type3ValueSum,b.RM_Type4ValueSum
		        ,b.RM_Finish,b.RM_Finish01
	        from #tmptmp a
	        left join ReportMonth b on a.RM_ProjectGuid=b.RM_ProjectGuid and a.RM_Stage=b.RM_Stage  and a.RM_ReportGuid=b.RM_ReportGuid
	        where CONVERT(int,(isnull(b.RM_Year,'0')+isnull(b.RM_Month,'0')))>=CONVERT(int,a.maxmonth) and b.RM_Stage=@strStage and b.RM_Status='A'
                and b.RM_CPType in ('01','02','03','04','05','07','11','18','20','21','22','23','24','25','26','27','28','29','30','31','32')
        )tmp
        group by maxmonth,RM_ProjectGuid,RM_CPType,RM_CPType,RM_Stage

		-------------------------Table 2 把兩張表join起來---------------------------------------
		select a.*,b.RM_CPType,b.sum1,b.sum2,b.sum3,b.sum4,b.sumFinsh,b.sumFinsh01,b.RM_Planning from #tmpcity a
		left join #tmpVal_1 b on a.city_I_Guid=b.RM_ProjectGuid and a.city_Stage=b.RM_Stage and a.city_maxmonth=b.maxmonth
		order by city_Item asc

		drop table #tmpVal_1
    end

if @strExType ='2'--住宅
    begin
        select maxmonth,RM_ProjectGuid,RM_CPType,RM_Stage
	        ,sum(isnull(RM_Type1ValueSum,0)) as sum1,sum(isnull(RM_Type2ValueSum,0)) as sum2
	        ,sum(isnull(RM_Type3ValueSum,0)) as sum3,sum(isnull(RM_Type4ValueSum,0)) as sum4
	        ,sum(isnull(RM_Finish,0)) as sumFinsh, sum(isnull(RM_Finish01,0)) as sumFinsh01
	        ,(select TOP 1 c.RM_Planning from ReportMonth c where c.RM_Status='A' and c.RM_ReportType='02' and c.RM_CPType=tmp.RM_CPType and CONVERT(int,(isnull(c.RM_Year,'0')+isnull(c.RM_Month,'0')))=CONVERT(int,tmp.maxmonth) and c.RM_ProjectGuid=tmp.RM_ProjectGuid) as RM_Planning
        into #tmpVal_2
        from 
        (
	        select a.maxmonth,b.RM_ProjectGuid,b.RM_ReportGuid,b.RM_CPType,b.RM_Stage,b.RM_Year,b.RM_Month
		        ,b.RM_Type1ValueSum,b.RM_Type2ValueSum,b.RM_Type3ValueSum,b.RM_Type4ValueSum
		        ,b.RM_Finish,b.RM_Finish01
	        from #tmptmp a
	        left join ReportMonth b on a.RM_ProjectGuid=b.RM_ProjectGuid and a.RM_Stage=b.RM_Stage  and a.RM_ReportGuid=b.RM_ReportGuid
	        where CONVERT(int,(isnull(b.RM_Year,'0')+isnull(b.RM_Month,'0')))>=CONVERT(int,a.maxmonth) and b.RM_Stage=@strStage and b.RM_Status='A'
                and b.RM_CPType in ('03','04','05','06','08','09','10','12','13','14','15','16','17','19','22','25')
        )tmp
        group by maxmonth,RM_ProjectGuid,RM_CPType,RM_CPType,RM_Stage

		-------------------------Table 2 把兩張表join起來---------------------------------------
		select a.*,b.RM_CPType,b.sum1,b.sum2,b.sum3,b.sum4,b.sumFinsh,b.sumFinsh01,b.RM_Planning from #tmpcity a
		left join #tmpVal_2 b on a.city_I_Guid=b.RM_ProjectGuid and a.city_Stage=b.RM_Stage and a.city_maxmonth=b.maxmonth
		order by city_Item asc

		drop table #tmpVal_2

    end
-------------------------END-------------------------

drop table #tmpcity
drop table #tmptmp

        ");
        #endregion
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@strStage", strStage);
        oCmd.Parameters.AddWithValue("@strExType", strExType);
        oda.Fill(ds);
        return ds;
    }

}