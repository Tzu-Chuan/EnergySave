using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;

/// <summary>
/// ReportMonth_DB 的摘要描述
/// </summary>
public class ReportMonth_DB
{
    public ReportMonth_DB()
    {
        //
        // TODO: 在這裡新增建構函式邏輯
        //
    }

    #region 全私用
    string M_ID = string.Empty;
    string str_stage = string.Empty;
    string I_Guid = string.Empty;
    string P_Period = string.Empty;
    #endregion
    #region 全公用
    public string _M_ID
    {
        set { M_ID = value; }
    }
    public string _str_stage
    {
        set { str_stage = value; }
    }
    public string _I_Guid
    {
        set { I_Guid = value; }
    }
    public string _P_Period
    {
        set { P_Period = value; }
    }
    #endregion

    #region 私用
    string RM_ID = string.Empty;
    string RM_Guid = string.Empty;
    string RM_ReportGuid = string.Empty;
    string RM_ProjectGuid = string.Empty;
    string RM_PGuid = string.Empty;
    string RM_Stage = string.Empty;
    string RM_Year = string.Empty;
    string RM_Month = string.Empty;
    string RM_CPType = string.Empty;
    Decimal RM_Finish;
    Int32 RM_Finish01;//無風管的累計完成數(台)
    Decimal RM_Planning;
    Decimal RM_Type1Value1;
    Decimal RM_Type1Value2;
    Decimal RM_Type1Value3;
    Decimal RM_Type1ValueSum;
    Decimal RM_Type2Value1;
    Decimal RM_Type2Value2;
    Decimal RM_Type2Value3;
    Decimal RM_Type2ValueSum;
    Decimal RM_Type3Value1;
    Decimal RM_Type3Value2;
    Decimal RM_Type3Value3;
    Decimal RM_Type3ValueSum;
    Decimal RM_Type4Value1;
    Decimal RM_Type4Value2;
    Decimal RM_Type4Value3;
    Decimal RM_Type4ValueSum;
    Decimal RM_PreVal;
    Decimal RM_ChkVal;
    Decimal RM_NotChkVal;
    string RM_Remark = string.Empty;
    string RM_Formula = string.Empty;
    DateTime RM_CreateDate;
    string RM_CreateId = string.Empty;
    DateTime RM_ModDate;
    string RM_ModId = string.Empty;
    string RM_Status = string.Empty;
    string RM_ReportType = string.Empty;
    string RM_P_ExType = string.Empty;
    string RM_P_ExDeviceType = string.Empty;
    #endregion
    #region 公用
    public string _RM_ID
    {
        set { RM_ID = value; }
    }
    public string _RM_Guid
    {
        set { RM_Guid = value; }
    }
    public string _RM_ReportGuid
    {
        set { RM_ReportGuid = value; }
    }
    public string _RM_ProjectGuid
    {
        set { RM_ProjectGuid = value; }
    }
    public string _RM_PGuid
    {
        set { RM_PGuid = value; }
    }
    public string _RM_Stage
    {
        set { RM_Stage = value; }
    }
    public string _RM_Year
    {
        set { RM_Year = value; }
    }
    public string _RM_Month
    {
        set { RM_Month = value; }
    }
    public string _RM_CPType
    {
        set { RM_CPType = value; }
    }
    public Decimal _RM_Finish
    {
        set { RM_Finish = value; }
    }
    public Int32 _RM_Finish01
    {
        set { RM_Finish01 = value; }
    }
    public Decimal _RM_Planning
    {
        set { RM_Planning = value; }
    }
    public Decimal _RM_Type1Value1
    {
        set { RM_Type1Value1 = value; }
    }
    public Decimal _RM_Type1Value2
    {
        set { RM_Type1Value2 = value; }
    }
    public Decimal _RM_Type1Value3
    {
        set { RM_Type1Value3 = value; }
    }
    public Decimal _RM_Type1ValueSum
    {
        set { RM_Type1ValueSum = value; }
    }
    public Decimal _RM_Type2Value1
    {
        set { RM_Type2Value1 = value; }
    }
    public Decimal _RM_Type2Value2
    {
        set { RM_Type2Value2 = value; }
    }
    public Decimal _RM_Type2Value3
    {
        set { RM_Type2Value3 = value; }
    }
    public Decimal _RM_Type2ValueSum
    {
        set { RM_Type2ValueSum = value; }
    }
    public Decimal _RM_Type3Value1
    {
        set { RM_Type3Value1 = value; }
    }
    public Decimal _RM_Type3Value2
    {
        set { RM_Type3Value2 = value; }
    }
    public Decimal _RM_Type3Value3
    {
        set { RM_Type3Value3 = value; }
    }
    public Decimal _RM_Type3ValueSum
    {
        set { RM_Type3ValueSum = value; }
    }
    public Decimal _RM_Type4Value1
    {
        set { RM_Type4Value1 = value; }
    }
    public Decimal _RM_Type4Value2
    {
        set { RM_Type4Value2 = value; }
    }
    public Decimal _RM_Type4Value3
    {
        set { RM_Type4Value3 = value; }
    }
    public Decimal _RM_Type4ValueSum
    {
        set { RM_Type4ValueSum = value; }
    }
    public Decimal _RM_PreVal
    {
        set { RM_PreVal = value; }
    }
    public Decimal _RM_ChkVal
    {
        set { RM_ChkVal = value; }
    }
    public Decimal _RM_NotChkVal
    {
        set { RM_NotChkVal = value; }
    }
    public string _RM_Remark
    {
        set { RM_Remark = value; }
    }
    public DateTime _RM_CreateDate
    {
        set { RM_CreateDate = value; }
    }
    public string _RM_CreateId
    {
        set { RM_CreateId = value; }
    }
    public DateTime _RM_ModDate
    {
        set { RM_ModDate = value; }
    }
    public string _RM_ModId
    {
        set { RM_ModId = value; }
    }
    public string _RM_Status
    {
        set { RM_Status = value; }
    }
    public string _RM_ReportType
    {
        set { RM_ReportType = value; }
    }
    public string _RM_P_ExType
    {
        set { RM_P_ExType = value; }
    }
    public string _RM_P_ExDeviceType
    {
        set { RM_P_ExDeviceType = value; }
    }
    public string _RM_Formula
    {
        set { RM_Formula = value; }
    }
    #endregion

    //撈承辦人資訊
    public DataTable selectMemberById()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
            select a.*,b.C_Item_cn 
            from Member a 
            left join CodeTable b on a.M_City = b.C_Item and C_Group='02'
            where a.M_ID=@M_ID 
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oCmd.Parameters.AddWithValue("@M_ID", M_ID);
        oda.Fill(ds);
        return ds;
    }

    //撈期程資料
    public DataTable selectStageDate()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
            select 
        ");
        if (str_stage=="1") {
            sb.Append(@"
                I_1_Sdate as sdate,I_1_Edate as edate
            ");
        }
        if (str_stage == "2")
        {
            sb.Append(@"
                I_2_Sdate as sdate,I_2_Edate as edate
            ");
        }
        if (str_stage == "3")
        {
            sb.Append(@"
                I_3_Sdate as sdate,I_3_Edate as edate
            ");
        }
        sb.Append(@"
            from ProjectInfo
            where I_Guid=@I_Guid 
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oCmd.Parameters.AddWithValue("@I_Guid", I_Guid);
        oCmd.Parameters.AddWithValue("@str_stage", str_stage);
        oda.Fill(ds);
        return ds;
    }

    //篩選月報資料
    public DataSet selectMonthReportBefore()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();
        
        sb.Append(@"
            
            select * from ReportMonth where RM_ProjectGuid=@I_Guid and RM_Stage=@P_Period and RM_Year=@RM_Year and RM_Month=@RM_Month and RM_Status<>'D' and RM_ReportType=@RM_ReportType
            
            select a.P_Guid,a.P_ParentId,a.P_Type,a.P_ItemName,a.P_Period,c.C_Item_cn ,a.P_Status,b.*,d.M_Name,d.M_Tel,d.M_Manager_ID,e.M_Name,f.RC_CheckType,f.RC_Status
                from PushItem a
                left join ReportMonth b on a.P_ParentId=b.RM_ProjectGuid  and a.P_ItemName=b.RM_CPType and b.RM_Stage=@P_Period and b.RM_Year=@RM_Year and b.RM_Month=@RM_Month
                left join CodeTable c on c.C_Group='07' and a.P_ItemName = c.C_Item
                left join Member d on d.M_ID=@M_ID
                left join Member e on d.M_id=@M_ID and d.M_Manager_ID = e.M_Guid
                left join ReportCheck f on b.RM_ReportGuid = f.RC_ReportGuid and RC_Status<>'D'
                where P_ParentId=@I_Guid and P_Period=@P_Period and (b.RM_Status<>'D' or b.RM_Status is null) and RM_ReportType=@RM_ReportType
                
        ");
        if (RM_ReportType == "01")
        {
            sb.Append(@"
                and P_Type<>'04'
            ");
        }
        else {
            sb.Append(@"
                and P_Type='04'
            ");
        }
        
        sb.Append(@"order by P_ItemName asc");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();
        oCmd.Parameters.AddWithValue("@M_ID", M_ID);
        oCmd.Parameters.AddWithValue("@I_Guid", I_Guid);
        oCmd.Parameters.AddWithValue("@P_Period", P_Period);
        oCmd.Parameters.AddWithValue("@RM_Year", RM_Year);
        oCmd.Parameters.AddWithValue("@RM_Month", RM_Month);
        oCmd.Parameters.AddWithValue("@RM_ReportType", RM_ReportType);
        oda.Fill(ds);
        return ds;
    }

    //撈月報資料
    public DataSet selectMonthReport()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
            select a.M_ID,a.M_Name,a.M_Tel,a.M_Manager_ID as bossid,b.M_Name as bossname
            from Member a left join Member b on a.M_Manager_ID=b.M_Guid
            where a.M_ID=@M_ID

            declare @report_guid nvarchar(50);
            declare @checkflag nvarchar(2);
            set rowcount 1;
            select @report_guid = RM_ReportGuid from ReportMonth where RM_ProjectGuid=@I_Guid and RM_Stage=@P_Period and RM_Year=@RM_Year and RM_Month=@RM_Month
            set rowcount 0;
            select @checkflag = RC_CheckType from ReportCheck where RC_ReportGuid = @report_guid and RC_Status<>'D'

            if @checkflag<>'Y' or @checkflag is null
                begin
                    select #tmp.* ,
                    (select sum(isnull(RM_Type4ValueSum,0.0)) from ReportMonth h left join ReportCheck j on h.RM_Stage=j.RC_Stage and j.RC_ReportType='01' where h.RM_Stage=#tmp.P_Period and  h.RM_Year+h.RM_Month<@RM_Year+@RM_Month and h.RM_CPType=#tmp.P_ItemName and j.RC_Status='A' and j.RC_CheckType='Y' and h.RM_Status<>'D' and h.RM_ReportGuid=j.RC_ReportGuid and h.RM_ProjectGuid=@I_Guid  ) as countFinishKW,--累計完成數 KW(無風管)
                    (select sum(isnull(RM_Type3ValueSum,0)) from ReportMonth i left join ReportCheck k on i.RM_Stage=k.RC_Stage and k.RC_ReportType='01' where i.RM_Stage=#tmp.P_Period and  i.RM_Year+i.RM_Month<@RM_Year+@RM_Month and i.RM_CPType=#tmp.P_ItemName and k.RC_Status='A' and k.RC_CheckType='Y' and i.RM_Status<>'D' and i.RM_ReportGuid=k.RC_ReportGuid and i.RM_ProjectGuid=@I_Guid  ) as countFinish03,--累計完成數(老舊 停車場 中型 大型)
                    (select sum(isnull(RM_Type2ValueSum,0)) from ReportMonth m left join ReportCheck n on m.RM_Stage=n.RC_Stage and n.RC_ReportType='01' where m.RM_Stage=#tmp.P_Period and  m.RM_Year+m.RM_Month<@RM_Year+@RM_Month and m.RM_CPType=#tmp.P_ItemName and n.RC_Status='A' and n.RC_CheckType='Y' and m.RM_Status<>'D' and m.RM_ReportGuid=n.RC_ReportGuid and m.RM_ProjectGuid=@I_Guid  ) as countFinish02,--累計核定數(老舊 停車場 中型 大型)
                    (select sum(isnull(RM_Type3ValueSum,0.0)) from ReportMonth o left join ReportCheck p on o.RM_Stage=p.RC_Stage and p.RC_ReportType='01' where o.RM_Stage=#tmp.P_Period and  o.RM_Year+o.RM_Month<@RM_Year+@RM_Month and o.RM_CPType=#tmp.P_ItemName and p.RC_Status='A' and p.RC_CheckType='Y' and o.RM_Status<>'D' and o.RM_ReportGuid=p.RC_ReportGuid and o.RM_ProjectGuid=@I_Guid  ) as countApplyKW,--累計申請數 KW(無風管)
                    (select sum(isnull(RM_Type1ValueSum,0)) from ReportMonth q left join ReportCheck r on q.RM_Stage=r.RC_Stage and r.RC_ReportType='01' where q.RM_Stage=#tmp.P_Period and  q.RM_Year+q.RM_Month<@RM_Year+@RM_Month and q.RM_CPType=#tmp.P_ItemName and r.RC_Status='A' and r.RC_CheckType='Y' and q.RM_Status<>'D' and q.RM_ReportGuid=r.RC_ReportGuid and q.RM_ProjectGuid=@I_Guid  ) as countApply01--累計完成數(老舊 停車場 中型 大型)
                    from (
                        select a.P_Guid,a.P_ParentId,a.P_Type,a.P_ItemName,a.P_Period,c.C_Item_cn ,a.P_Status,b.*,d.M_Name as M_Name,d.M_Tel,d.M_Manager_ID,e.M_Name as M__Manage_Name,f.RC_CheckType,f.RC_Status,f.RC_CheckDate,
                                g.I_Finish_item1_1,g.I_Finish_item1_2,g.I_Finish_item1_3,g.I_Finish_item2_1,g.I_Finish_item2_2,g.I_Finish_item2_3,
                                g.I_Finish_item3_1,g.I_Finish_item3_2,g.I_Finish_item3_3,g.I_Finish_item4_1,g.I_Finish_item4_2,g.I_Finish_item4_3,
                                g.I_Finish_item5_1,g.I_Finish_item5_2,g.I_Finish_item5_3
                        from PushItem a
                        left join ReportMonth b on a.P_ParentId=b.RM_ProjectGuid  and a.P_ItemName=b.RM_CPType and b.RM_Stage=@P_Period and b.RM_Year=@RM_Year and b.RM_Month=@RM_Month and (b.RM_Status<>'D' or b.RM_Status is null) and b.RM_ReportType='01'
                        left join CodeTable c on c.C_Group='07' and a.P_ItemName = c.C_Item
                        left join Member d on d.M_ID=@M_ID
                        left join Member e on d.M_id=@M_ID and d.M_Manager_ID = e.M_Guid
                        left join ReportCheck f on b.RM_ReportGuid = f.RC_ReportGuid and RC_Status<>'D'
                        left join ProjectInfo g on g.I_Guid = @I_Guid and g.I_Status<>'D'
                        where P_ParentId=@I_Guid and P_Period=@P_Period  and a.P_Status<>'D' and P_Type='03' -- and RM_Year=@RM_Year and RM_Month=@RM_Month
                    )#tmp
                    order by #tmp.P_ItemName asc
                end
            else
                begin
                     select #tmp.* ,
                    (select sum(isnull(RM_Type4ValueSum,0.0)) from ReportMonth h left join ReportCheck j on h.RM_Stage=j.RC_Stage and j.RC_ReportType='01' where h.RM_Stage=#tmp.P_Period and  h.RM_Year+h.RM_Month<@RM_Year+@RM_Month and h.RM_CPType=#tmp.P_ItemName and j.RC_Status='A' and j.RC_CheckType='Y' and h.RM_Status<>'D' and h.RM_ReportGuid=j.RC_ReportGuid  and h.RM_ProjectGuid=@I_Guid ) as countFinishKW,--累計完成數 KW(無風管)
                    (select sum(isnull(RM_Type3ValueSum,0)) from ReportMonth i left join ReportCheck k on i.RM_Stage=k.RC_Stage and k.RC_ReportType='01' where i.RM_Stage=#tmp.P_Period and  i.RM_Year+i.RM_Month<@RM_Year+@RM_Month and i.RM_CPType=#tmp.P_ItemName and k.RC_Status='A' and k.RC_CheckType='Y' and i.RM_Status<>'D' and i.RM_ReportGuid=k.RC_ReportGuid  and i.RM_ProjectGuid=@I_Guid ) as countFinish03,--累計完成數(老舊 停車場 中型 大型)
                    (select sum(isnull(RM_Type2ValueSum,0)) from ReportMonth m left join ReportCheck n on m.RM_Stage=n.RC_Stage and n.RC_ReportType='01' where m.RM_Stage=#tmp.P_Period and  m.RM_Year+m.RM_Month<@RM_Year+@RM_Month and m.RM_CPType=#tmp.P_ItemName and n.RC_Status='A' and n.RC_CheckType='Y' and m.RM_Status<>'D' and m.RM_ReportGuid=n.RC_ReportGuid  and m.RM_ProjectGuid=@I_Guid ) as countFinish02,--累計核定數(老舊 停車場 中型 大型)
                    (select sum(isnull(RM_Type3ValueSum,0.0)) from ReportMonth o left join ReportCheck p on o.RM_Stage=p.RC_Stage and p.RC_ReportType='01' where o.RM_Stage=#tmp.P_Period and  o.RM_Year+o.RM_Month<@RM_Year+@RM_Month and o.RM_CPType=#tmp.P_ItemName and p.RC_Status='A' and p.RC_CheckType='Y' and o.RM_Status<>'D' and o.RM_ReportGuid=p.RC_ReportGuid and o.RM_ProjectGuid=@I_Guid  ) as countApplyKW,--累計申請數 KW(無風管)
                    (select sum(isnull(RM_Type1ValueSum,0)) from ReportMonth q left join ReportCheck r on q.RM_Stage=r.RC_Stage and r.RC_ReportType='01' where q.RM_Stage=#tmp.P_Period and  q.RM_Year+q.RM_Month<@RM_Year+@RM_Month and q.RM_CPType=#tmp.P_ItemName and r.RC_Status='A' and r.RC_CheckType='Y' and q.RM_Status<>'D' and q.RM_ReportGuid=r.RC_ReportGuid and q.RM_ProjectGuid=@I_Guid  ) as countApply01--累計完成數(老舊 停車場 中型 大型)
                    from (
                        select a.P_Guid,a.P_ParentId,a.P_Type,a.P_ItemName,a.P_Period,c.C_Item_cn ,a.P_Status,b.*,d.M_Name as M_Name,d.M_Tel,d.M_Manager_ID,e.M_Name as M_Manage_Name,f.RC_CheckType,f.RC_Status,f.RC_CheckDate,
                                g.I_Finish_item1_1,g.I_Finish_item1_2,g.I_Finish_item1_3,g.I_Finish_item2_1,g.I_Finish_item2_2,g.I_Finish_item2_3,
                                g.I_Finish_item3_1,g.I_Finish_item3_2,g.I_Finish_item3_3,g.I_Finish_item4_1,g.I_Finish_item4_2,g.I_Finish_item4_3,
                                g.I_Finish_item5_1,g.I_Finish_item5_2,g.I_Finish_item5_3
                        from PushItem a
                        left join ReportMonth b on a.P_ParentId=b.RM_ProjectGuid and a.P_ItemName=b.RM_CPType and a.P_Guid=b.RM_PGuid and b.RM_Stage=@P_Period and b.RM_Year=@RM_Year and b.RM_Month=@RM_Month and (b.RM_Status<>'D' or b.RM_Status is null) and b.RM_ReportType='01'
                        left join CodeTable c on c.C_Group='07' and a.P_ItemName = c.C_Item
                        left join Member d on d.M_ID=@M_ID
                        left join Member e on d.M_id=@M_ID and d.M_Manager_ID = e.M_Guid
                        left join ReportCheck f on b.RM_ReportGuid = f.RC_ReportGuid and RC_Status<>'D'
                        left join ProjectInfo g on g.I_Guid = @I_Guid and g.I_Status<>'D'
                        where P_ParentId=@I_Guid and P_Period=@P_Period and b.RM_Status<>'D' and P_Type='03'
                    )#tmp
                    order by #tmp.P_ItemName asc
                end
        ");
        
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();
        oCmd.Parameters.AddWithValue("@M_ID", M_ID);
        oCmd.Parameters.AddWithValue("@I_Guid", I_Guid);
        oCmd.Parameters.AddWithValue("@P_Period", P_Period); 
        oCmd.Parameters.AddWithValue("@RM_Year", RM_Year);
        oCmd.Parameters.AddWithValue("@RM_Month", RM_Month);
        oda.Fill(ds);
        return ds;
    }

    //修改月報資料
    public void modMonthReport(string modType)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        if (modType == "add")
        {//新增
            oCmd.CommandText = @"
                insert into ReportMonth(RM_Guid,RM_ReportGuid,RM_ProjectGuid,RM_Stage,RM_Year,RM_Month,RM_CPType,RM_Finish,RM_Finish01,RM_Planning
                    ,RM_Type1Value1,RM_Type1Value2,RM_Type1Value3,RM_Type1ValueSum,RM_Type2Value1,RM_Type2Value2,RM_Type2Value3,RM_Type2ValueSum
                    ,RM_Type3Value1,RM_Type3Value2,RM_Type3Value3,RM_Type3ValueSum,RM_Type4Value1,RM_Type4Value2,RM_Type4Value3,RM_Type4ValueSum
                    ,RM_PreVal,RM_ChkVal,RM_NotChkVal,RM_Remark,RM_CreateDate,RM_CreateId,RM_ModDate,RM_ModId,RM_Status,RM_PGuid,RM_ReportType,RM_P_ExType,RM_P_ExDeviceType,RM_Formula)
                values(newid(),@RM_ReportGuid,@RM_ProjectGuid,@RM_Stage,@RM_Year,@RM_Month,@RM_CPType,@RM_Finish,@RM_Finish01,@RM_Planning
                    ,@RM_Type1Value1,@RM_Type1Value2,@RM_Type1Value3,@RM_Type1ValueSum,@RM_Type2Value1,@RM_Type2Value2,@RM_Type2Value3,@RM_Type2ValueSum
                    ,@RM_Type3Value1,@RM_Type3Value2,@RM_Type3Value3,@RM_Type3ValueSum,@RM_Type4Value1,@RM_Type4Value2,@RM_Type4Value3,@RM_Type4ValueSum
                    ,@RM_PreVal,@RM_ChkVal,@RM_NotChkVal,@RM_Remark,@RM_ModDate,@RM_ModId,@RM_ModDate,@RM_ModId,'A',@RM_PGuid,@RM_ReportType,@RM_P_ExType,@RM_P_ExDeviceType,@RM_Formula)
            ";
            oCmd.Parameters.AddWithValue("@RM_ReportGuid", RM_ReportGuid);
        }
        else
        {//修改
            oCmd.CommandText = @"
                update ReportMonth set 
                RM_Finish=@RM_Finish,RM_Finish01=@RM_Finish01,RM_Planning=@RM_Planning
                ,RM_Type1Value1=@RM_Type1Value1,RM_Type1Value2=@RM_Type1Value2,RM_Type1Value3=@RM_Type1Value3,RM_Type1ValueSum=@RM_Type1ValueSum
                ,RM_Type2Value1=@RM_Type2Value1,RM_Type2Value2=@RM_Type2Value2,RM_Type2Value3=@RM_Type2Value3,RM_Type2ValueSum=@RM_Type2ValueSum
                ,RM_Type3Value1=@RM_Type3Value1,RM_Type3Value2=@RM_Type3Value2,RM_Type3Value3=@RM_Type3Value3,RM_Type3ValueSum=@RM_Type3ValueSum
                ,RM_Type4Value1=@RM_Type4Value1,RM_Type4Value2=@RM_Type4Value2,RM_Type4Value3=@RM_Type4Value3,RM_Type4ValueSum=@RM_Type4ValueSum
                ,RM_PreVal=@RM_PreVal,RM_ChkVal=@RM_ChkVal,RM_NotChkVal=@RM_NotChkVal,RM_Remark=@RM_Remark
                ,RM_ModId=@RM_ModId,RM_ModDate=@RM_ModDate,RM_Status='A',RM_Formula=@RM_Formula
                where RM_ProjectGuid=@RM_ProjectGuid and RM_Stage=@RM_Stage and RM_Year=@RM_Year and RM_Month=@RM_Month and RM_CPType=@RM_CPType and RM_P_ExType=@RM_P_ExType and RM_ReportType=@RM_ReportType and RM_PGuid=@RM_PGuid -- and RM_P_ExDeviceType=@RM_P_ExDeviceType
            ";
            
        }
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@RM_ProjectGuid", RM_ProjectGuid);
        oCmd.Parameters.AddWithValue("@RM_PGuid", RM_PGuid);
        oCmd.Parameters.AddWithValue("@RM_CPType", RM_CPType);
        oCmd.Parameters.AddWithValue("@RM_Stage", RM_Stage);
        oCmd.Parameters.AddWithValue("@RM_Year", RM_Year);
        oCmd.Parameters.AddWithValue("@RM_Month", RM_Month);
        oCmd.Parameters.AddWithValue("@RM_Finish", RM_Finish);
        oCmd.Parameters.AddWithValue("@RM_Finish01", RM_Finish01);
        oCmd.Parameters.AddWithValue("@RM_Planning", RM_Planning);
        oCmd.Parameters.AddWithValue("@RM_Type1Value1", RM_Type1Value1);
        oCmd.Parameters.AddWithValue("@RM_Type1Value2", RM_Type1Value2);
        oCmd.Parameters.AddWithValue("@RM_Type1Value3", RM_Type1Value3);
        oCmd.Parameters.AddWithValue("@RM_Type1ValueSum", RM_Type1ValueSum);
        oCmd.Parameters.AddWithValue("@RM_Type2Value1", RM_Type2Value1);
        oCmd.Parameters.AddWithValue("@RM_Type2Value2", RM_Type2Value2);
        oCmd.Parameters.AddWithValue("@RM_Type2Value3", RM_Type2Value3);
        oCmd.Parameters.AddWithValue("@RM_Type2ValueSum", RM_Type2ValueSum);
        oCmd.Parameters.AddWithValue("@RM_Type3Value1", RM_Type3Value1);
        oCmd.Parameters.AddWithValue("@RM_Type3Value2", RM_Type3Value2);
        oCmd.Parameters.AddWithValue("@RM_Type3Value3", RM_Type3Value3);
        oCmd.Parameters.AddWithValue("@RM_Type3ValueSum", RM_Type3ValueSum);
        oCmd.Parameters.AddWithValue("@RM_Type4Value1", RM_Type4Value1);
        oCmd.Parameters.AddWithValue("@RM_Type4Value2", RM_Type4Value2);
        oCmd.Parameters.AddWithValue("@RM_Type4Value3", RM_Type4Value3);
        oCmd.Parameters.AddWithValue("@RM_Type4ValueSum", RM_Type4ValueSum);
        oCmd.Parameters.AddWithValue("@RM_PreVal", RM_PreVal);
        oCmd.Parameters.AddWithValue("@RM_ChkVal", RM_ChkVal);
        oCmd.Parameters.AddWithValue("@RM_NotChkVal", RM_NotChkVal);
        oCmd.Parameters.AddWithValue("@RM_Remark", RM_Remark);
        oCmd.Parameters.AddWithValue("@RM_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@RM_ModId", RM_ModId);
        oCmd.Parameters.AddWithValue("@RM_ReportType", RM_ReportType);
        oCmd.Parameters.AddWithValue("@RM_P_ExType", "");
        oCmd.Parameters.AddWithValue("@RM_P_ExDeviceType", "");
        oCmd.Parameters.AddWithValue("@RM_Formula", RM_Formula);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    //檢查有無月報資料 (新增/修改)用
    public DataTable selectMonthReportByPjno()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
            select * from ReportMonth where RM_ProjectGuid=@RM_ProjectGuid and RM_Stage=@RM_Stage and RM_Year=@RM_Year and RM_Month=@RM_Month and RM_ReportType=@RM_ReportType and RM_Status='A'
        ");
        if (RM_CPType!="") {
            sb.Append(@"
                 and RM_CPType=@RM_CPType
            ");
            oCmd.Parameters.AddWithValue("@RM_CPType", RM_CPType);
        }
        //if (RM_ReportType != "")
        //{
        //    sb.Append(@"
        //         and RM_ReportType=@RM_ReportType
        //    ");
        //    oCmd.Parameters.AddWithValue("@RM_ReportType", RM_ReportType);
        //}
        //if (RM_P_ExType != null)
        //{
        //    sb.Append(@"
        //         and RM_P_ExType=@RM_P_ExType
        //    ");
        //    oCmd.Parameters.AddWithValue("@RM_P_ExType", RM_P_ExType);
        //}
        //if (RM_P_ExDeviceType != null)
        //{
        //    sb.Append(@"
        //         and RM_P_ExDeviceType=@RM_P_ExDeviceType
        //    ");
        //    oCmd.Parameters.AddWithValue("@RM_P_ExDeviceType", RM_P_ExDeviceType);
        //}

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();
        oCmd.Parameters.AddWithValue("@RM_ProjectGuid", RM_ProjectGuid);
        oCmd.Parameters.AddWithValue("@RM_Stage", RM_Stage);
        oCmd.Parameters.AddWithValue("@RM_Year", RM_Year);
        oCmd.Parameters.AddWithValue("@RM_Month", RM_Month);
        oCmd.Parameters.AddWithValue("@RM_ReportType", RM_ReportType);

        oda.Fill(dt);
        return dt;
    }

    //將月報資料 update RM_Status='D'
    public void deleteMonthReport()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);

        oCmd.CommandText = @"
            update ReportMonth set RM_Status='D',RM_ModId=@RM_ModId,RM_ModDate=@RM_ModDate where RM_ID=@RM_ID and RM_PGuid=@RM_PGuid
        ";
        
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@RM_ID", RM_ID);
        oCmd.Parameters.AddWithValue("@RM_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@RM_ModId", RM_ModId);
        oCmd.Parameters.AddWithValue("@RM_PGuid", RM_PGuid);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    //撈月報資料 DetailReportMonth (設備汰換)
    public DataTable selectMonthReportDetail()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
            select c.I_People,d.M_Name,d.M_Tel,d.M_City,d.M_Office,e.M_Name as bossname
            ,a.*,b.P_Guid,b.P_ItemName,c.I_1_Sdate,c.I_1_Edate,c.I_2_Sdate,c.I_2_Edate,f.RC_CheckDate,f.RC_CheckType,f.RC_Status
            ,c.I_3_Sdate,c.I_3_Edate,g.C_Item_cn,h.C_Item_cn as cityname
            from ReportMonth a
            left join PushItem b on a.RM_PGuid = b.P_Guid
            left join ProjectInfo c on a.RM_ProjectGuid = c.I_Guid and I_Status<>'D'
            left join Member d on c.I_People = d.M_Guid and d.M_Status<>'D'
            left join Member e on d.M_Manager_ID=e.M_Guid and e.M_Status<>'D'
            left join ReportCheck f on a.RM_ReportGuid=f.RC_ReportGuid and f.RC_ReportType='01' and f.RC_Status<>'D'
            left join CodeTable g on g.C_Group='07' and b.P_ItemName=g.C_Item
            left join CodeTable h on d.M_City=h.C_Item and h.C_Group='02'
            where a.RM_ReportGuid=@RM_ReportGuid and RM_Status<>'D' and RM_ReportType='01'
  
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();
        oCmd.Parameters.AddWithValue("@RM_ReportGuid", RM_ReportGuid);
        oda.Fill(dt);
        return dt;
    }

    //撈月報資料 DetailReportMonth (擴大補助)
    public DataTable selectMonthReportDetailEx()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        #region 舊code
//        sb.Append(@"
//select c.I_People,d.M_Name,d.M_Tel,d.M_City,d.M_Office,e.M_Name as bossname
//,a.*,b.P_Guid,b.P_ItemName,c.I_1_Sdate,c.I_1_Edate,c.I_2_Sdate,c.I_2_Edate,f.RC_CheckDate,f.RC_CheckType,f.RC_Status
//,c.I_3_Sdate,c.I_3_Edate,g.C_Item_cn,h.C_Item_cn as cityname
//,b.P_ExFinish,i.C_Item_cn as P_ItemName_c,b.P_ExType,b.P_ExDeviceType,
//case a.RM_P_ExType when '01' then '行政' when '02' then '設備' else '' end as P_ExType_c,
//case a.RM_P_ExDeviceType when '01' then '空調' when '02' then '照明' when '03' then '非照明' else '' end as P_ExDeviceType_c
//from ReportMonth a
//left join PushItem b on a.RM_PGuid = b.P_Guid
//left join ProjectInfo c on a.RM_ProjectGuid = c.I_Guid and I_Status<>'D'
//left join Member d on c.I_People = d.M_Guid and d.M_Status<>'D'
//left join Member e on d.M_Manager_ID=e.M_Guid and e.M_Status<>'D'
//left join ReportCheck f on a.RM_ReportGuid=f.RC_ReportGuid and f.RC_ReportType='03' and f.RC_Status<>'D'
//left join CodeTable g on g.C_Group='07' and b.P_ItemName=g.C_Item
//left join CodeTable h on d.M_City=h.C_Item and h.C_Group='02'
//left join CodeTable i on b.P_ItemName=i.C_Item and i.C_Group='09'
//where a.RM_ReportGuid=@RM_ReportGuid and RM_Status<>'D' and RM_ReportType='02'
  
//        ");
        #endregion

        sb.Append(@"
declare @I_Guid nvarchar(50)=''
declare @RM_Year nvarchar(50)=''
declare @RM_Month nvarchar(50)=''

select top 1 @I_Guid=RM_ProjectGuid,@RM_Year=RM_Year,@RM_Month=RM_Month from ReportMonth where RM_ReportGuid=@RM_ReportGuid
--select @I_Guid,@RM_Year,@RM_Month

select #tmp.*,
isnull((select sum(isnull(RM_Type4ValueSum,0.0)) from ReportMonth h left join ReportCheck j on h.RM_Stage=j.RC_Stage and j.RC_ReportType='03' where h.RM_Stage=#tmp.P_Period and  h.RM_Year+h.RM_Month<@RM_Year+@RM_Month and h.RM_CPType=#tmp.P_ItemName and j.RC_Status='A' and j.RC_CheckType='Y' and h.RM_Status<>'D' and h.RM_ReportGuid=j.RC_ReportGuid  and h.RM_ProjectGuid=@I_Guid ),0) as countFinishKW,--累計完成數 KW(無風管)
isnull((select sum(isnull(RM_Type3ValueSum,0)) from ReportMonth i left join ReportCheck k on i.RM_Stage=k.RC_Stage and k.RC_ReportType='03' where i.RM_Stage=#tmp.P_Period and  i.RM_Year+i.RM_Month<@RM_Year+@RM_Month and i.RM_CPType=#tmp.P_ItemName and k.RC_Status='A' and k.RC_CheckType='Y' and i.RM_Status<>'D' and i.RM_ReportGuid=k.RC_ReportGuid  and i.RM_ProjectGuid=@I_Guid ),0) as countFinish03,--累計完成數
isnull((select sum(isnull(RM_Type2ValueSum,0)) from ReportMonth m left join ReportCheck n on m.RM_Stage=n.RC_Stage and n.RC_ReportType='03' where m.RM_Stage=#tmp.P_Period and  m.RM_Year+m.RM_Month<@RM_Year+@RM_Month and m.RM_CPType=#tmp.P_ItemName and n.RC_Status='A' and n.RC_CheckType='Y' and m.RM_Status<>'D' and m.RM_ReportGuid=n.RC_ReportGuid  and m.RM_ProjectGuid=@I_Guid ),0) as countFinish02,--累計核定數
isnull((select sum(isnull(RM_Type3ValueSum,0.0)) from ReportMonth o left join ReportCheck p on o.RM_Stage=p.RC_Stage and p.RC_ReportType='03' where o.RM_Stage=#tmp.P_Period and  o.RM_Year+o.RM_Month<@RM_Year+@RM_Month and o.RM_CPType=#tmp.P_ItemName and p.RC_Status='A' and p.RC_CheckType='Y' and o.RM_Status<>'D' and o.RM_ReportGuid=p.RC_ReportGuid and o.RM_ProjectGuid=@I_Guid  ),0) as countApplyKW,--累計申請數 KW(無風管)
isnull((select sum(isnull(RM_Type1ValueSum,0)) from ReportMonth q left join ReportCheck r on q.RM_Stage=r.RC_Stage and r.RC_ReportType='03' where q.RM_Stage=#tmp.P_Period and  q.RM_Year+q.RM_Month<@RM_Year+@RM_Month and q.RM_CPType=#tmp.P_ItemName and r.RC_Status='A' and r.RC_CheckType='Y' and q.RM_Status<>'D' and q.RM_ReportGuid=r.RC_ReportGuid and q.RM_ProjectGuid=@I_Guid  ),0) as countApply01--累計完成數
from 
(
	select c.I_People,d.M_Name,d.M_Tel,d.M_City,d.M_Office,e.M_Name as bossname
	,a.*,b.P_Guid,b.P_ItemName,c.I_1_Sdate,c.I_1_Edate,c.I_2_Sdate,c.I_2_Edate,f.RC_CheckDate,f.RC_CheckType,f.RC_Status
	,c.I_3_Sdate,c.I_3_Edate,g.C_Item_cn,h.C_Item_cn as cityname
	,b.P_ExFinish,i.C_Item_cn as P_ItemName_c,b.P_ExType,b.P_ExDeviceType,b.P_Period,
	case a.RM_P_ExType when '01' then '行政' when '02' then '設備' else '' end as P_ExType_c,
	case a.RM_P_ExDeviceType when '01' then '空調' when '02' then '照明' when '03' then '非照明' else '' end as P_ExDeviceType_c
	from ReportMonth a
	left join PushItem b on a.RM_PGuid = b.P_Guid
	left join ProjectInfo c on a.RM_ProjectGuid = c.I_Guid and I_Status<>'D'
	left join Member d on c.I_People = d.M_Guid and d.M_Status<>'D'
	left join Member e on d.M_Manager_ID=e.M_Guid and e.M_Status<>'D'
	left join ReportCheck f on a.RM_ReportGuid=f.RC_ReportGuid and f.RC_ReportType='03' and f.RC_Status<>'D'
	left join CodeTable g on g.C_Group='07' and b.P_ItemName=g.C_Item
	left join CodeTable h on d.M_City=h.C_Item and h.C_Group='02'
	left join CodeTable i on b.P_ItemName=i.C_Item and i.C_Group='09'
	where a.RM_ReportGuid=@RM_ReportGuid and RM_Status<>'D' and RM_ReportType='02'
)#tmp
  
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();
        oCmd.Parameters.AddWithValue("@RM_ReportGuid", RM_ReportGuid);
        oda.Fill(dt);
        return dt;
    }

    //月報送審前 檢查有沒有儲存
    public DataTable selectMonthReportByGuid()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
            select * from ReportMonth
            where RM_ReportGuid=@RM_ReportGuid and RM_Status<>'D'
  
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();
        oCmd.Parameters.AddWithValue("@RM_ReportGuid", RM_ReportGuid);
        oda.Fill(dt);
        return dt;
    }

    //發審核不通過MAIL 撈出承辦人信箱BY ReportGuid
    public DataSet selectMailByReportguid(string strReportGuid)
    {
        string strmail = "";
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
            declare @strPepoleGuid nvarchar(50)=''
            declare @strManagerGuid nvarchar(50)=''
            select top 1 @strPepoleGuid=RC_PeopleGuid from ReportCheck where RC_ReportGuid=@RC_ReportGuid and RC_Status='A'
            select top 1 @strManagerGuid=M_Manager_ID from Member where M_Guid=@strPepoleGuid and M_Status='A'
            select top 1 M_Email from Member where M_Guid=@strPepoleGuid and M_Status='A'
            select top 1 M_Email from Member where M_Guid=@strManagerGuid and M_Status='A'
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();
        oCmd.Parameters.AddWithValue("@RC_ReportGuid", strReportGuid);
        oda.Fill(ds);
        return ds;
    }

    //20181016 撈月報列表(設備汰換)
    public DataSet getMonthList(string mGuid, string pStart, string pEnd)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
--傳進來的參數
declare @strYear nvarchar(4)=@RM_Year --yyyy 西元年4碼
declare @strMonth nvarchar(1)=@RM_Month --月 01.02.03.04  2碼
declare @strStage nvarchar(1)=@RM_Stage --期數 1.2.3 1碼
declare @M_Guid nvarchar(50)=@mGuid --承辦人M_ID
declare @M_City nvarchar(2) = (select top 1 M_City from Member where M_Guid=@M_Guid and M_Status='A')

create table #tmpAll(
	 RM_ReportGuid nvarchar(50)
	,RM_Stage nvarchar(1)
	,RM_Year nvarchar(4)
	,RM_Month nvarchar(2)
	,RM_CreateDate datetime
	,RC_CheckType nvarchar(2) NULL
	,RC_Status nvarchar(2) NULL
    ,MTypeNum nvarchar(2)
    ,MTypeName nvarchar(10)
)

--先過濾出資料進#tmpAll
insert into #tmpAll(RM_ReportGuid,RM_Stage,RM_Year,RM_Month,RM_CreateDate,RC_CheckType,RC_Status,MTypeNum,MTypeName)
select RM_ReportGuid,RM_Stage,RM_Year,RM_Month,MIN(RM_CreateDate),RC_CheckType,RC_Status,RM_ReportType,strTypeName
from
(
    select RM_ReportGuid,RM_Stage,RM_Year,RM_Month,RM_CreateDate,RC_CheckType,RC_Status,RM_ReportType,case RM_ReportType when '01' then '設備汰換' when '02' then '擴大補助' else '' end as strTypeName
    from ReportMonth
    left join ReportCheck on RM_ReportGuid=RC_ReportGuid and RC_Status<>'D'
    left join ProjectInfo on RM_ProjectGuid=I_Guid and I_Status='A' and I_Flag='Y'
    where I_City=@M_City  and RM_Status='A'
        ");
        if (RM_ReportType!="") {
            sb.Append(@" and RM_ReportType=@RM_ReportType  ");
        }
        if (RM_Year!="") {
            sb.Append(@" and RM_Year=@RM_Year  ");
        }
        if (RM_Month != "")
        {
            sb.Append(@"  and RM_Month=@RM_Month ");
        }
        if (RM_Stage != "")
        {
            sb.Append(@" and RM_Stage=@RM_Stage  ");
        }
        sb.Append(@"
)tmp
group by RM_ReportGuid,RM_Stage,RM_Year,RM_Month,RC_CheckType,RC_Status,RM_ReportType,strTypeName
order by RM_Stage desc,RM_Year desc,RM_Month desc

--總筆數
select count(*) as total from #tmpAll

--分頁資料
select * from (
           --select ROW_NUMBER() over (order by RM_CreateDate) itemNo,#tmpAll.*
           select ROW_NUMBER() over (order by RC_CheckType,RM_Year desc,RM_Month desc,RM_Stage desc,MTypeNum asc) itemNo,#tmpAll.*
		   from #tmpAll
)#tmp where itemNo between @pStart and @pEnd

drop table #tmpAll 
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@mGuid", mGuid);
        oCmd.Parameters.AddWithValue("@RM_Year", RM_Year);
        oCmd.Parameters.AddWithValue("@RM_Month", RM_Month);
        oCmd.Parameters.AddWithValue("@RM_Stage", RM_Stage);
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oCmd.Parameters.AddWithValue("@RM_ReportType", RM_ReportType);
        oda.Fill(ds);
        return ds;
    }

    //新增月報前檢查有沒有新增過 (設備汰換)
    public DataTable chkReportMonth(string mGuid, string year, string month, string stage)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
            declare @PerGuid nvarchar(50)=@mGuid
            declare @ProjectID nvarchar(50)=(select I_Guid from ProjectInfo where I_People=@PerGuid)
            select * from ReportMonth 
            left join ReportCheck on RM_ReportGuid = RC_ReportGuid and RM_ProjectGuid=@ProjectID
            where RM_Stage=@RM_Stage and RM_Year=@RM_Year and RM_Month=@RM_Month and RM_ProjectGuid=@ProjectID
            and ((RC_Status='A' and RC_CheckType='Y') or ((RC_Status='A' and RC_CheckType='Y') or (RC_Status='A' and RC_CheckType is null))) and RM_ReportType='01'
        ");
        

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();
        oCmd.Parameters.AddWithValue("@mGuid", mGuid);
        oCmd.Parameters.AddWithValue("@RM_Stage", stage);
        oCmd.Parameters.AddWithValue("@RM_Year", year);
        oCmd.Parameters.AddWithValue("@RM_Month", month);

        oda.Fill(dt);
        return dt;
    }

    //撈月報資料 (擴大補助)
    public DataSet selectMonthReportEx()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"

select a.M_ID,a.M_Name,a.M_Tel,a.M_Manager_ID as bossid,b.M_Name as bossname
from Member a left join Member b on a.M_Manager_ID=b.M_Guid
where a.M_ID=@M_ID

declare @report_guid nvarchar(50);
declare @checkflag nvarchar(2);
set rowcount 1;
select @report_guid = RM_ReportGuid from ReportMonth where RM_ProjectGuid=@I_Guid and RM_Stage=@P_Period and RM_Year=@RM_Year and RM_Month=@RM_Month and RM_ReportType='02'
set rowcount 0;
select @checkflag = RC_CheckType from ReportCheck where RC_ReportGuid = @report_guid and RC_Status<>'D'

declare @rcount int=0;
declare @ifY nvarchar(4)='',@ifM nvarchar(2)='',@ifYM nvarchar(7)='';

if @checkflag<>'Y' or @checkflag is null
    begin
        select #tmp.* ,
        (select sum(isnull(RM_Type4ValueSum,0.0)) from ReportMonth h left join ReportCheck j on h.RM_Stage=j.RC_Stage and j.RC_ReportType='03' where h.RM_Stage=#tmp.P_Period and  h.RM_Year+h.RM_Month<@RM_Year+@RM_Month and h.RM_CPType=#tmp.P_ItemName and j.RC_Status='A' and j.RC_CheckType='Y' and h.RM_Status<>'D' and h.RM_ReportGuid=j.RC_ReportGuid and h.RM_ProjectGuid=@I_Guid  ) as countFinishKW,--累計完成數 KW(無風管)
        (select sum(isnull(RM_Type3ValueSum,0)) from ReportMonth i left join ReportCheck k on i.RM_Stage=k.RC_Stage and k.RC_ReportType='03' where i.RM_Stage=#tmp.P_Period and  i.RM_Year+i.RM_Month<@RM_Year+@RM_Month and i.RM_CPType=#tmp.P_ItemName and k.RC_Status='A' and k.RC_CheckType='Y' and i.RM_Status<>'D' and i.RM_ReportGuid=k.RC_ReportGuid and i.RM_ProjectGuid=@I_Guid  ) as countFinish03,--累計完成數(老舊 停車場 中型 大型)
        (select sum(isnull(RM_Type2ValueSum,0)) from ReportMonth m left join ReportCheck n on m.RM_Stage=n.RC_Stage and n.RC_ReportType='03' where m.RM_Stage=#tmp.P_Period and  m.RM_Year+m.RM_Month<@RM_Year+@RM_Month and m.RM_CPType=#tmp.P_ItemName and n.RC_Status='A' and n.RC_CheckType='Y' and m.RM_Status<>'D' and m.RM_ReportGuid=n.RC_ReportGuid and m.RM_ProjectGuid=@I_Guid  ) as countFinish02,--累計核定數(老舊 停車場 中型 大型)
        (select sum(isnull(RM_Type3ValueSum,0.0)) from ReportMonth o left join ReportCheck p on o.RM_Stage=p.RC_Stage and p.RC_ReportType='03' where o.RM_Stage=#tmp.P_Period and  o.RM_Year+o.RM_Month<@RM_Year+@RM_Month and o.RM_CPType=#tmp.P_ItemName and p.RC_Status='A' and p.RC_CheckType='Y' and o.RM_Status<>'D' and o.RM_ReportGuid=p.RC_ReportGuid and o.RM_ProjectGuid=@I_Guid  ) as countApplyKW,--累計申請數 KW(無風管)
        (select sum(isnull(RM_Type1ValueSum,0)) from ReportMonth q left join ReportCheck r on q.RM_Stage=r.RC_Stage and r.RC_ReportType='03' where q.RM_Stage=#tmp.P_Period and  q.RM_Year+q.RM_Month<@RM_Year+@RM_Month and q.RM_CPType=#tmp.P_ItemName and r.RC_Status='A' and r.RC_CheckType='Y' and q.RM_Status<>'D' and q.RM_ReportGuid=r.RC_ReportGuid and q.RM_ProjectGuid=@I_Guid  ) as countApply01--累計完成數(老舊 停車場 中型 大型)
        into #tmpAll
        from (
            select a.P_Guid,a.P_ParentId,a.P_Type,a.P_ItemName,a.P_Period,c.C_Item_cn ,a.P_Status,b.*,d.M_Name as M_Name,d.M_Tel,d.M_Manager_ID,e.M_Name as M__Manage_Name,f.RC_CheckType,f.RC_Status,f.RC_CheckDate,
                    a.P_ExFinish,h.C_Item_cn as P_ItemName_c
					--a.P_ExType,a.P_ExDeviceType,a.P_ExType,a.P_ExDeviceType,
                    --case a.P_ExType when '01' then '行政' when '02' then '設備' else '' end as P_ExType_c,
                    --case a.P_ExDeviceType when '01' then '空調' when '02' then '照明' when '03' then '非照明' else '' end as P_ExDeviceType_c
            from PushItem a
            --left join Check_Point cp on a.P_Type='04' and a.P_Guid = cp.CP_ParentId and cp.CP_ProjectId=@I_Guid and cp.CP_Status='A'
			left join ReportMonth b on a.P_ParentId=b.RM_ProjectGuid  and a.P_ItemName=b.RM_CPType and b.RM_Stage=@P_Period and b.RM_Year=@RM_Year and b.RM_Month=@RM_Month and (b.RM_Status<>'D' or b.RM_Status is null) and a.P_Type='04' and a.P_Guid=b.RM_PGuid and a.P_Type='04' and b.RM_ReportType='02'
            left join CodeTable c on c.C_Group='04' and a.P_Type = c.C_Item
            left join Member d on d.M_ID=@M_ID
            left join Member e on d.M_id=@M_ID and d.M_Manager_ID = e.M_Guid
            left join ReportCheck f on b.RM_ReportGuid = f.RC_ReportGuid and RC_Status<>'D'
            left join ProjectInfo g on g.I_Guid = @I_Guid and g.I_Status<>'D'
            left join CodeTable h on a.P_ItemName=h.C_Item and h.C_Group='09'
            where P_ParentId=@I_Guid and P_Period=@P_Period  and a.P_Status<>'D' and a.P_Type='04'  and P_ItemName<>'99' -- and RM_Year=@RM_Year and RM_Month=@RM_Month and (b.RM_ReportType='02' and b.RM_P_ExType<>'01')
        )#tmp
        
        -- and ((convert(int,RM_Year)+convert(int,RM_Month))<(convert(int,@RM_Year)+convert(int,@RM_Month)))
		select @rcount =count(*) from ReportMonth where RM_ReportType='02' and RM_Stage=@P_Period and RM_Formula is not null and RM_ProjectGuid=@I_Guid and RM_Status='A'
		if @rcount>0
			begin
				select @ifY=Min(RM_Year),@ifM=Min(RM_Month) from ReportMonth  where RM_ReportType='02' and RM_Stage=@P_Period and RM_Formula is not null and RM_ProjectGuid=@I_Guid and RM_Status='A'
				select * into #tmpTemp from ReportMonth where RM_ReportType='02' and RM_Stage=@P_Period and RM_Year=@ifY and RM_Month=@ifM and RM_ProjectGuid=@I_Guid and RM_Status='A'
				update #tmpAll set #tmpAll.RM_Formula=b.RM_Formula from #tmpTemp b where #tmpAll.P_Guid=b.RM_PGuid
                drop table #tmpTemp
			end

        --select @ifYM=REPLACE(@ifYM,'/','')
		--if @ifYM=@RM_Year+@RM_Month
		if @rcount=0
			begin
				select 'Y' as strFirst,b.RC_Status,a.* from #tmpAll a left join ReportCheck b on RM_ReportGuid=b.RC_ReportGuid and b.RC_ReportType='03' and b.RC_Status='A'
				order by a.P_ItemName asc
			end
		else
			begin
				select 'N' as strFirst,b.RC_Status,a.* from #tmpAll a left join ReportCheck b on RM_ReportGuid=b.RC_ReportGuid and b.RC_ReportType='03' and b.RC_Status='A'
				order by a.P_ItemName asc
			end

		drop table #tmpAll
		
    end
else
    begin
        select #tmp.* ,
        (select sum(isnull(RM_Type4ValueSum,0.0)) from ReportMonth h left join ReportCheck j on h.RM_Stage=j.RC_Stage and j.RC_ReportType='03' where h.RM_Stage=#tmp.P_Period and  h.RM_Year+h.RM_Month<@RM_Year+@RM_Month and h.RM_CPType=#tmp.P_ItemName and j.RC_Status='A' and j.RC_CheckType='Y' and h.RM_Status<>'D' and h.RM_ReportGuid=j.RC_ReportGuid  and h.RM_ProjectGuid=@I_Guid ) as countFinishKW,--累計完成數 KW(無風管)
        (select sum(isnull(RM_Type3ValueSum,0)) from ReportMonth i left join ReportCheck k on i.RM_Stage=k.RC_Stage and k.RC_ReportType='03' where i.RM_Stage=#tmp.P_Period and  i.RM_Year+i.RM_Month<@RM_Year+@RM_Month and i.RM_CPType=#tmp.P_ItemName and k.RC_Status='A' and k.RC_CheckType='Y' and i.RM_Status<>'D' and i.RM_ReportGuid=k.RC_ReportGuid  and i.RM_ProjectGuid=@I_Guid ) as countFinish03,--累計完成數(老舊 停車場 中型 大型)
        (select sum(isnull(RM_Type2ValueSum,0)) from ReportMonth m left join ReportCheck n on m.RM_Stage=n.RC_Stage and n.RC_ReportType='03' where m.RM_Stage=#tmp.P_Period and  m.RM_Year+m.RM_Month<@RM_Year+@RM_Month and m.RM_CPType=#tmp.P_ItemName and n.RC_Status='A' and n.RC_CheckType='Y' and m.RM_Status<>'D' and m.RM_ReportGuid=n.RC_ReportGuid  and m.RM_ProjectGuid=@I_Guid ) as countFinish02,--累計核定數(老舊 停車場 中型 大型)
        (select sum(isnull(RM_Type3ValueSum,0.0)) from ReportMonth o left join ReportCheck p on o.RM_Stage=p.RC_Stage and p.RC_ReportType='03' where o.RM_Stage=#tmp.P_Period and  o.RM_Year+o.RM_Month<@RM_Year+@RM_Month and o.RM_CPType=#tmp.P_ItemName and p.RC_Status='A' and p.RC_CheckType='Y' and o.RM_Status<>'D' and o.RM_ReportGuid=p.RC_ReportGuid and o.RM_ProjectGuid=@I_Guid  ) as countApplyKW,--累計申請數 KW(無風管)
        (select sum(isnull(RM_Type1ValueSum,0)) from ReportMonth q left join ReportCheck r on q.RM_Stage=r.RC_Stage and r.RC_ReportType='03' where q.RM_Stage=#tmp.P_Period and  q.RM_Year+q.RM_Month<@RM_Year+@RM_Month and q.RM_CPType=#tmp.P_ItemName and r.RC_Status='A' and r.RC_CheckType='Y' and q.RM_Status<>'D' and q.RM_ReportGuid=r.RC_ReportGuid and q.RM_ProjectGuid=@I_Guid  ) as countApply01--累計完成數(老舊 停車場 中型 大型)
        into #tmpAll2
        from (
            select a.P_Guid,a.P_ParentId,a.P_Type,a.P_ItemName,a.P_Period,c.C_Item_cn ,a.P_Status,b.*,d.M_Name as M_Name,d.M_Tel,d.M_Manager_ID,e.M_Name as M_Manage_Name,f.RC_CheckType,f.RC_Status,f.RC_CheckDate,
                    a.P_ExFinish,h.C_Item_cn as P_ItemName_c
					--a.P_ExType,a.P_ExDeviceType,
                    --case a.P_ExType when '01' then '行政' when '02' then '設備' else '' end as P_ExType_c,
                    --case a.P_ExDeviceType when '01' then '空調' when '02' then '照明' when '03' then '非照明' else '' end as P_ExDeviceType_c
            from PushItem a
			--left join Check_Point cp on a.P_Type='04' and a.P_Guid = cp.CP_ParentId and cp.CP_ProjectId=@I_Guid and cp.CP_Status='A'
            left join ReportMonth b on a.P_ParentId=b.RM_ProjectGuid and a.P_ItemName=b.RM_CPType and a.P_Guid=b.RM_PGuid and b.RM_Stage=@P_Period and b.RM_Year=@RM_Year and b.RM_Month=@RM_Month and (b.RM_Status<>'D' or b.RM_Status is null) and a.P_Type='04'  and a.P_Guid=b.RM_PGuid and b.RM_ReportType='02'
            left join CodeTable c on c.C_Group='04' and a.P_Type = c.C_Item
            left join Member d on d.M_ID=@M_ID
            left join Member e on d.M_id=@M_ID and d.M_Manager_ID = e.M_Guid
            left join ReportCheck f on b.RM_ReportGuid = f.RC_ReportGuid and RC_Status<>'D'
            left join ProjectInfo g on g.I_Guid = @I_Guid and g.I_Status<>'D'
            left join CodeTable h on a.P_ItemName=h.C_Item and h.C_Group='09'
            where P_ParentId=@I_Guid and P_Period=@P_Period and b.RM_Status<>'D' and P_Type='04'  and P_ItemName<>'99'
        )#tmp
        
		select @rcount =count(*) from ReportMonth where RM_ReportType='02' and RM_Stage=@P_Period and RM_Formula is not null and RM_ProjectGuid=@I_Guid and RM_Status='A'
		if @rcount>0
			begin
				select @ifY=Min(RM_Year),@ifM=Min(RM_Month) from ReportMonth  where RM_ReportType='02' and RM_Stage=@P_Period and RM_Formula is not null and RM_ProjectGuid=@I_Guid and RM_Status='A'
				select * into #tmpTemp2 from ReportMonth where RM_ReportType='02' and RM_Stage=@P_Period and RM_Year=@ifY and RM_Month=@ifM and RM_ProjectGuid=@I_Guid and RM_Status='A'
				update #tmpAll2 set #tmpAll2.RM_Formula=b.RM_Formula from #tmpTemp2 b where #tmpAll2.P_Guid=b.RM_PGuid
                drop table #tmpTemp2
			end

        --select @ifYM=REPLACE(@ifYM,'/','')
		--if @ifYM=@RM_Year+@RM_Month
		if @rcount=0
			begin
				select 'Y' as strFirst,b.RC_Status,a.* from #tmpAll2 a left join ReportCheck b on RM_ReportGuid=b.RC_ReportGuid and b.RC_ReportType='03' and b.RC_Status='A'
				order by a.P_ItemName asc
			end
		else
			begin
				select 'N' as strFirst,b.RC_Status,a.* from #tmpAll2 a left join ReportCheck b on RM_ReportGuid=b.RC_ReportGuid and b.RC_ReportType='03' and b.RC_Status='A'
				order by a.P_ItemName asc
			end

		drop table #tmpAll2
		
    end
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();
        oCmd.Parameters.AddWithValue("@M_ID", M_ID);
        oCmd.Parameters.AddWithValue("@I_Guid", I_Guid);
        oCmd.Parameters.AddWithValue("@P_Period", P_Period);
        oCmd.Parameters.AddWithValue("@RM_Year", RM_Year);
        oCmd.Parameters.AddWithValue("@RM_Month", RM_Month);
        oda.Fill(ds);
        return ds;
    }
}
