using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;

/// <summary>
/// Progress_DB 的摘要描述
/// </summary>
public class Progress_DB
{
    #region 私用
    string P_ID = string.Empty;
    string P_Guid = string.Empty;
    string P_ParentId = string.Empty;
    string P_ProjectId = string.Empty;
    string P_Value = string.Empty;
    string P_CreateId = string.Empty;
    string P_ModId = string.Empty;
    string P_Status = string.Empty;

    DateTime P_CreateDate;
    DateTime P_ModDate;
    #endregion
    #region 公用
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
    public string _P_ProjectId
    {
        set { P_ProjectId = value; }
    }
    public string _P_Value
    {
        set { P_Value = value; }
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
    public DateTime _P_CreateDate
    {
        set { P_CreateDate = value; }
    }
    public DateTime _P_ModDate
    {
        set { P_ModDate = value; }
    }
    #endregion

    public DataTable SelectList(string ProjectId, string pType)
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
where CP_ProjectId=@ProjectId and CP_Status='A' 
group by P_Type
)#t

select pui.P_Type,pui.P_ItemName,pui.P_WorkRatio,CP_Guid,CP_ParentId,CP_Point,CP_ReserveYear,RIGHT(REPLICATE('0', 1) + CP_ReserveMonth, 2) CP_ReserveMonth
,CP_Desc,p.P_Value,TypeTotal 
from PushItem pui
left join Check_Point on CP_ParentId=P_Guid
left join Process p on p.P_ParentId=CP_Guid
left join #tmp on pui.P_Type=tp

where CP_ProjectId=@ProjectId and CP_Status='A' ");

        if (pType != "")
            sb.Append(@"and P_Type=@pType ");

        sb.Append(@"order by pui.P_Sort,pui.P_ID,pui.P_Type,CP_ReserveYear,CP_ReserveMonth 
drop table #tmp
");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oCmd.Parameters.AddWithValue("@ProjectId", ProjectId);
        oCmd.Parameters.AddWithValue("@pType", pType);
        oda.Fill(ds);
        return ds;
    }

    public void saveData()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"insert into Process (
M_Guid,
M_Account,
M_Pwd,
M_Name,
M_JobTitle,
M_Tel,
M_Ext,
M_Fax,
M_Phone,
M_Email,
M_Addr,
M_City,
M_Office,
M_Competence,
M_Manager_ID,
M_CreateId,
M_ModId,
M_Status
) values (
@M_Guid,
@M_Account,
@M_Pwd,
@M_Name,
@M_JobTitle,
@M_Tel,
@M_Ext,
@M_Fax,
@M_Phone,
@M_Email,
@M_Addr,
@M_City,
@M_Office,
@M_Competence,
@M_Manager_ID,
@M_CreateId,
@M_ModId,
@M_Status
) ";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@P_ID", P_ID);
        oCmd.Parameters.AddWithValue("@P_Guid", P_Guid);
        oCmd.Parameters.AddWithValue("@P_ParentId", P_ParentId);
        oCmd.Parameters.AddWithValue("@P_ProjectId", P_ProjectId);
        oCmd.Parameters.AddWithValue("@P_Value", P_Value);
        oCmd.Parameters.AddWithValue("@P_CreateId", P_CreateId);
        oCmd.Parameters.AddWithValue("@P_ModId", P_ModId);
        oCmd.Parameters.AddWithValue("@P_ModDate", P_ModDate);
        oCmd.Parameters.AddWithValue("@P_Status", "A");

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }
}