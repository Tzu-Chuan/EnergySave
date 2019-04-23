using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;

/// <summary>
/// MoneyExecute_DB 的摘要描述
/// </summary>
public class MoneyExecute_DB
{
    string KeyWord = string.Empty;
    string str_city = string.Empty;
    string str_stage = string.Empty;
    public string _KeyWord
    {
        set { KeyWord = value; }
    }
    public string _str_city
    {
        set { str_city = value; }
    }
    public string _str_stage
    {
        set { str_stage = value; }
    }

    #region 私用
    string PR_ID = string.Empty;
    string PR_Stage = string.Empty;
    string PR_City = string.Empty;
    string PR_PlanTitle = string.Empty;
    string PR_Office = string.Empty;
    string PR_Money = string.Empty;
    string PR_Steps = string.Empty;
    string PR_CreateId = string.Empty;
    string PR_ModId = string.Empty;
    string PR_Status = string.Empty;
    string PR_CaseName = string.Empty;
    string PR_CaseMoney = string.Empty;
    string PR_SelfMoney = string.Empty;
    DateTime PR_CreateDate;
    DateTime PR_ModDate;
    #endregion
    #region 公用
    public string _PR_ID
    {
        set { PR_ID = value; }
    }
    public string _PR_Stage
    {
        set { PR_Stage = value; }
    }
    public string _PR_City
    {
        set { PR_City = value; }
    }
    public string _PR_PlanTitle
    {
        set { PR_PlanTitle = value; }
    }
    public string _PR_Office
    {
        set { PR_Office = value; }
    }
    public string _PR_Money
    {
        set { PR_Money = value; }
    }
    public string _PR_Steps
    {
        set { PR_Steps = value; }
    }
    public string _PR_CreateId
    {
        set { PR_CreateId = value; }
    }
    public string _PR_ModId
    {
        set { PR_ModId = value; }
    }
    public string _PR_Status
    {
        set { PR_Status = value; }
    }
    public string _PR_CaseName
    {
        set { PR_CaseName = value; }
    }
    public string _PR_CaseMoney
    {
        set { PR_CaseMoney = value; }
    }
    public string _PR_SelfMoney
    {
        set { PR_SelfMoney = value; }
    }
    public DateTime _PR_CreateDate
    {
        set { PR_CreateDate = value; }
    }
    public DateTime _PR_ModDate
    {
        set { PR_ModDate = value; }
    }
    #endregion

    public DataSet getMoneyList()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select a.* ,b.C_Item_cn as planNamer,c.C_Item_cn as CityName
,REPLACE(CONVERT(varchar(100),CAST(a.PR_Money as money),1),'.00','') as PR_Money_m
,REPLACE(CONVERT(varchar(100),CAST(a.PR_CaseMoney as money),1),'.00','') as PR_CaseMoney_m
,REPLACE(CONVERT(varchar(100),CAST(a.PR_SelfMoney as money),1),'.00','') as PR_SelfMoney_m
from PayRun a
left join CodeTable b on a.PR_PlanTitle = b.C_Item and b.C_Group='08'
left join CodeTable c on a.PR_City = c.C_Item and c.C_Group='02'
where PR_Status='A'
");
        if (PR_City != "")
            sb.Append(@" and PR_City=@PR_City ");
        if (PR_Stage != "")
            sb.Append(@" and PR_Stage=@PR_Stage ");

        sb.Append(@"
--執行經費
if @PR_City<>''
    begin
        if @PR_Stage='1'
            begin
                select I_Money_item1_all as MoneyAll from ProjectInfo where I_City=@PR_City and I_Status='A' and I_Flag='Y'
            end
        if @PR_Stage='2'
            begin
                select I_Money_item2_all as MoneyAll from ProjectInfo where I_City=@PR_City and I_Status='A' and I_Flag='Y'
            end
        if @PR_Stage='3'
            begin
                select I_Money_item3_all as MoneyAll from ProjectInfo where I_City=@PR_City and I_Status='A' and I_Flag='Y'
            end
    end
else
    begin
        select '' as MoneyAll
    end
        ");
        
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@PR_City", PR_City);
        oCmd.Parameters.AddWithValue("@PR_Stage", PR_Stage);
        oda.Fill(ds);
        return ds;
    }
    
    //前台撈該機關當期的執行狀況列表&當期經費
    public DataSet getMoney()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@" 

--當期執行經費列表
SELECT a.*
,REPLACE(CONVERT(varchar(100),CAST(a.PR_Money as money),1),'.00','') as PR_Money_m
,REPLACE(CONVERT(varchar(100),CAST(a.PR_CaseMoney as money),1),'.00','') as PR_CaseMoney_m
,REPLACE(CONVERT(varchar(100),CAST(a.PR_SelfMoney as money),1),'.00','') as PR_SelfMoney_m
,b.C_Item_cn 
from PayRun a left join CodeTable b on b.C_Group='08' and a.PR_PlanTitle=b.C_Item 
 where a.PR_Stage=@stage and a.PR_City=@M_City and a.PR_Status='A'  
order by a.PR_ID

--執行經費
if @stage='1'
    begin
        select I_Money_item1_all as MoneyAll from ProjectInfo where I_City=@M_City and I_Status='A' and I_Flag='Y'
    end
if @stage='2'
    begin
        select I_Money_item2_all as MoneyAll from ProjectInfo where I_City=@M_City and I_Status='A' and I_Flag='Y'
    end
if @stage='3'
    begin
        select I_Money_item3_all as MoneyAll from ProjectInfo where I_City=@M_City and I_Status='A' and I_Flag='Y'
    end

--執行機關中文名稱
select C_Item_cn from CodeTable where C_Item=@M_City

");
        oCmd.Parameters.AddWithValue("@M_City", str_city);
        oCmd.Parameters.AddWithValue("@stage", str_stage);

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();
        oda.Fill(ds);
        return ds;
    }

    //前台經費執行 新增
    public void addMoney()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        
        oCmd.CommandText = @"
            Insert into PayRun (PR_Stage,PR_City,PR_PlanTitle,PR_Office,PR_Money,PR_CaseName,PR_CaseMoney,PR_SelfMoney,PR_Steps,PR_CreateId,PR_Status)
            values(@PR_Stage,@PR_City,@PR_PlanTitle,@PR_Office,@PR_Money,@PR_CaseName,@PR_CaseMoney,@PR_SelfMoney,@PR_Steps,@PR_CreateId,'A')
        ";
        
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@PR_Stage", PR_Stage);
        oCmd.Parameters.AddWithValue("@PR_City", PR_City);
        oCmd.Parameters.AddWithValue("@PR_PlanTitle", PR_PlanTitle);
        oCmd.Parameters.AddWithValue("@PR_Office", PR_Office);
        oCmd.Parameters.AddWithValue("@PR_Money", PR_Money);
        oCmd.Parameters.AddWithValue("@PR_CaseName", PR_CaseName);
        oCmd.Parameters.AddWithValue("@PR_CaseMoney", PR_CaseMoney);
        oCmd.Parameters.AddWithValue("@PR_SelfMoney", PR_SelfMoney);
        oCmd.Parameters.AddWithValue("@PR_Steps", PR_Steps);
        oCmd.Parameters.AddWithValue("@PR_CreateId", PR_CreateId);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    //前台經費執行 修改
    public void modMoney()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        
        oCmd.CommandText = @"
            update PayRun set PR_Stage=@PR_Stage,PR_City=@PR_City,PR_PlanTitle=@PR_PlanTitle,PR_Office=@PR_Office,PR_Money=@PR_Money
                            ,PR_CaseName=@PR_CaseName,PR_CaseMoney=@PR_CaseMoney,PR_SelfMoney=@PR_SelfMoney,PR_Steps=@PR_Steps,PR_ModId=@PR_ModId,PR_ModDate=@PR_ModDate
            where PR_ID=@PR_ID
        ";
        
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@PR_ID", PR_ID);
        oCmd.Parameters.AddWithValue("@PR_Stage", PR_Stage);
        oCmd.Parameters.AddWithValue("@PR_City", PR_City);
        oCmd.Parameters.AddWithValue("@PR_PlanTitle", PR_PlanTitle);
        oCmd.Parameters.AddWithValue("@PR_Office", PR_Office);
        oCmd.Parameters.AddWithValue("@PR_Money", PR_Money);
        oCmd.Parameters.AddWithValue("@PR_CaseName", PR_CaseName);
        oCmd.Parameters.AddWithValue("@PR_CaseMoney", PR_CaseMoney);
        oCmd.Parameters.AddWithValue("@PR_SelfMoney", PR_SelfMoney);
        oCmd.Parameters.AddWithValue("@PR_Steps", PR_Steps);
        oCmd.Parameters.AddWithValue("@PR_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@PR_ModId", PR_ModId);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    //刪除經費資料
    public void delMoney()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"
        update PayRun set PR_ModId = @PR_ModId,PR_ModDate = @PR_ModDate,PR_Status = 'D'
        where PR_ID = @PR_ID
        ";
        oCmd.CommandType = new CommandType();
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("PR_ID", PR_ID);
        oCmd.Parameters.AddWithValue("@PR_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@PR_ModId", PR_ModId);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    //撈經費資料 BY PR_ID
    public DataTable getMoneyById()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select a.*,b.C_Item_cn as PlanName,c.C_Item_cn as CityName
from PayRun a 
left join CodeTable b on b.C_Group='08' and a.PR_PlanTitle=b.C_Item
left join CodeTable c on c.C_Group='02' and a.PR_City=c.C_Item
where PR_ID = @PR_ID 
            ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();
        oCmd.Parameters.AddWithValue("@PR_ID", PR_ID);
        oda.Fill(dt);
        return dt;
    }
}