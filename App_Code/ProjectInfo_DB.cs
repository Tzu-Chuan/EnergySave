using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;
/// <summary>
/// ProjectInfo_DB 的摘要描述
/// </summary>
public class ProjectInfo_DB
{
    public ProjectInfo_DB()
    {
        //
        // TODO: 在這裡新增建構函式邏輯
        //
    }
    #region 全私用
    string M_Competence = string.Empty;
    string saveType = string.Empty;
    string str_keyword = string.Empty;
    string M_ID = string.Empty;
    #endregion
    #region 全公用
    public string _M_Competence
    {
        set { M_Competence = value; }
    }
    public string _saveType
    {
        set { saveType = value; }
    }
    public string _str_keyword
    {
        set { str_keyword = value; }
    }
    public string _M_ID
    {
        set { M_ID = value; }
    }
    #endregion

    #region 私用
    string I_ID = string.Empty;
    string I_GUID = string.Empty;
    string I_City = string.Empty;
    string I_Office = string.Empty;
    string I_People = string.Empty;
    string I_1_Sdate = string.Empty;
    string I_1_Edate = string.Empty;
    string I_2_Sdate = string.Empty;
    string I_2_Edate = string.Empty;
    string I_3_Sdate = string.Empty;
    string I_3_Edate = string.Empty;
    string I_Money_item1_1 = string.Empty;
    string I_Money_item1_2 = string.Empty;
    string I_Money_item1_3 = string.Empty;
    string I_Money_item1_all = string.Empty;
    string I_Money_item2_1 = string.Empty;
    string I_Money_item2_2 = string.Empty;
    string I_Money_item2_3 = string.Empty;
    string I_Money_item2_all = string.Empty;
    string I_Money_item3_1 = string.Empty;
    string I_Money_item3_2 = string.Empty;
    string I_Money_item3_3 = string.Empty;
    string I_Money_item3_all = string.Empty;
    string I_Money_item4_1 = string.Empty;
    string I_Money_item4_2 = string.Empty;
    string I_Money_item4_3 = string.Empty;
    string I_Money_item4_all = string.Empty;
    string I_Other_Oneself = string.Empty;
    Decimal I_Other_Oneself_Money;
    string I_Other_Other = string.Empty;
    string I_Other_Other_name= string.Empty;
    Decimal I_Other_Other_Money;
    string I_Target = string.Empty;
    string I_Summary = string.Empty;
    Decimal I_Finish_item1_1;
    Decimal I_Finish_item1_2;
    Decimal I_Finish_item1_3;
    Decimal I_Finish_item1_all;
    Decimal I_Finish_item2_1;
    Decimal I_Finish_item2_2;
    Decimal I_Finish_item2_3;
    Decimal I_Finish_item2_all;
    Decimal I_Finish_item3_1;
    Decimal I_Finish_item3_2;
    Decimal I_Finish_item3_3;
    Decimal I_Finish_item3_all;
    Decimal I_Finish_item4_1;
    Decimal I_Finish_item4_2;
    Decimal I_Finish_item4_3;
    Decimal I_Finish_item4_all;
    Decimal I_Finish_item5_1;
    Decimal I_Finish_item5_2;
    Decimal I_Finish_item5_3;
    Decimal I_Finish_item5_all;
    string I_Boss = string.Empty;
    DateTime I_Createdate;
    DateTime I_Modifydate;
    string I_Flag = string.Empty;
    string I_Status = string.Empty;
    DateTime IC_Checkdate;
    string I_Examine = string.Empty;
    string I_CreateId = string.Empty;
    string I_ModId = string.Empty;
    #endregion
    #region 公用
    public string _I_ID
    {
        set { I_ID = value; }
    }
    public string _I_GUID
    {
        set { I_GUID = value; }
    }
    public string _I_City
    {
        set { I_City = value; }
    }
    public string _I_Office
    {
        set { I_Office = value; }
    }
    public string _I_People
    {
        set { I_People = value; }
    }
    public string _I_1_Sdate
    {
        set { I_1_Sdate = value; }
    }
    public string _I_1_Edate
    {
        set { I_1_Edate = value; }
    }
    public string _I_2_Sdate
    {
        set { I_2_Sdate = value; }
    }
    public string _I_2_Edate
    {
        set { I_2_Edate = value; }
    }
    public string _I_3_Sdate
    {
        set { I_3_Sdate = value; }
    }
    public string _I_3_Edate
    {
        set { I_3_Edate = value; }
    }
    public string _I_Money_item1_1
    {
        set { I_Money_item1_1 = value; }
    }
    public string _I_Money_item1_2
    {
        set { I_Money_item1_2 = value; }
    }
    public string _I_Money_item1_3
    {
        set { I_Money_item1_3 = value; }
    }
    public string _I_Money_item1_all
    {
        set { I_Money_item1_all = value; }
    }
    public string _I_Money_item2_1
    {
        set { I_Money_item2_1 = value; }
    }
    public string _I_Money_item2_2
    {
        set { I_Money_item2_2 = value; }
    }
    public string _I_Money_item2_3
    {
        set { I_Money_item2_3 = value; }
    }
    public string _I_Money_item2_all
    {
        set { I_Money_item2_all = value; }
    }

    public string _I_Money_item4_1
    {
        set { I_Money_item4_1 = value; }
    }
    public string _I_Money_item4_2
    {
        set { I_Money_item4_2 = value; }
    }
    public string _I_Money_item4_3
    {
        set { I_Money_item4_3 = value; }
    }
    public string _I_Money_item4_all
    {
        set { I_Money_item4_all = value; }
    }


    public string _I_Money_item3_1
    {
        set { I_Money_item3_1 = value; }
    }
    public string _I_Money_item3_2
    {
        set { I_Money_item3_2 = value; }
    }
    public string _I_Money_item3_3
    {
        set { I_Money_item3_3 = value; }
    }
    public string _I_Money_item3_all
    {
        set { I_Money_item3_all = value; }
    }
    public string _I_Other_Oneself
    {
        set { I_Other_Oneself = value; }
    }
    public Decimal _I_Other_Oneself_Money
    {
        set { I_Other_Oneself_Money = value; }
    }
    public string _I_Other_Other
    {
        set { I_Other_Other = value; }
    }
    public string _I_Other_Other_name
    {
        set { I_Other_Other_name = value; }
    }
    public Decimal _I_Other_Other_Money
    {
        set { I_Other_Other_Money = value; }
    }
    public string _I_Target
    {
        set { I_Target = value; }
    }
    public string _I_Summary
    {
        set { I_Summary = value; }
    }
    public Decimal _I_Finish_item1_1
    {
        set { I_Finish_item1_1 = value; }
    }
    public Decimal _I_Finish_item1_2
    {
        set { I_Finish_item1_2 = value; }
    }
    public Decimal _I_Finish_item1_3
    {
        set { I_Finish_item1_3 = value; }
    }
    public Decimal _I_Finish_item1_all
    {
        set { I_Finish_item1_all = value; }
    }
    public Decimal _I_Finish_item2_1
    {
        set { I_Finish_item2_1 = value; }
    }
    public Decimal _I_Finish_item2_2
    {
        set { I_Finish_item2_2 = value; }
    }
    public Decimal _I_Finish_item2_3
    {
        set { I_Finish_item2_3 = value; }
    }
    public Decimal _I_Finish_item2_all
    {
        set { I_Finish_item2_all = value; }
    }
    public Decimal _I_Finish_item3_1
    {
        set { I_Finish_item3_1 = value; }
    }
    public Decimal _I_Finish_item3_2
    {
        set { I_Finish_item3_2 = value; }
    }
    public Decimal _I_Finish_item3_3
    {
        set { I_Finish_item3_3 = value; }
    }
    public Decimal _I_Finish_item3_all
    {
        set { I_Finish_item3_all = value; }
    }
    public Decimal _I_Finish_item4_1
    {
        set { I_Finish_item4_1 = value; }
    }
    public Decimal _I_Finish_item4_2
    {
        set { I_Finish_item4_2 = value; }
    }
    public Decimal _I_Finish_item4_3
    {
        set { I_Finish_item4_3 = value; }
    }
    public Decimal _I_Finish_item4_all
    {
        set { I_Finish_item4_all = value; }
    }
    public Decimal _I_Finish_item5_1
    {
        set { I_Finish_item5_1 = value; }
    }
    public Decimal _I_Finish_item5_2
    {
        set { I_Finish_item5_2 = value; }
    }
    public Decimal _I_Finish_item5_3
    {
        set { I_Finish_item5_3 = value; }
    }
    public Decimal _I_Finish_item5_all
    {
        set { I_Finish_item5_all = value; }
    }
    public string _I_Boss
    {
        set { I_Boss = value; }
    }
    public DateTime _I_Createdate
    {
        set { I_Createdate = value; }
    }
    public DateTime _I_Modifydate
    {
        set { I_Modifydate = value; }
    }
    public string _I_Flag
    {
        set { I_Flag = value; }
    }
    public string _I_Status
    {
        set { I_Status = value; }
    }
    public string _I_Examine
    {
        set { I_Examine = value; }
    }
    public DateTime _IC_Checkdate
    {
        set { IC_Checkdate = value; }
    }
    public string _I_CreateId
    {
        set { I_CreateId = value; }
    }
    public string _I_ModId
    {
        set { I_ModId = value; }
    }
    #endregion

    //撈計畫資料 ProjectInfo 
    //1. 每個人近來填寫撈屬於自己的計畫資料 BY I_People(填寫人(承辦人))
    //2. 參考別人的計畫 BY I_ID
    public DataTable getProjectInfo()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
            select a.* 
            ,b.M_Name as Cname,b.M_JobTitle as JobTitle,b.M_Tel as Tel,b.M_Phone as Phone,b.M_Fax as Fax,b.M_Email as Email,b.M_Addr as Addr,c.C_Item_cn as CityName,
            d.M_NAme as ManagerCname,d.M_JobTitle as ManagerJobTitle,d.M_Tel as ManagerTel,d.M_Phone as ManagerPhone,d.M_Fax as ManagerFax,e.C_Item_cn as ManagerCityName,
            d.M_Email as ManagerEmail,d.M_Addr as ManagerAddr,
            (select I_Guid from  ProjectInfo f where f.I_City= a.I_City and f.I_Flag='Y')  as chk_flag,b.M_ID,b.M_Name,b.M_Office
            from ProjectInfo a
            left join Member b on  a.I_People = b.M_Guid--承辦人資料
            left join CodeTable c on c.C_Group='02' and a.I_City=c.C_Item --承辦人機關名稱
            left join Member d on  a.I_Boss = d.M_Guid--承辦主管資料
            left join CodeTable e on e.C_Group='02' and a.I_City=e.C_Item and a.I_Boss = d.M_Guid --承辦主管機關名稱
            where 1=1 
        ");
        if (I_ID!="") {//管理這看計畫資料會帶計畫ID
            sb.Append(@" and I_ID=@I_ID ");
            oCmd.Parameters.AddWithValue("@I_ID", I_ID);
        }
        else
        {//承辦人看資料帶計畫GUID
            sb.Append(@" and I_Guid=@I_Guid ");
            oCmd.Parameters.AddWithValue("@I_Guid", I_GUID);
        }
        

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        
        
        oda.Fill(ds);
        return ds;
    }

    //修改計畫基本資料
    public void modProjectInfo()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        if (M_ID != "")
        {//管理者進來修改
            //承辦人修改
            oCmd.CommandText = @"
                declare @I_City nvarchar(5);
                declare @I_Boss nvarchar(50);
                declare @I_Office nvarchar(50);
                declare @I_People nvarchar(50);
                select @I_City=M_City,@I_Boss=M_Manager_ID,@I_Office=M_Office,@I_People=M_Guid from Member where M_ID=@M_ID and M_Status<>'D'
                update ProjectInfo set
                I_1_Sdate=@I_1_Sdate,I_1_Edate=@I_1_Edate,I_2_Sdate=@I_2_Sdate,I_2_Edate=@I_2_Edate,I_3_Sdate=@I_3_Sdate,I_3_Edate=@I_3_Edate,
                I_Money_item1_1=@I_Money_item1_1,I_Money_item1_2=@I_Money_item1_2,I_Money_item1_3=@I_Money_item1_3,I_Money_item1_all=@I_Money_item1_all,
                I_Money_item2_1=@I_Money_item2_1,I_Money_item2_2=@I_Money_item2_2,I_Money_item2_3=@I_Money_item2_3,I_Money_item2_all=@I_Money_item2_all,
                I_Money_item3_1=@I_Money_item3_1,I_Money_item3_2=@I_Money_item3_2,I_Money_item3_3=@I_Money_item3_3,I_Money_item3_all=@I_Money_item3_all,
                I_Money_item4_1=@I_Money_item4_1,I_Money_item4_2=@I_Money_item4_2,I_Money_item4_3=@I_Money_item4_3,I_Money_item4_all=@I_Money_item4_all,
                I_Other_Oneself=@I_Other_Oneself,I_Other_Oneself_Money=@I_Other_Oneself_Money,I_Other_Other=@I_Other_Other,I_Other_Other_name=@I_Other_Other_name,I_Other_Other_Money=@I_Other_Other_Money,
                I_Target=@I_Target,I_Summary=@I_Summary,
                I_Finish_item1_1=@I_Finish_item1_1,I_Finish_item1_2=@I_Finish_item1_2,I_Finish_item1_3=@I_Finish_item1_3,I_Finish_item1_all=@I_Finish_item1_all,
                I_Finish_item2_1=@I_Finish_item2_1,I_Finish_item2_2=@I_Finish_item2_2,I_Finish_item2_3=@I_Finish_item2_3,I_Finish_item2_all=@I_Finish_item2_all,
                I_Finish_item3_1=@I_Finish_item3_1,I_Finish_item3_2=@I_Finish_item3_2,I_Finish_item3_3=@I_Finish_item3_3,I_Finish_item3_all=@I_Finish_item3_all,
                I_Finish_item4_1=@I_Finish_item4_1,I_Finish_item4_2=@I_Finish_item4_2,I_Finish_item4_3=@I_Finish_item4_3,I_Finish_item4_all=@I_Finish_item4_all,
                I_Finish_item5_1=@I_Finish_item5_1,I_Finish_item5_2=@I_Finish_item5_2,I_Finish_item5_3=@I_Finish_item5_3,I_Finish_item5_all=@I_Finish_item5_all,
                I_Modifydate=@I_Modifydate,I_ModId=@I_ModId,I_City=@I_City,I_Boss=@I_Boss,I_Office=@I_Office,I_People=@I_People
                
                where I_ID=@I_ID and I_Guid=@I_GUID
            ";
            oCmd.Parameters.AddWithValue("@M_ID", M_ID);
        }
        else {
            //承辦人修改
            oCmd.CommandText = @"
                update ProjectInfo set
                I_1_Sdate=@I_1_Sdate,I_1_Edate=@I_1_Edate,I_2_Sdate=@I_2_Sdate,I_2_Edate=@I_2_Edate,I_3_Sdate=@I_3_Sdate,I_3_Edate=@I_3_Edate,
                I_Money_item1_1=@I_Money_item1_1,I_Money_item1_2=@I_Money_item1_2,I_Money_item1_3=@I_Money_item1_3,I_Money_item1_all=@I_Money_item1_all,
                I_Money_item2_1=@I_Money_item2_1,I_Money_item2_2=@I_Money_item2_2,I_Money_item2_3=@I_Money_item2_3,I_Money_item2_all=@I_Money_item2_all,
                I_Money_item3_1=@I_Money_item3_1,I_Money_item3_2=@I_Money_item3_2,I_Money_item3_3=@I_Money_item3_3,I_Money_item3_all=@I_Money_item3_all,
                I_Money_item4_1=@I_Money_item4_1,I_Money_item4_2=@I_Money_item4_2,I_Money_item4_3=@I_Money_item4_3,I_Money_item4_all=@I_Money_item4_all,
                I_Other_Oneself=@I_Other_Oneself,I_Other_Oneself_Money=@I_Other_Oneself_Money,I_Other_Other=@I_Other_Other,I_Other_Other_name=@I_Other_Other_name,I_Other_Other_Money=@I_Other_Other_Money,
                I_Target=@I_Target,I_Summary=@I_Summary,
                I_Finish_item1_1=@I_Finish_item1_1,I_Finish_item1_2=@I_Finish_item1_2,I_Finish_item1_3=@I_Finish_item1_3,I_Finish_item1_all=@I_Finish_item1_all,
                I_Finish_item2_1=@I_Finish_item2_1,I_Finish_item2_2=@I_Finish_item2_2,I_Finish_item2_3=@I_Finish_item2_3,I_Finish_item2_all=@I_Finish_item2_all,
                I_Finish_item3_1=@I_Finish_item3_1,I_Finish_item3_2=@I_Finish_item3_2,I_Finish_item3_3=@I_Finish_item3_3,I_Finish_item3_all=@I_Finish_item3_all,
                I_Finish_item4_1=@I_Finish_item4_1,I_Finish_item4_2=@I_Finish_item4_2,I_Finish_item4_3=@I_Finish_item4_3,I_Finish_item4_all=@I_Finish_item4_all,
                I_Finish_item5_1=@I_Finish_item5_1,I_Finish_item5_2=@I_Finish_item5_2,I_Finish_item5_3=@I_Finish_item5_3,I_Finish_item5_all=@I_Finish_item5_all,
                I_Modifydate=@I_Modifydate,I_ModId=@I_ModId
                where I_ID=@I_ID and I_Guid=@I_GUID
            ";
        }
        
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@I_ID", I_ID);
        oCmd.Parameters.AddWithValue("@I_GUID", I_GUID);
        oCmd.Parameters.AddWithValue("@I_1_Sdate", I_1_Sdate);
        oCmd.Parameters.AddWithValue("@I_1_Edate", I_1_Edate);
        oCmd.Parameters.AddWithValue("@I_2_Sdate", I_2_Sdate);
        oCmd.Parameters.AddWithValue("@I_2_Edate", I_2_Edate);
        oCmd.Parameters.AddWithValue("@I_3_Sdate", I_3_Sdate);
        oCmd.Parameters.AddWithValue("@I_3_Edate", I_3_Edate);
        oCmd.Parameters.AddWithValue("@I_Money_item1_1", I_Money_item1_1);
        oCmd.Parameters.AddWithValue("@I_Money_item1_2", I_Money_item1_2);
        oCmd.Parameters.AddWithValue("@I_Money_item1_3", I_Money_item1_3);
        oCmd.Parameters.AddWithValue("@I_Money_item1_all", I_Money_item1_all);
        oCmd.Parameters.AddWithValue("@I_Money_item2_1", I_Money_item2_1);
        oCmd.Parameters.AddWithValue("@I_Money_item2_2", I_Money_item2_2);
        oCmd.Parameters.AddWithValue("@I_Money_item2_3", I_Money_item2_3);
        oCmd.Parameters.AddWithValue("@I_Money_item2_all", I_Money_item2_all);
        oCmd.Parameters.AddWithValue("@I_Money_item3_1", I_Money_item3_1);
        oCmd.Parameters.AddWithValue("@I_Money_item3_2", I_Money_item3_2);
        oCmd.Parameters.AddWithValue("@I_Money_item3_3", I_Money_item3_3);
        oCmd.Parameters.AddWithValue("@I_Money_item3_all", I_Money_item3_all);

        oCmd.Parameters.AddWithValue("@I_Money_item4_1", I_Money_item4_1);
        oCmd.Parameters.AddWithValue("@I_Money_item4_2", I_Money_item4_2);
        oCmd.Parameters.AddWithValue("@I_Money_item4_3", I_Money_item4_3);
        oCmd.Parameters.AddWithValue("@I_Money_item4_all", I_Money_item4_all);

        oCmd.Parameters.AddWithValue("@I_Other_Oneself", I_Other_Oneself);
        oCmd.Parameters.AddWithValue("@I_Other_Oneself_Money", I_Other_Oneself_Money);
        oCmd.Parameters.AddWithValue("@I_Other_Other", I_Other_Other);
        oCmd.Parameters.AddWithValue("@I_Other_Other_name", I_Other_Other_name);
        oCmd.Parameters.AddWithValue("@I_Other_Other_Money", I_Other_Other_Money);
        oCmd.Parameters.AddWithValue("@I_Target", I_Target);
        oCmd.Parameters.AddWithValue("@I_Summary", I_Summary);
        oCmd.Parameters.AddWithValue("@I_Finish_item1_1", I_Finish_item1_1);
        oCmd.Parameters.AddWithValue("@I_Finish_item1_2", I_Finish_item1_2);
        oCmd.Parameters.AddWithValue("@I_Finish_item1_3", I_Finish_item1_3);
        oCmd.Parameters.AddWithValue("@I_Finish_item1_all", I_Finish_item1_all);
        oCmd.Parameters.AddWithValue("@I_Finish_item2_1", I_Finish_item2_1);
        oCmd.Parameters.AddWithValue("@I_Finish_item2_2", I_Finish_item2_2);
        oCmd.Parameters.AddWithValue("@I_Finish_item2_3", I_Finish_item2_3);
        oCmd.Parameters.AddWithValue("@I_Finish_item2_all", I_Finish_item2_all);
        oCmd.Parameters.AddWithValue("@I_Finish_item3_1", I_Finish_item3_1);
        oCmd.Parameters.AddWithValue("@I_Finish_item3_2", I_Finish_item3_2);
        oCmd.Parameters.AddWithValue("@I_Finish_item3_3", I_Finish_item3_3);
        oCmd.Parameters.AddWithValue("@I_Finish_item3_all", I_Finish_item3_all);
        oCmd.Parameters.AddWithValue("@I_Finish_item4_1", I_Finish_item4_1);
        oCmd.Parameters.AddWithValue("@I_Finish_item4_2", I_Finish_item4_2);
        oCmd.Parameters.AddWithValue("@I_Finish_item4_3", I_Finish_item4_3);
        oCmd.Parameters.AddWithValue("@I_Finish_item4_all", I_Finish_item4_all);
        oCmd.Parameters.AddWithValue("@I_Finish_item5_1", I_Finish_item5_1);
        oCmd.Parameters.AddWithValue("@I_Finish_item5_2", I_Finish_item5_2);
        oCmd.Parameters.AddWithValue("@I_Finish_item5_3", I_Finish_item5_3);
        oCmd.Parameters.AddWithValue("@I_Finish_item5_all", I_Finish_item5_all);
        oCmd.Parameters.AddWithValue("@I_Modifydate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@I_ModId", I_ModId);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    //抓計畫列表 by機關
    public DataTable selectProjectList()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();
        
        if (str_keyword != "")
        {
            sb.Append(@"
                select I_ID,I_Guid,I_1_Sdate,I_3_Edate,I_City,I_Office,I_People,I_Createdate,I_Modifydate,C_Item_cn as cityName,M_Name as personNmae,M_ID,I_Flag
                ,(select I_Guid from ProjectInfo b where ProjectInfo.I_City = b.I_City and I_Status<>'D' and I_Flag='Y' ) as chk_flag
                from ProjectInfo
                left join CodeTable on I_City = C_Item and C_Group='02'
                left join Member on I_People=M_Guid
                where 1=1 and I_Status<>'D'
            ");
            if (M_Competence == "01" || M_Competence == "02")
            {
                sb.Append(@" and I_City = @I_City ");
                oCmd.Parameters.AddWithValue("@I_City", I_City);
            }
            sb.Append(@" and (upper(C_Item_cn) LIKE '%' + upper(@str_keyword) + '%' or upper(I_Office) LIKE '%' + upper(@str_keyword) + '%' or upper(M_Name) LIKE '%' + upper(@str_keyword) + '%' ) ");
            oCmd.Parameters.AddWithValue("@str_keyword", str_keyword);
        }
        else {
            if (M_Competence == "SA")
            {
                //管理者
                sb.Append(@"
                    select I_ID,I_Guid,I_1_Sdate,I_3_Edate,I_City,I_Office,I_People,I_Createdate,I_Modifydate,C_Item_cn as cityName,M_Name as personNmae,I_Flag,M_ID
                    ,(select I_Guid from ProjectInfo b where ProjectInfo.I_City = b.I_City and I_Status<>'D' and I_Flag='Y' ) as chk_flag
                    from ProjectInfo
                    left join CodeTable on I_City = C_Item and C_Group='02'
                    left join Member on I_People=M_Guid
                    where I_Status<>'D'
                    order by I_City ASC
                ");
            }
            else {
                sb.Append(@"
                    select I_ID,I_Guid,I_1_Sdate,I_3_Edate,I_City,I_Office,I_People,I_Createdate,I_Modifydate,C_Item_cn as cityName,M_Name as personNmae,I_Flag,M_ID
                    ,(select I_Guid from ProjectInfo b where ProjectInfo.I_City = b.I_City and I_Status<>'D' and I_Flag='Y' ) as chk_flag
                    from ProjectInfo
                    left join CodeTable on I_City = C_Item and C_Group='02'
                    left join Member on I_People=M_Guid
                    where I_People=@I_People  and I_Status<>'D'
                    union all
                    select I_ID,I_Guid,I_1_Sdate,I_3_Edate,I_City,I_Office,I_People,I_Createdate,I_Modifydate,C_Item_cn as cityName,M_Name as personNmae,I_Flag,M_ID
                    ,(select I_Guid from ProjectInfo b where ProjectInfo.I_City = b.I_City and I_Status<>'D' and I_Flag='Y' ) as chk_flag
                    from ProjectInfo
                    left join CodeTable on I_City = C_Item and C_Group='02'
                    left join Member on I_People=M_Guid
                    where I_People!=@I_People and I_City=@I_City and I_Status<>'D'
                ");
                oCmd.Parameters.AddWithValue("@I_City", I_City);
                oCmd.Parameters.AddWithValue("@I_People", I_People);
            }
            
            
        }
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();


        oda.Fill(ds);
        return ds;
    }

    //insert 資料進去
    public void addProjectInfo()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);

        oCmd.CommandText = @"
            declare @new_Guid nvarchar(50)
            declare @I_City nvarchar(10)
            declare @I_office nvarchar(10)
            declare @I_Boss nvarchar(50)
            select @new_Guid= newid()
            select @I_City=M_City,@I_office=M_office,@I_Boss=M_Manager_ID from Member where M_Guid=@I_People and M_Status<>'D'
            insert into ProjectInfo(I_Guid,I_City,I_Office,I_People,I_1_Sdate,I_1_Edate,I_2_Sdate,I_2_Edate,I_3_Sdate,I_3_Edate
                ,I_Money_item1_1,I_Money_item1_2,I_Money_item1_3,I_Money_item1_all
                ,I_Money_item2_1,I_Money_item2_2,I_Money_item2_3,I_Money_item2_all
                ,I_Money_item3_1,I_Money_item3_2,I_Money_item3_3,I_Money_item3_all
                ,I_Money_item4_1,I_Money_item4_2,I_Money_item4_3,I_Money_item4_all
                ,I_Other_Oneself,I_Other_Oneself_Money,I_Other_Other,I_Other_Other_name,I_Other_Other_Money,I_Target,I_Summary
                ,I_Finish_item1_1,I_Finish_item1_2,I_Finish_item1_3,I_Finish_item1_all
                ,I_Finish_item2_1,I_Finish_item2_2,I_Finish_item2_3,I_Finish_item2_all
                ,I_Finish_item3_1,I_Finish_item3_2,I_Finish_item3_3,I_Finish_item3_all
                ,I_Finish_item4_1,I_Finish_item4_2,I_Finish_item4_3,I_Finish_item4_all
                ,I_Finish_item5_1,I_Finish_item5_2,I_Finish_item5_3,I_Finish_item5_all
                ,I_Boss,I_Createdate,I_Modifydate,I_CreateId,I_ModId,I_Status)
            values(@I_GUID,@I_City,@I_Office,@I_People,@I_1_Sdate,@I_1_Edate,@I_2_Sdate,@I_2_Edate,@I_3_Sdate,@I_3_Edate
                ,@I_Money_item1_1,@I_Money_item1_2,@I_Money_item1_3,@I_Money_item1_all
                ,@I_Money_item2_1,@I_Money_item2_2,@I_Money_item2_3,@I_Money_item2_all
                ,@I_Money_item3_1,@I_Money_item3_2,@I_Money_item3_3,@I_Money_item3_all
                ,@I_Money_item4_1,@I_Money_item4_2,@I_Money_item4_3,@I_Money_item4_all
                ,@I_Other_Oneself,@I_Other_Oneself_Money,@I_Other_Other,@I_Other_Other_name,@I_Other_Other_Money,@I_Target,@I_Summary
                ,@I_Finish_item1_1,@I_Finish_item1_2,@I_Finish_item1_3,@I_Finish_item1_all
                ,@I_Finish_item2_1,@I_Finish_item2_2,@I_Finish_item2_3,@I_Finish_item2_all
                ,@I_Finish_item3_1,@I_Finish_item3_2,@I_Finish_item3_3,@I_Finish_item3_all
                ,@I_Finish_item4_1,@I_Finish_item4_2,@I_Finish_item4_3,@I_Finish_item4_all
                ,@I_Finish_item5_1,@I_Finish_item5_2,@I_Finish_item5_3,@I_Finish_item5_all
                ,@I_Boss,@I_Createdate,@I_Createdate,@I_CreateId,@I_CreateId,'A')
        ";
        
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@I_People", I_People);
        oCmd.Parameters.AddWithValue("@I_GUID", I_GUID);
        oCmd.Parameters.AddWithValue("@I_1_Sdate", I_1_Sdate);
        oCmd.Parameters.AddWithValue("@I_1_Edate", I_1_Edate);
        oCmd.Parameters.AddWithValue("@I_2_Sdate", I_2_Sdate);
        oCmd.Parameters.AddWithValue("@I_2_Edate", I_2_Edate);
        oCmd.Parameters.AddWithValue("@I_3_Sdate", I_3_Sdate);
        oCmd.Parameters.AddWithValue("@I_3_Edate", I_3_Edate);
        oCmd.Parameters.AddWithValue("@I_Money_item1_1", I_Money_item1_1);
        oCmd.Parameters.AddWithValue("@I_Money_item1_2", I_Money_item1_2);
        oCmd.Parameters.AddWithValue("@I_Money_item1_3", I_Money_item1_3);
        oCmd.Parameters.AddWithValue("@I_Money_item1_all", I_Money_item1_all);
        oCmd.Parameters.AddWithValue("@I_Money_item2_1", I_Money_item2_1);
        oCmd.Parameters.AddWithValue("@I_Money_item2_2", I_Money_item2_2);
        oCmd.Parameters.AddWithValue("@I_Money_item2_3", I_Money_item2_3);
        oCmd.Parameters.AddWithValue("@I_Money_item2_all", I_Money_item2_all);
        oCmd.Parameters.AddWithValue("@I_Money_item3_1", I_Money_item3_1);
        oCmd.Parameters.AddWithValue("@I_Money_item3_2", I_Money_item3_2);
        oCmd.Parameters.AddWithValue("@I_Money_item3_3", I_Money_item3_3);
        oCmd.Parameters.AddWithValue("@I_Money_item3_all", I_Money_item3_all);

        oCmd.Parameters.AddWithValue("@I_Money_item4_1", I_Money_item4_1);
        oCmd.Parameters.AddWithValue("@I_Money_item4_2", I_Money_item4_2);
        oCmd.Parameters.AddWithValue("@I_Money_item4_3", I_Money_item4_3);
        oCmd.Parameters.AddWithValue("@I_Money_item4_all", I_Money_item4_all);

        oCmd.Parameters.AddWithValue("@I_Other_Oneself", I_Other_Oneself);
        oCmd.Parameters.AddWithValue("@I_Other_Oneself_Money", I_Other_Oneself_Money);
        oCmd.Parameters.AddWithValue("@I_Other_Other", I_Other_Other);
        oCmd.Parameters.AddWithValue("@I_Other_Other_name", I_Other_Other_name);
        oCmd.Parameters.AddWithValue("@I_Other_Other_Money", I_Other_Other_Money);
        oCmd.Parameters.AddWithValue("@I_Target", I_Target);
        oCmd.Parameters.AddWithValue("@I_Summary", I_Summary);
        oCmd.Parameters.AddWithValue("@I_Finish_item1_1", I_Finish_item1_1);
        oCmd.Parameters.AddWithValue("@I_Finish_item1_2", I_Finish_item1_2);
        oCmd.Parameters.AddWithValue("@I_Finish_item1_3", I_Finish_item1_3);
        oCmd.Parameters.AddWithValue("@I_Finish_item1_all", I_Finish_item1_all);
        oCmd.Parameters.AddWithValue("@I_Finish_item2_1", I_Finish_item2_1);
        oCmd.Parameters.AddWithValue("@I_Finish_item2_2", I_Finish_item2_2);
        oCmd.Parameters.AddWithValue("@I_Finish_item2_3", I_Finish_item2_3);
        oCmd.Parameters.AddWithValue("@I_Finish_item2_all", I_Finish_item2_all);
        oCmd.Parameters.AddWithValue("@I_Finish_item3_1", I_Finish_item3_1);
        oCmd.Parameters.AddWithValue("@I_Finish_item3_2", I_Finish_item3_2);
        oCmd.Parameters.AddWithValue("@I_Finish_item3_3", I_Finish_item3_3);
        oCmd.Parameters.AddWithValue("@I_Finish_item3_all", I_Finish_item3_all);
        oCmd.Parameters.AddWithValue("@I_Finish_item4_1", I_Finish_item4_1);
        oCmd.Parameters.AddWithValue("@I_Finish_item4_2", I_Finish_item4_2);
        oCmd.Parameters.AddWithValue("@I_Finish_item4_3", I_Finish_item4_3);
        oCmd.Parameters.AddWithValue("@I_Finish_item4_all", I_Finish_item4_all);
        oCmd.Parameters.AddWithValue("@I_Finish_item5_1", I_Finish_item5_1);
        oCmd.Parameters.AddWithValue("@I_Finish_item5_2", I_Finish_item5_2);
        oCmd.Parameters.AddWithValue("@I_Finish_item5_3", I_Finish_item5_3);
        oCmd.Parameters.AddWithValue("@I_Finish_item5_all", I_Finish_item5_all);
        oCmd.Parameters.AddWithValue("@I_Createdate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@I_CreateId", I_ModId);
        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    //複製計畫 update
    public void copyProjectByIDUpdate()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"
            update ProjectInfo set
            I_1_Sdate=a.I_1_Sdate,I_1_Edate=a.I_1_Edate,I_2_Sdate=a.I_2_Sdate,I_2_Edate=a.I_2_Edate,I_3_Sdate=a.I_3_Sdate,I_3_Edate=a.I_3_Edate,
            I_Money_item1_1=a.I_Money_item1_1,I_Money_item1_2=a.I_Money_item1_2,I_Money_item1_3=a.I_Money_item1_3,I_Money_item1_all=a.I_Money_item1_all,
            I_Money_item2_1=a.I_Money_item2_1,I_Money_item2_2=a.I_Money_item2_2,I_Money_item2_3=a.I_Money_item2_3,I_Money_item2_all=a.I_Money_item2_all,
            I_Money_item3_1=a.I_Money_item3_1,I_Money_item3_2=a.I_Money_item3_2,I_Money_item3_3=a.I_Money_item3_3,I_Money_item3_all=a.I_Money_item3_all,
            I_Other_Oneself=a.I_Other_Oneself,I_Other_Oneself_Money=a.I_Other_Oneself_Money,I_Other_Other=a.I_Other_Other,I_Other_Other_name=a.I_Other_Other_name,I_Other_Other_Money=a.I_Other_Other_Money,
            I_Target=a.I_Target,I_Summary=a.I_Summary,
            I_Finish_item1_1=a.I_Finish_item1_1,I_Finish_item1_2=a.I_Finish_item1_2,I_Finish_item1_3=a.I_Finish_item1_3,I_Finish_item1_all=a.I_Finish_item1_all,
            I_Finish_item2_1=a.I_Finish_item2_1,I_Finish_item2_2=a.I_Finish_item2_2,I_Finish_item2_3=a.I_Finish_item2_3,I_Finish_item2_all=a.I_Finish_item2_all,
            I_Finish_item3_1=a.I_Finish_item3_1,I_Finish_item3_2=a.I_Finish_item3_2,I_Finish_item3_3=a.I_Finish_item3_3,I_Finish_item3_all=a.I_Finish_item3_all,
            I_Finish_item4_1=a.I_Finish_item4_1,I_Finish_item4_2=a.I_Finish_item4_2,I_Finish_item4_3=a.I_Finish_item4_3,I_Finish_item4_all=a.I_Finish_item4_all,
            I_Finish_item5_1=a.I_Finish_item5_1,I_Finish_item5_2=a.I_Finish_item5_2,I_Finish_item5_3=a.I_Finish_item5_3,I_Finish_item5_all=a.I_Finish_item5_all,
            I_Modifydate=@I_Modifydate
            from (
                select  b.I_1_Sdate,b.I_1_Edate,b.I_2_Sdate,b.I_2_Edate,b.I_3_Sdate,b.I_3_Edate,
                        b.I_Money_item1_1,b.I_Money_item1_2,b.I_Money_item1_3,b.I_Money_item1_all,
                        b.I_Money_item2_1,b.I_Money_item2_2,b.I_Money_item2_3,b.I_Money_item2_all,
                        b.I_Money_item3_1,b.I_Money_item3_2,b.I_Money_item3_3,b.I_Money_item3_all,
                        b.I_Other_Oneself,b.I_Other_Oneself_Money,b.I_Other_Other,b.I_Other_Other_name,b.I_Other_Other_Money,
                        b.I_Target,b.I_Summary,
                        b.I_Finish_item1_1,b.I_Finish_item1_2,b.I_Finish_item1_3,b.I_Finish_item1_all,
                        b.I_Finish_item2_1,b.I_Finish_item2_2,b.I_Finish_item2_3,b.I_Finish_item2_all,
                        b.I_Finish_item3_1,b.I_Finish_item3_2,b.I_Finish_item3_3,b.I_Finish_item3_all,
                        b.I_Finish_item4_1,b.I_Finish_item4_2,b.I_Finish_item4_3,b.I_Finish_item4_all,
                        b.I_Finish_item5_1,b.I_Finish_item5_2,b.I_Finish_item5_3,b.I_Finish_item5_all
                from ProjectInfo b
                where b.I_ID = @I_ID
            ) a
            where I_People=@I_People
        ";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DateTime dtnow = DateTime.Now;
        oCmd.Parameters.AddWithValue("@I_Modifydate", dtnow);
        oCmd.Parameters.AddWithValue("@I_People", I_People);
        oCmd.Parameters.AddWithValue("@I_ID", I_ID);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    //複製計畫 insert
    public void copyProjectByIDInsert()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        if (M_Competence == "SA")
        {
            //管理者複製不需要複製人員
            oCmd.CommandText = @"
                insert into ProjectInfo(
                    I_Guid,
                    I_1_Sdate,I_1_Edate,I_2_Sdate,I_2_Edate,I_3_Sdate,I_3_Edate,
                    I_Money_item1_1,I_Money_item1_2,I_Money_item1_3,I_Money_item1_all,
                    I_Money_item2_1,I_Money_item2_2,I_Money_item2_3,I_Money_item2_all,
                    I_Money_item3_1,I_Money_item3_2,I_Money_item3_3,I_Money_item3_all,
                    I_Other_Oneself,I_Other_Oneself_Money,I_Other_Other,I_Other_Other_name,I_Other_Other_Money,
                    I_Target,I_Summary,
                    I_Finish_item1_1,I_Finish_item1_2,I_Finish_item1_3,I_Finish_item1_all,
                    I_Finish_item2_1,I_Finish_item2_2,I_Finish_item2_3,I_Finish_item2_all,
                    I_Finish_item3_1,I_Finish_item3_2,I_Finish_item3_3,I_Finish_item3_all,
                    I_Finish_item4_1,I_Finish_item4_2,I_Finish_item4_3,I_Finish_item4_all,
                    I_Finish_item5_1,I_Finish_item5_2,I_Finish_item5_3,I_Finish_item5_all,
                    I_Createdate,I_Modifydate,I_Status,I_CreateId,I_ModId
                )
                select  NEWID(),
                        I_1_Sdate,I_1_Edate,I_2_Sdate,I_2_Edate,I_3_Sdate,I_3_Edate,
                        I_Money_item1_1,I_Money_item1_2,I_Money_item1_3,I_Money_item1_all,
                        I_Money_item2_1,I_Money_item2_2,I_Money_item2_3,I_Money_item2_all,
                        I_Money_item3_1,I_Money_item3_2,I_Money_item3_3,I_Money_item3_all,
                        I_Other_Oneself,I_Other_Oneself_Money,I_Other_Other,I_Other_Other_name,I_Other_Other_Money,
                        I_Target,I_Summary,
                        I_Finish_item1_1,I_Finish_item1_2,I_Finish_item1_3,I_Finish_item1_all,
                        I_Finish_item2_1,I_Finish_item2_2,I_Finish_item2_3,I_Finish_item2_all,
                        I_Finish_item3_1,I_Finish_item3_2,I_Finish_item3_3,I_Finish_item3_all,
                        I_Finish_item4_1,I_Finish_item4_2,I_Finish_item4_3,I_Finish_item4_all,
                        I_Finish_item5_1,I_Finish_item5_2,I_Finish_item5_3,I_Finish_item5_all,
                        @I_Modifydate,@I_Modifydate,'A',@I_People,@I_People
                from ProjectInfo 
                where I_ID = @I_ID
            ";
        }
        else {
            oCmd.CommandText = @"
                insert into ProjectInfo(
                    I_Guid,I_City,I_Office,I_People,
                    I_1_Sdate,I_1_Edate,I_2_Sdate,I_2_Edate,I_3_Sdate,I_3_Edate,
                    I_Money_item1_1,I_Money_item1_2,I_Money_item1_3,I_Money_item1_all,
                    I_Money_item2_1,I_Money_item2_2,I_Money_item2_3,I_Money_item2_all,
                    I_Money_item3_1,I_Money_item3_2,I_Money_item3_3,I_Money_item3_all,
                    I_Other_Oneself,I_Other_Oneself_Money,I_Other_Other,I_Other_Other_name,I_Other_Other_Money,
                    I_Target,I_Summary,
                    I_Finish_item1_1,I_Finish_item1_2,I_Finish_item1_3,I_Finish_item1_all,
                    I_Finish_item2_1,I_Finish_item2_2,I_Finish_item2_3,I_Finish_item2_all,
                    I_Finish_item3_1,I_Finish_item3_2,I_Finish_item3_3,I_Finish_item3_all,
                    I_Finish_item4_1,I_Finish_item4_2,I_Finish_item4_3,I_Finish_item4_all,
                    I_Finish_item5_1,I_Finish_item5_2,I_Finish_item5_3,I_Finish_item5_all,
                    I_Boss,I_Createdate,I_Modifydate,I_Status,I_CreateId,I_ModId
                )
                select  NEWID(),M_City,M_Office,@I_People,
                        b.I_1_Sdate,b.I_1_Edate,b.I_2_Sdate,b.I_2_Edate,b.I_3_Sdate,b.I_3_Edate,
                        b.I_Money_item1_1,b.I_Money_item1_2,b.I_Money_item1_3,b.I_Money_item1_all,
                        b.I_Money_item2_1,b.I_Money_item2_2,b.I_Money_item2_3,b.I_Money_item2_all,
                        b.I_Money_item3_1,b.I_Money_item3_2,b.I_Money_item3_3,b.I_Money_item3_all,
                        b.I_Other_Oneself,b.I_Other_Oneself_Money,b.I_Other_Other,b.I_Other_Other_name,b.I_Other_Other_Money,
                        b.I_Target,b.I_Summary,
                        b.I_Finish_item1_1,b.I_Finish_item1_2,b.I_Finish_item1_3,b.I_Finish_item1_all,
                        b.I_Finish_item2_1,b.I_Finish_item2_2,b.I_Finish_item2_3,b.I_Finish_item2_all,
                        b.I_Finish_item3_1,b.I_Finish_item3_2,b.I_Finish_item3_3,b.I_Finish_item3_all,
                        b.I_Finish_item4_1,b.I_Finish_item4_2,b.I_Finish_item4_3,b.I_Finish_item4_all,
                        b.I_Finish_item5_1,b.I_Finish_item5_2,b.I_Finish_item5_3,b.I_Finish_item5_all,
                        M_Manager_ID,@I_Modifydate,@I_Modifydate,'A',@I_People,@I_People
                from ProjectInfo b left join Member on M_Guid=@I_People
                where b.I_ID = @I_ID
            ";
            
        }
        
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DateTime dtnow = DateTime.Now;
        oCmd.Parameters.AddWithValue("@I_People", I_People);
        oCmd.Parameters.AddWithValue("@I_Modifydate", dtnow);
        oCmd.Parameters.AddWithValue("@I_ID", I_ID);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    //定稿 insert 一筆到ProjectInfo_Check
    public void submitProject()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"
            insert into ProjectInfo_Check(IC_GUID,IC_City,IC_Office,IC_People,IC_1_Sdate,IC_1_edate,IC_2_Sdate,IC_2_edate,IC_3_Sdate,IC_3_edate,
            IC_Money_item1_1,IC_Money_item1_2,IC_Money_item1_3,IC_Money_item1_all,
            IC_Money_item2_1,IC_Money_item2_2,IC_Money_item2_3,IC_Money_item2_all,
            IC_Money_item3_1,IC_Money_item3_2,IC_Money_item3_3,IC_Money_item3_all,
            IC_Other_Oneself,IC_Other_Oneself_Money,IC_Other_Other,IC_Other_Other_name,IC_Other_Other_Money,IC_Target,IC_Summary,
            IC_Finish_item1_1,IC_Finish_item1_2,IC_Finish_item1_3,IC_Finish_item1_all,
            IC_Finish_item2_1,IC_Finish_item2_2,IC_Finish_item2_3,IC_Finish_item2_all,
            IC_Finish_item3_1,IC_Finish_item3_2,IC_Finish_item3_3,IC_Finish_item3_all,
            IC_Finish_item4_1,IC_Finish_item4_2,IC_Finish_item4_3,IC_Finish_item4_all,
            IC_Finish_item5_1,IC_Finish_item5_2,IC_Finish_item5_3,IC_Finish_item5_all,
            IC_Boss,IC_Submitdate,IC_Status,IC_Flag
            )
            select NEWID(),I_City,I_Office,I_People,I_1_Sdate,I_1_edate,I_2_Sdate,I_2_edate,I_3_Sdate,I_3_edate,
            I_Money_item1_1,I_Money_item1_2,I_Money_item1_3,I_Money_item1_all,
            I_Money_item2_1,I_Money_item2_2,I_Money_item2_3,I_Money_item2_all,
            I_Money_item3_1,I_Money_item3_2,I_Money_item3_3,I_Money_item3_all,
            I_Other_Oneself,I_Other_Oneself_Money,I_Other_Other,I_Other_Other_name,I_Other_Other_Money,I_Target,I_Summary,
            I_Finish_item1_1,I_Finish_item1_2,I_Finish_item1_3,I_Finish_item1_all,
            I_Finish_item2_1,I_Finish_item2_2,I_Finish_item2_3,I_Finish_item2_all,
            I_Finish_item3_1,I_Finish_item3_2,I_Finish_item3_3,I_Finish_item3_all,
            I_Finish_item4_1,I_Finish_item4_2,I_Finish_item4_3,I_Finish_item4_all,
            I_Finish_item5_1,I_Finish_item5_2,I_Finish_item5_3,I_Finish_item5_all,
            I_Boss,@IC_Submitdate,'A',0
            from ProjectInfo
            where I_People=@I_People

            update ProjectInfo set I_Flag='Y' where I_People=@I_People
        ";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DateTime dtnow = DateTime.Now;
        oCmd.Parameters.AddWithValue("@IC_Submitdate", dtnow);
        oCmd.Parameters.AddWithValue("@I_People", I_People);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    //判斷現在有沒有定稿的資料 by機關
    public DataTable selectCheckProject()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();
        
            sb.Append(@"
                select * from ProjectInfo where I_City=@I_City and I_Flag='Y'  and I_Status<>'D'
            ");
            oCmd.Parameters.AddWithValue("@I_City", I_City);

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();


        oda.Fill(ds);
        return ds;
    }

    //撈所有承辦人
    public DataTable selPeople()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();
        
        sb.Append(@"
            select M_ID,M_Guid,M_Name,M_City,C_Item_cn as cityName
            from Member
            left join CodeTable on M_City = C_Item and C_Group='02'
            where M_Status<>'D'
            order by M_City ASC
        ");
       
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();


        oda.Fill(ds);
        return ds;
    }

    //撈管理者剛剛複製的計畫資料I_ID
    public DataTable selCopyIIDForAdmin()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
            select top 1 I_ID from ProjectInfo where I_CreateId=@I_CreateId order by I_Createdate DESC
        ");
        oCmd.Parameters.AddWithValue("@I_CreateId", I_People);
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();


        oda.Fill(ds);
        return ds;
    }

    //撈有沒有該承辦員的計畫基本資料
    public DataTable selProjectByPeople()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
            select top 1 I_ID from ProjectInfo where I_People=@I_People and I_Status<>'D'
        ");
        oCmd.Parameters.AddWithValue("@I_People", I_People);
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();


        oda.Fill(ds);
        return ds;
    }


    public DataSet getPreview(string ProjectID)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
--基本資料
select *,(select C_Item_cn from CodeTable where C_Group='02' and C_Item=I_City) CityName 
from ProjectInfo 
left join ProjectDate on I_City=PD_Type
where I_Guid=@ProjectID and I_Status='A'

--承辦人資料
declare @mguid nvarchar(50)=(select I_People from ProjectInfo where I_Guid=@ProjectID and I_Status='A')
select * from Member where M_Guid=@mguid and M_Status='A'

--承辦主管
declare @mrguid nvarchar(50)=(select M_Manager_ID from Member where M_Guid=@mguid and M_Status='A')
select * from Member where M_Guid=@mrguid and M_Status='A' ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@ProjectID", ProjectID);
        oda.Fill(ds);
        return ds;
    }

    //送出定稿
    public DataTable submit_Info()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"update ProjectInfo set 
I_Flag='Y',
I_Modifydate=@I_Modifydate,
I_ModId=@I_ModId 
where I_Guid=@I_Guid ");

        oCmd.Parameters.AddWithValue("@I_Guid", I_GUID);
        oCmd.Parameters.AddWithValue("@I_Modifydate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@I_ModId", I_ModId);
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();


        oda.Fill(ds);
        return ds;
    }

    public DataTable getFlag()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select I_Flag from ProjectInfo where I_Guid=@I_Guid ");

        oCmd.Parameters.AddWithValue("@I_Guid", I_GUID);
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();


        oda.Fill(ds);
        return ds;
    }

    public DataTable CityFlagCount()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select COUNT(*) num from ProjectInfo where I_City=@I_City and I_Flag='Y' and I_Status='A' ");

        oCmd.Parameters.AddWithValue("@I_City", I_City);
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();


        oda.Fill(ds);
        return ds;
    }

    public DataSet getAdminCityFlag()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
--該縣市有無定稿資料
select COUNT(*) num from ProjectInfo
where I_City=(select I_City from ProjectInfo where I_Guid=@I_Guid) and I_Flag='Y' and I_Status='A' 

--該筆定稿資料
select I_Flag from ProjectInfo
where I_Guid=@I_GUID
 and I_Status='A' 
");

        oCmd.Parameters.AddWithValue("@I_Guid", I_GUID);
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();


        oda.Fill(ds);
        return ds;
    }


    //撈有無定稿資料 by 機關代碼
    public DataTable getFlagByCity()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select I_Guid from ProjectInfo where I_City=@I_City and I_Status<>'D' and I_Flag='Y' ");

        oCmd.Parameters.AddWithValue("@I_City", I_City);
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();


        oda.Fill(ds);
        return ds;
    }
    
    public DataTable checkCity()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select I_City from ProjectInfo where I_Guid=@I_Guid and I_Status='A' ");

        oCmd.Parameters.AddWithValue("@I_Guid", I_GUID);
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();


        oda.Fill(ds);
        return ds;
    }

    public DataSet getProjectListBySA(string pStart,string pEnd)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT COUNT(*) total from ProjectInfo 
		left join Member on I_People=M_Guid
        left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=I_City
        left join ProjectDate on PD_Type=I_City
where I_Status='A' ");

        if (str_keyword != "")
        {
            sb.Append(@"and ((upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(I_Office) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Name) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@"select * from (
	select ROW_NUMBER() over (order by I_City,I_Modifydate desc) itemNo,
	city_type.C_Item_cn City,a.I_ID,a.I_Guid,a.I_City,a.I_Office,a.I_People,a.I_Flag
	,a.I_Modifydate,PD_StartDate,PD_EndDate,M_ID,M_Name,M_Office
	,(select COUNT(*) total from ProjectInfo b where b.I_City=a.I_City and b.I_Flag='Y') CityFlag
	from ProjectInfo a
	left join Member on I_People=M_Guid
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=I_City
    left join ProjectDate on PD_Type=I_City
	where I_Status='A'  ");
        

        if (str_keyword != "")
        {
            sb.Append(@"and ((upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(I_Office) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Name) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.Parameters.AddWithValue("@KeyWord", str_keyword);
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oda.Fill(ds);
        return ds;
    }

    public DataTable getProjectListByPerson()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select 
I_ID,I_Guid,city_type.C_Item_cn City,I_City,I_Office,I_People,I_Flag,I_Modifydate,PD_StartDate,PD_EndDate,M_ID,M_Name,M_Office
from ProjectInfo
left join ProjectDate on I_City=PD_Type
left join Member on I_People=M_Guid
left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=I_City
where I_Status='A' and I_People=@I_People
union all
select 
I_ID,I_Guid,city_type.C_Item_cn City,I_City,I_Office,I_People,I_Flag,I_Modifydate,PD_StartDate,PD_EndDate,M_ID,M_Name,M_Office
from ProjectInfo
left join ProjectDate on I_City=PD_Type
left join Member on I_People=M_Guid
left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=I_City
where I_City=@I_City and I_People<>@I_People ");
        
        oCmd.Parameters.AddWithValue("@I_People", I_People);
        oCmd.Parameters.AddWithValue("@I_City", I_City);
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oda.Fill(ds);
        return ds;
    }

    public DataTable getProjectListByManager()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select 
I_ID,I_Guid,city_type.C_Item_cn City,I_City,I_Office,I_People,I_Flag,I_Modifydate,PD_StartDate,PD_EndDate,M_ID,M_Name,M_Office
from ProjectInfo
left join ProjectDate on I_City=PD_Type
left join Member on I_People=M_Guid
left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=I_City
where I_Status='A' and I_City=@I_City 
order by I_Modifydate desc,I_ID desc ");
        
        oCmd.Parameters.AddWithValue("@I_City", I_City);
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();

        oda.Fill(ds);
        return ds;
    }

    //計畫退回草稿
    public void backProject()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"
            update ProjectInfo set I_Flag='' where I_ID=@I_ID and I_Status='A'
        ";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DateTime dtnow = DateTime.Now;
        oCmd.Parameters.AddWithValue("@I_ID", I_ID);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }
}