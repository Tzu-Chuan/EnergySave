using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;

/// <summary>
/// News_DB 的摘要描述
/// </summary>
public class News_DB
{
    #region 私用
    string N_ID = string.Empty;
    string N_Guid = string.Empty;
    string N_Date = string.Empty;
    string N_Title = string.Empty;
    string N_Content = string.Empty;
    string N_CreateId = string.Empty;
    DateTime N_CreateDate;
    string N_ModId = string.Empty;
    DateTime N_ModDate;
    string N_Status = string.Empty;
    string N_Flag = string.Empty;
    string strKeyWord = string.Empty;
    string strDate = string.Empty;
    #endregion
    #region 公用
    public string _N_ID
    {
        set { N_ID = value; }
    }
    public string _N_Guid
    {
        set { N_Guid = value; }
    }
    public string _N_Date
    {
        set { N_Date = value; }
    }
    public string _N_Title
    {
        set { N_Title = value; }
    }
    public string _N_Content
    {
        set { N_Content = value; }
    }
    public string _N_CreateId
    {
        set { N_CreateId = value; }
    }
    public string _N_ModId
    {
        set { N_ModId = value; }
    }
    public string _N_Status
    {
        set { N_Status = value; }
    }
    public string _N_Flag
    {
        set { N_Flag = value; }
    }
    public DateTime _N_CreateDate
    {
        set { N_CreateDate = value; }
    }
    
    public DateTime _N_ModDate
    {
        set { N_ModDate = value; }
    }
    public string _strKeyWord
    {
        set { strKeyWord = value; }
    }
    public string _strDate
    {
        set { strDate = value; }
    }
    #endregion

    public News_DB()
    {
        //
        // TODO: 在這裡新增建構函式邏輯
        //
    }

    //取得公告列表資料
    public DataSet getNewsList(string pStart, string pEnd)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select * into #tmpAll from News where N_Status='A'
");
        if (strKeyWord != "")
        {
            sb.Append(@" and ((upper(N_Title) LIKE '%' + upper(@strKeyWord) + '%') or (upper(N_Content) LIKE '%' + upper(@strKeyWord) + '%')) ");
        }
        if (N_Date != "") {
            sb.Append(@" and N_Date=@N_Date ");
        }
        sb.Append(@" order by N_Date desc,N_CreateDate desc ");

        sb.Append(@"
        --總筆數
select count(*) as total from #tmpAll

--分頁資料
select * from(
           select ROW_NUMBER() over(order by N_Date desc, N_CreateDate desc) itemNo,#tmpAll.*
		   from #tmpAll
)#tmp where itemNo between @pStart and @pEnd

drop table #tmpAll 
");
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@strKeyWord", strKeyWord);
        oCmd.Parameters.AddWithValue("@N_Date", N_Date);
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oda.Fill(ds);
        return ds;
    }

    //取得公告資料 BY ID
    public DataTable getNewsByID()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select * from News where N_ID=@N_ID
");


        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable dt = new DataTable();
        
        oCmd.Parameters.AddWithValue("@N_ID", N_ID);
        oda.Fill(dt);
        return dt;
    }

    //新增公告
    public void addNews()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"
insert into News (N_Guid,N_Date,N_Title,N_Content,N_CreateId,N_ModId,N_CreateDate,N_ModDate,N_Status,N_Flag) 
values (@N_Guid,@N_Date,@N_Title,@N_Content,@N_CreateId,@N_ModId,@N_CreateDate,@N_ModDate,@N_Status,@N_Flag) 
";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@N_Guid", N_Guid);
        oCmd.Parameters.AddWithValue("@N_Date", N_Date);
        oCmd.Parameters.AddWithValue("@N_Title", N_Title);
        oCmd.Parameters.AddWithValue("@N_Content", N_Content);
        oCmd.Parameters.AddWithValue("@N_CreateId", N_CreateId);
        oCmd.Parameters.AddWithValue("@N_ModId", N_CreateId);
        oCmd.Parameters.AddWithValue("@N_CreateDate", N_CreateDate);
        oCmd.Parameters.AddWithValue("@N_ModDate", N_CreateDate);
        oCmd.Parameters.AddWithValue("@N_Status", "A");
        oCmd.Parameters.AddWithValue("@N_Flag", "");

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    //修改公告
    public void updateNews()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"
update News set N_Date=@N_Date,N_Title=@N_Title,N_Content=@N_Content,N_ModId=@N_ModId,N_ModDate=@N_ModDate where N_ID=@N_ID
";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@N_ID", N_ID);
        oCmd.Parameters.AddWithValue("@N_Date", N_Date);
        oCmd.Parameters.AddWithValue("@N_Title", N_Title);
        oCmd.Parameters.AddWithValue("@N_Content", N_Content);
        oCmd.Parameters.AddWithValue("@N_ModId", N_ModId);
        oCmd.Parameters.AddWithValue("@N_ModDate", N_ModDate);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    //刪除公告
    public void deleteNews()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"
update News set N_Status='D' where N_ID=@N_ID 
";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@N_ID", N_ID);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }
}