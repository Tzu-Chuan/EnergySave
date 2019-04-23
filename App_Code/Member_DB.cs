using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Configuration;

/// <summary>
/// Member_DB 的摘要描述
/// </summary>
public class Member_DB
{
    string KeyWord = string.Empty;
    public string _KeyWord
    {
        set { KeyWord = value; }
    }
    #region 私用
    string M_ID = string.Empty;
    string M_Guid = string.Empty;
    string M_Account = string.Empty;
    string M_Pwd = string.Empty;
    string M_Name = string.Empty;
    string M_JobTitle = string.Empty;
    string M_Tel = string.Empty;
    string M_Ext = string.Empty;
    string M_Fax = string.Empty;
    string M_Phone = string.Empty;
    string M_Email = string.Empty;
    string M_Addr = string.Empty;
    string M_City = string.Empty;
    string M_Office = string.Empty;
    string M_Competence = string.Empty;
    string M_Manager_ID = string.Empty;
    string M_CreateId = string.Empty;
    string M_ModId = string.Empty;
    string M_Status = string.Empty;

    DateTime M_CreateDate;
    DateTime M_ModDate;
    #endregion
    #region 公用
    public string _M_ID
    {
        set { M_ID = value; }
    }
    public string _M_Guid
    {
        set { M_Guid = value; }
    }
    public string _M_Account
    {
        set { M_Account = value; }
    }
    public string _M_Pwd
    {
        set { M_Pwd = value; }
    }
    public string _M_Name
    {
        set { M_Name = value; }
    }
    public string _M_JobTitle
    {
        set { M_JobTitle = value; }
    }
    public string _M_Tel
    {
        set { M_Tel = value; }
    }
    public string _M_Ext
    {
        set { M_Ext = value; }
    }
    public string _M_Fax
    {
        set { M_Fax = value; }
    }
    public string _M_Phone
    {
        set { M_Phone = value; }
    }
    public string _M_Email
    {
        set { M_Email = value; }
    }
    public string _M_Addr
    {
        set { M_Addr = value; }
    }
    public string _M_City
    {
        set { M_City = value; }
    }
    public string _M_Office
    {
        set { M_Office = value; }
    }
    public string _M_Competence
    {
        set { M_Competence = value; }
    }
    public string _M_Manager_ID
    {
        set { M_Manager_ID = value; }
    }
    public string _M_CreateId
    {
        set { M_CreateId = value; }
    }
    public string _M_ModId
    {
        set { M_ModId = value; }
    }
    public string _M_Status
    {
        set { M_Status = value; }
    }
    public DateTime _M_CreateDate
    {
        set { M_CreateDate = value; }
    }
    public DateTime _M_ModDate
    {
        set { M_ModDate = value; }
    }
    #endregion

    public DataSet getMemberList(string pStart, string pEnd, string sortStr)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT COUNT(*) total from Member 
        left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
        left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
where M_Status<>'D' ");
        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Account) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_JobTitle) LIKE '%' + upper(@KeyWord) + '%')
 or (upper(M_Tel) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Ext) LIKE '%' + upper(@KeyWord) + '%') or (upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Office) LIKE '%' + upper(@KeyWord) + '%')
 or (upper(M_Email) LIKE '%' + upper(@KeyWord) + '%') or (upper(comp_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@"select * from (
           select ROW_NUMBER() over (order by " + sortStr + @") itemNo,
        Member.*,city_type.C_Item_cn City,comp_type.C_Item_cn Comp from Member
        left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
        left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
         where M_Status<>'D' ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Account) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_JobTitle) LIKE '%' + upper(@KeyWord) + '%')
 or (upper(M_Tel) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Ext) LIKE '%' + upper(@KeyWord) + '%') or (upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Office) LIKE '%' + upper(@KeyWord) + '%')
 or (upper(M_Email) LIKE '%' + upper(@KeyWord) + '%') or (upper(comp_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@KeyWord", KeyWord);
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oda.Fill(ds);
        return ds;
    }

    public void addMember()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"insert into Member (
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
        oCmd.Parameters.AddWithValue("@M_Guid", M_Guid);
        oCmd.Parameters.AddWithValue("@M_Account", M_Account);
        oCmd.Parameters.AddWithValue("@M_Pwd", M_Pwd);
        oCmd.Parameters.AddWithValue("@M_Name", M_Name);
        oCmd.Parameters.AddWithValue("@M_JobTitle", M_JobTitle);
        oCmd.Parameters.AddWithValue("@M_Tel", M_Tel);
        oCmd.Parameters.AddWithValue("@M_Ext", M_Ext);
        oCmd.Parameters.AddWithValue("@M_Fax", M_Fax);
        oCmd.Parameters.AddWithValue("@M_Phone", M_Phone);
        oCmd.Parameters.AddWithValue("@M_Email", M_Email);
        oCmd.Parameters.AddWithValue("@M_Addr", M_Addr);
        oCmd.Parameters.AddWithValue("@M_City", M_City);
        oCmd.Parameters.AddWithValue("@M_Office", M_Office);
        oCmd.Parameters.AddWithValue("@M_Competence", M_Competence);
        oCmd.Parameters.AddWithValue("@M_Manager_ID", M_Manager_ID);
        oCmd.Parameters.AddWithValue("@M_CreateId", M_CreateId);
        oCmd.Parameters.AddWithValue("@M_ModId", M_ModId);
        oCmd.Parameters.AddWithValue("@M_Status", "A");

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    public void modMember()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"update Member set
M_Account=@M_Account,
M_Pwd=@M_Pwd,
M_Name=@M_Name,
M_JobTitle=@M_JobTitle,
M_Tel=@M_Tel,
M_Ext=@M_Ext,
M_Fax=@M_Fax,
M_Phone=@M_Phone,
M_Email=@M_Email,
M_Addr=@M_Addr,
M_City=@M_City,
M_Office=@M_Office,
M_Competence=@M_Competence,
M_Manager_ID=@M_Manager_ID,
M_ModDate=@M_ModDate,
M_ModId=@M_ModId
where M_ID=@M_ID
";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@M_ID", M_ID);
        oCmd.Parameters.AddWithValue("@M_Account", M_Account);
        oCmd.Parameters.AddWithValue("@M_Pwd", M_Pwd);
        oCmd.Parameters.AddWithValue("@M_Name", M_Name);
        oCmd.Parameters.AddWithValue("@M_JobTitle", M_JobTitle);
        oCmd.Parameters.AddWithValue("@M_Tel", M_Tel);
        oCmd.Parameters.AddWithValue("@M_Ext", M_Ext);
        oCmd.Parameters.AddWithValue("@M_Fax", M_Fax);
        oCmd.Parameters.AddWithValue("@M_Phone", M_Phone);
        oCmd.Parameters.AddWithValue("@M_Email", M_Email);
        oCmd.Parameters.AddWithValue("@M_Addr", M_Addr);
        oCmd.Parameters.AddWithValue("@M_City", M_City);
        oCmd.Parameters.AddWithValue("@M_Office", M_Office);
        oCmd.Parameters.AddWithValue("@M_Competence", M_Competence);
        oCmd.Parameters.AddWithValue("@M_Manager_ID", M_Manager_ID);
        oCmd.Parameters.AddWithValue("@M_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@M_ModId", M_ModId);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    public void DeleteMember()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"update Member set
M_Status=@M_Status,
M_ModDate=@M_ModDate,
M_ModId=@M_ModId
where M_ID=@M_ID
";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@M_ID", M_ID);
        oCmd.Parameters.AddWithValue("@M_Status", "D");
        oCmd.Parameters.AddWithValue("@M_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@M_ModId", M_ModId);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    public DataTable getMemberById()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select a.*,b.M_Name Manager from Member a 
left join Member b on a.M_Manager_ID=b.M_Guid
where a.M_ID=@M_ID ");
        
        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oCmd.Parameters.AddWithValue("@M_ID", M_ID);
        oda.Fill(ds);
        return ds;
    }

    public DataTable checkEmailRepeat()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select * from Member where M_Email=@M_Email and M_Status<>'D' ");
        if(M_ID!="")
            sb.Append(@"and M_ID<>@M_ID ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oCmd.Parameters.AddWithValue("@M_ID", M_ID);
        oCmd.Parameters.AddWithValue("@M_Email", M_Email);
        oda.Fill(ds);
        return ds;
    }

    public DataTable getMemberByGuid()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
            select * 
            ,(select C_Item_cn from CodeTable where C_Group='02' and C_Item=M_City ) as CityName
            ,(select b.M_Name from Member b where M_Manager_ID=b.M_Guid) as ManagerCname
            from Member
            where M_Guid=@M_Guid 
        ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oCmd.Parameters.AddWithValue("@M_Guid", M_Guid);
        oda.Fill(ds);
        return ds;
    }
    
    public DataSet getManagerList(string pStart, string pEnd)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT COUNT(*) total from Member 
        left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
        left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
where M_Competence='02' and M_Status<>'D' ");

        if (M_City != "")
            sb.Append(@"and M_City=@M_City ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Account) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_JobTitle) LIKE '%' + upper(@KeyWord) + '%')
 or (upper(M_Tel) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Ext) LIKE '%' + upper(@KeyWord) + '%') or (upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Office) LIKE '%' + upper(@KeyWord) + '%')
 or (upper(M_Email) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@"select * from (
           select ROW_NUMBER() over (order by M_CreateDate desc,M_ID desc) itemNo,
        Member.*,city_type.C_Item_cn City,comp_type.C_Item_cn Comp from Member
        left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
        left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
         where M_Competence='02' and M_Status<>'D' ");

        if (M_City != "")
            sb.Append(@"and M_City=@M_City ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Account) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_JobTitle) LIKE '%' + upper(@KeyWord) + '%')
 or (upper(M_Tel) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Ext) LIKE '%' + upper(@KeyWord) + '%') or (upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Office) LIKE '%' + upper(@KeyWord) + '%')
 or (upper(M_Email) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@KeyWord", KeyWord);
        oCmd.Parameters.AddWithValue("@M_City", M_City);
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oda.Fill(ds);
        return ds;
    }

    public DataSet getContractorList(string pStart, string pEnd,string ckNew)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"SELECT COUNT(*) total from Member 
        left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
        left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
        left join ProjectInfo on M_Guid=I_People and I_Status<>'D'
where M_Competence='01' and M_Status<>'D' ");

        if (M_City != "")
            sb.Append(@"and M_City=@M_City ");

        if (ckNew == "new")
            sb.Append(@"and (select COUNT(*) from ProjectInfo where I_Status='A' and I_People=M_Guid)=0 ");
        else
            sb.Append(@"and (select COUNT(*) from ProjectInfo where I_Status='A' and I_People=M_Guid)>0 ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Account) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_JobTitle) LIKE '%' + upper(@KeyWord) + '%')
 or (upper(M_Tel) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Ext) LIKE '%' + upper(@KeyWord) + '%') or (upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Office) LIKE '%' + upper(@KeyWord) + '%')
 or (upper(M_Email) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@"select * from (
           select ROW_NUMBER() over (order by M_CreateDate desc,M_ID desc) itemNo,
        Member.*,city_type.C_Item_cn City,comp_type.C_Item_cn Comp,I_Guid from Member
        left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
        left join CodeTable comp_type on comp_type.C_Group='03' and comp_type.C_Item=M_Competence
        left join ProjectInfo on M_Guid=I_People and I_Status<>'D'
         where M_Competence='01' and M_Status<>'D' ");

        if (M_City != "")
            sb.Append(@"and M_City=@M_City ");

        if (ckNew == "new")
            sb.Append(@"and (select COUNT(*) from ProjectInfo where I_Status='A' and I_People=M_Guid)=0 ");
        else
            sb.Append(@"and (select COUNT(*) from ProjectInfo where I_Status='A' and I_People=M_Guid)>0 ");

        if (KeyWord != "")
        {
            sb.Append(@"and ((upper(M_Account) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Name) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_JobTitle) LIKE '%' + upper(@KeyWord) + '%')
 or (upper(M_Tel) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Ext) LIKE '%' + upper(@KeyWord) + '%') or (upper(city_type.C_Item_cn) LIKE '%' + upper(@KeyWord) + '%') or (upper(M_Office) LIKE '%' + upper(@KeyWord) + '%')
 or (upper(M_Email) LIKE '%' + upper(@KeyWord) + '%')) ");
        }

        sb.Append(@")#tmp where itemNo between @pStart and @pEnd ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();

        oCmd.Parameters.AddWithValue("@KeyWord", KeyWord);
        oCmd.Parameters.AddWithValue("@M_City", M_City);
        oCmd.Parameters.AddWithValue("@pStart", pStart);
        oCmd.Parameters.AddWithValue("@pEnd", pEnd);
        oda.Fill(ds);
        return ds;
    }

    public string getProgectGuidByPersonId()
    {
        string tmpstr = string.Empty;
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select I_Guid from ProjectInfo where I_People=(SELECT M_Guid from Member where M_ID=@M_ID) and I_Status='A' ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oCmd.Parameters.AddWithValue("@M_ID", M_ID);
        oda.Fill(ds);

        if (ds.Rows.Count > 0)
            tmpstr = ds.Rows[0]["I_Guid"].ToString();

        return tmpstr;
    }

    public void modPersonalFile()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"update Member set
M_Account=@M_Account,
M_Pwd=@M_Pwd,
M_Name=@M_Name,
M_JobTitle=@M_JobTitle,
M_Tel=@M_Tel,
M_Ext=@M_Ext,
M_Fax=@M_Fax,
M_Phone=@M_Phone,
M_Email=@M_Email,
M_Addr=@M_Addr,
M_City=@M_City,
M_Office=@M_Office,
M_ModDate=@M_ModDate,
M_ModId=@M_ModId
where M_ID=@M_ID
";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@M_ID", M_ID);
        oCmd.Parameters.AddWithValue("@M_Account", M_Account);
        oCmd.Parameters.AddWithValue("@M_Pwd", M_Pwd);
        oCmd.Parameters.AddWithValue("@M_Name", M_Name);
        oCmd.Parameters.AddWithValue("@M_JobTitle", M_JobTitle);
        oCmd.Parameters.AddWithValue("@M_Tel", M_Tel);
        oCmd.Parameters.AddWithValue("@M_Ext", M_Ext);
        oCmd.Parameters.AddWithValue("@M_Fax", M_Fax);
        oCmd.Parameters.AddWithValue("@M_Phone", M_Phone);
        oCmd.Parameters.AddWithValue("@M_Email", M_Email);
        oCmd.Parameters.AddWithValue("@M_Addr", M_Addr);
        oCmd.Parameters.AddWithValue("@M_City", M_City);
        oCmd.Parameters.AddWithValue("@M_Office", M_Office);
        oCmd.Parameters.AddWithValue("@M_ModDate", DateTime.Now);
        oCmd.Parameters.AddWithValue("@M_ModId", M_ModId);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    public DataTable getProjectDate()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select PD_StartDate,PD_EndDate from ProjectDate
where PD_Type=(SELECT M_City from Member where M_ID=@M_ID) ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oCmd.Parameters.AddWithValue("@M_ID", M_ID);
        oda.Fill(ds);
        return ds;

    }

    public string getGuidById()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select M_Guid from Member where M_ID=@M_ID ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oCmd.Parameters.AddWithValue("@M_ID", M_ID);
        oda.Fill(ds);

        string tmp = string.Empty;
        if (ds.Rows.Count > 0)
            tmp = ds.Rows[0]["M_Guid"].ToString();

        return tmp;
    }

    public void changeConntractor(string NewID,string OrgID)
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        oCmd.CommandText = @"
            update ProjectInfo set I_People=@NewID where I_Status='A' and I_People=@OrgID 
            update Member set M_Status='D' where M_Guid=@OrgID
        ";
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        oCmd.Parameters.AddWithValue("@NewID", NewID);
        oCmd.Parameters.AddWithValue("@OrgID", OrgID);

        oCmd.Connection.Open();
        oCmd.ExecuteNonQuery();
        oCmd.Connection.Close();
    }

    public DataTable getPersonByGuid()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"select Member.*,city_type.C_Item_cn City from Member 
	left join CodeTable city_type on city_type.C_Group='02' and city_type.C_Item=M_City
where M_Guid=@M_Guid ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataTable ds = new DataTable();
        oCmd.Parameters.AddWithValue("@M_Guid", M_Guid);
        oda.Fill(ds);
        return ds;
    }

    public DataSet getPersonInfoById()
    {
        SqlCommand oCmd = new SqlCommand();
        oCmd.Connection = new SqlConnection(ConfigurationManager.AppSettings["ConnectionString"]);
        StringBuilder sb = new StringBuilder();

        sb.Append(@"
select Member.*,city_tyoe.C_Item_cn City from Member 
left join CodeTable city_tyoe on C_Group='02' and C_Item=M_City
where M_ID=@M_ID

select * from Member where M_Guid=(select M_Manager_ID from Member where M_ID=@M_ID) ");

        oCmd.CommandText = sb.ToString();
        oCmd.CommandType = CommandType.Text;
        SqlDataAdapter oda = new SqlDataAdapter(oCmd);
        DataSet ds = new DataSet();
        oCmd.Parameters.AddWithValue("@M_ID", M_ID);
        oda.Fill(ds);
        return ds;
    }
}