<%@ WebHandler Language="C#" Class="ProjectInfo" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Data;
using System.Collections.Generic;
using System.Xml;

public class ProjectInfo : IHttpHandler,IRequiresSessionState
{
    public class personInfoTooL
    {
        public string M_ID { get; set; }//ID
        public string M_Guid { get; set; }//GUID
        public string M_Name { get; set; }//姓名
        public string M_City { get; set; }//機關代碼
        public string CityName { get; set; }//機關名稱
    }
    public class pjdTooL
    {
        public string PD_StartDate { get; set; }//起始日
        public string PD_EndDate { get; set; }//結束日
        public string I_Flag { get; set; }
        public string I_People { get; set; }
        public string user { get; set; }
        public string M_competence { get; set; }
        public string CityName { get; set; }
        public string I_Office { get; set; }
        public string M_ID { get; set; }
        public string chk_flag { get; set; }
    }
    public class projectInfoTooL
    {
        public string I_ID { get; set; }
        public string I_GUID { get; set; }
        public string I_City { get; set; }
        public string I_Office { get; set; }
        public string I_People { get; set; }
        public string I_1_Sdate { get; set; }
        public string I_1_Edate { get; set; }
        public string I_2_Sdate { get; set; }
        public string I_2_Edate { get; set; }
        public string I_3_Sdate { get; set; }
        public string I_3_Edate { get; set; }
        //設備汰換經費
        public string I_Money_item1_1 { get; set; }
        public string I_Money_item1_2 { get; set; }
        public string I_Money_item1_3 { get; set; }
        public string I_Money_item1_all { get; set; }
        public string I_Money_item2_1 { get; set; }
        public string I_Money_item2_2 { get; set; }
        public string I_Money_item2_3 { get; set; }
        public string I_Money_item2_all { get; set; }
        public string I_Money_item3_1 { get; set; }
        public string I_Money_item3_2 { get; set; }
        public string I_Money_item3_3 { get; set; }
        public string I_Money_item3_all { get; set; }
        //擴大補助經費
        public string I_Money_item4_1 { get; set; }
        public string I_Money_item4_2 { get; set; }
        public string I_Money_item4_3 { get; set; }
        public string I_Money_item4_all { get; set; }
        public string I_Other_Oneself { get; set; }
        public string I_Other_Oneself_Money { get; set; }
        public string I_Other_Other { get; set; }
        public string I_Other_Other_name { get; set; }
        public string I_Other_Other_Money { get; set; }
        public string I_Target { get; set; }
        public string I_Summary { get; set; }
        public string I_Finish_item1_1 { get; set; }
        public string I_Finish_item1_2 { get; set; }
        public string I_Finish_item1_3 { get; set; }
        public string I_Finish_item1_all { get; set; }
        public string I_Finish_item2_1 { get; set; }
        public string I_Finish_item2_2 { get; set; }
        public string I_Finish_item2_3 { get; set; }
        public string I_Finish_item2_all { get; set; }
        public string I_Finish_item3_1 { get; set; }
        public string I_Finish_item3_2 { get; set; }
        public string I_Finish_item3_3 { get; set; }
        public string I_Finish_item3_all { get; set; }
        public string I_Finish_item4_1 { get; set; }
        public string I_Finish_item4_2 { get; set; }
        public string I_Finish_item4_3 { get; set; }
        public string I_Finish_item4_all { get; set; }
        public string I_Finish_item5_1 { get; set; }
        public string I_Finish_item5_2 { get; set; }
        public string I_Finish_item5_3 { get; set; }
        public string I_Finish_item5_all { get; set; }
        public string I_Boss { get; set; }
        public string CityName { get; set; }
        public string Cname { get; set; }
        public string JobTitle { get; set; }
        public string Tel { get; set; }
        public string Phone { get; set; }
        public string Fax { get; set; }
        public string Email { get; set; }
        public string Addr { get; set; }
        public string ManagerCname { get; set; }
        public string ManagerJobTitle { get; set; }
        public string ManagerTel { get; set; }
        public string ManagerPhone { get; set; }
        public string ManagerFax { get; set; }
        public string ManagerCityName { get; set; }
        public string ManagerEmail { get; set; }
        public string ManagerAddr { get; set; }
        public string I_Flag { get; set; }
        public string user { get; set; }
        public string chk_flag { get; set; }
        public string all_dates { get; set; }
        public string all_datee { get; set; }
        public string M_competence { get; set; }
        public string M_ID { get; set; }
        public string M_Name { get; set; }
        public string M_Office { get; set; }
    }
    Member_DB mb = new Member_DB();
    ProjectInfo_DB pj = new ProjectInfo_DB();
    ProjectDate_DB pjd = new ProjectDate_DB();
    CodeTable_DB codeDB = new CodeTable_DB();
    Log_DB l_db = new Log_DB();
    public void ProcessRequest(HttpContext context)
    {
        if (LogInfo.mGuid == "")
        {
            context.Response.Write("timeout");
            return;
        }
        string period = (context.Request["period"] != null) ? context.Request["period"].ToString() : "";
        string type = (context.Request["type"] != null) ? context.Request["type"].ToString() : "";
        string person_id = (context.Request["person_id"] != null) ? context.Request["person_id"].ToString() : "";

        mb._M_ID = person_id;
        string project_id = mb.getProgectGuidByPersonId();

        //check 是否為同縣市承辦人
        if (LogInfo.competence != "SA")
        {
            pj._I_GUID = project_id;
            DataTable ccy = pj.checkCity();
            if (ccy.Rows.Count > 0)
            {
                if (ccy.Rows[0]["I_City"].ToString() != LogInfo.city)
                {
                    context.Response.Write("DiffCity");
                    return;
                }
            }
        }

        string str_func = string.IsNullOrEmpty(context.Request.Form["func"]) ? "" : context.Request.Form["func"].ToString().Trim();
        switch (str_func) {
            case "load_projectinfo"://撈計畫資料
                try
                {
                    //string mid = string.IsNullOrEmpty(context.Request.Form["mid"]) ? "" : context.Request.Form["mid"].ToString().Trim();
                    //mb._M_ID = mid;
                    //string iguid = mb.getProgectGuidByPersonId();
                    string iguid = string.IsNullOrEmpty(context.Request.Form["str_tmpguid"]) ? "" : context.Request.Form["str_tmpguid"].ToString().Trim();
                    DataTable dt = new DataTable();
                    pj._I_GUID = iguid;
                    dt = pj.getProjectInfo();
                    string sday = "";
                    string eday = "";
                    if (dt.Rows.Count > 0)
                    {
                        pjd._PD_Type = dt.Rows[0]["I_City"].ToString().Trim();
                        DataTable dtDate = pjd.SelectList();
                        if (dtDate.Rows.Count > 0 && dtDate.Rows[0]["PD_StartDate"] != null && dtDate.Rows[0]["PD_StartDate"].ToString().Trim() != "" && dtDate.Rows[0]["PD_EndDate"] != null && dtDate.Rows[0]["PD_EndDate"].ToString().Trim() != "")
                        {
                            List<projectInfoTooL> eList = new List<projectInfoTooL>();
                            projectInfoTooL e = new projectInfoTooL();
                            e.I_ID = dt.Rows[0]["I_ID"].ToString().Trim();//
                            e.I_GUID = dt.Rows[0]["I_Guid"].ToString().Trim();//
                            e.I_City = dt.Rows[0]["I_City"].ToString().Trim();//
                            e.I_Office = dt.Rows[0]["I_Office"].ToString().Trim();//
                            e.I_People = dt.Rows[0]["I_People"].ToString().Trim();//
                            e.I_1_Sdate = dt.Rows[0]["I_1_Sdate"].ToString().Trim();//
                            e.I_1_Edate = dt.Rows[0]["I_1_Edate"].ToString().Trim();//
                            e.I_2_Sdate = dt.Rows[0]["I_2_Sdate"].ToString().Trim();//
                            e.I_2_Edate = dt.Rows[0]["I_2_Edate"].ToString().Trim();//
                            e.I_3_Sdate = dt.Rows[0]["I_3_Sdate"].ToString().Trim();//
                            e.I_3_Edate = dt.Rows[0]["I_3_Edate"].ToString().Trim();//
                            e.I_Money_item1_1 = dt.Rows[0]["I_Money_item1_1"].ToString().Trim();//
                            e.I_Money_item1_2 = dt.Rows[0]["I_Money_item1_2"].ToString().Trim();//
                            e.I_Money_item1_3 = dt.Rows[0]["I_Money_item1_3"].ToString().Trim();//
                            e.I_Money_item1_all = dt.Rows[0]["I_Money_item1_all"].ToString().Trim();//
                            e.I_Money_item2_1 = dt.Rows[0]["I_Money_item2_1"].ToString().Trim();//
                            e.I_Money_item2_2 = dt.Rows[0]["I_Money_item2_2"].ToString().Trim();//
                            e.I_Money_item2_3 = dt.Rows[0]["I_Money_item2_3"].ToString().Trim();//
                            e.I_Money_item2_all = dt.Rows[0]["I_Money_item2_all"].ToString().Trim();//
                            e.I_Money_item3_1 = dt.Rows[0]["I_Money_item3_1"].ToString().Trim();//
                            e.I_Money_item3_2 = dt.Rows[0]["I_Money_item3_2"].ToString().Trim();//
                            e.I_Money_item3_3 = dt.Rows[0]["I_Money_item3_3"].ToString().Trim();//
                            e.I_Money_item3_all = dt.Rows[0]["I_Money_item3_all"].ToString().Trim();//
                            e.I_Money_item4_1 = dt.Rows[0]["I_Money_item4_1"].ToString().Trim();//
                            e.I_Money_item4_2 = dt.Rows[0]["I_Money_item4_2"].ToString().Trim();//
                            e.I_Money_item4_3 = dt.Rows[0]["I_Money_item4_3"].ToString().Trim();//
                            e.I_Money_item4_all = dt.Rows[0]["I_Money_item4_all"].ToString().Trim();//
                            e.I_Other_Oneself = dt.Rows[0]["I_Other_Oneself"].ToString().Trim();//
                            e.I_Other_Oneself_Money = dt.Rows[0]["I_Other_Oneself_Money"].ToString().Trim();//
                            e.I_Other_Other = dt.Rows[0]["I_Other_Other"].ToString().Trim();//
                            e.I_Other_Other_name = dt.Rows[0]["I_Other_Other_name"].ToString().Trim();//
                            e.I_Other_Other_Money = dt.Rows[0]["I_Other_Other_Money"].ToString().Trim();//
                            e.I_Target = dt.Rows[0]["I_Target"].ToString().Trim();//
                            e.I_Summary = dt.Rows[0]["I_Summary"].ToString().Trim();//
                            e.I_Finish_item1_1 = dt.Rows[0]["I_Finish_item1_1"].ToString().Trim();//
                            e.I_Finish_item1_2 = dt.Rows[0]["I_Finish_item1_2"].ToString().Trim();//
                            e.I_Finish_item1_3 = dt.Rows[0]["I_Finish_item1_3"].ToString().Trim();//
                            e.I_Finish_item1_all = dt.Rows[0]["I_Finish_item1_all"].ToString().Trim();//
                            e.I_Finish_item2_1 = dt.Rows[0]["I_Finish_item2_1"].ToString().Trim();//
                            e.I_Finish_item2_2 = dt.Rows[0]["I_Finish_item2_2"].ToString().Trim();//
                            e.I_Finish_item2_3 = dt.Rows[0]["I_Finish_item2_3"].ToString().Trim();//
                            e.I_Finish_item2_all = dt.Rows[0]["I_Finish_item2_all"].ToString().Trim();//
                            e.I_Finish_item3_1 = dt.Rows[0]["I_Finish_item3_1"].ToString().Trim();//
                            e.I_Finish_item3_2 = dt.Rows[0]["I_Finish_item3_2"].ToString().Trim();//
                            e.I_Finish_item3_3 = dt.Rows[0]["I_Finish_item3_3"].ToString().Trim();//
                            e.I_Finish_item3_all = dt.Rows[0]["I_Finish_item3_all"].ToString().Trim();//
                            e.I_Finish_item4_1 = dt.Rows[0]["I_Finish_item4_1"].ToString().Trim();//
                            e.I_Finish_item4_2 = dt.Rows[0]["I_Finish_item4_2"].ToString().Trim();//
                            e.I_Finish_item4_3 = dt.Rows[0]["I_Finish_item4_3"].ToString().Trim();//
                            e.I_Finish_item4_all = dt.Rows[0]["I_Finish_item4_all"].ToString().Trim();//
                            e.I_Finish_item5_1 = dt.Rows[0]["I_Finish_item5_1"].ToString().Trim();//
                            e.I_Finish_item5_2 = dt.Rows[0]["I_Finish_item5_2"].ToString().Trim();//
                            e.I_Finish_item5_3 = dt.Rows[0]["I_Finish_item5_3"].ToString().Trim();//
                            e.I_Finish_item5_all = dt.Rows[0]["I_Finish_item5_all"].ToString().Trim();//
                            e.I_Boss = dt.Rows[0]["I_Boss"].ToString().Trim();//
                            e.CityName = dt.Rows[0]["CityName"].ToString().Trim();//
                            e.Cname = dt.Rows[0]["Cname"].ToString().Trim();//
                            e.JobTitle = dt.Rows[0]["JobTitle"].ToString().Trim();//
                            e.Tel = dt.Rows[0]["Tel"].ToString().Trim();//
                            e.Phone = dt.Rows[0]["Phone"].ToString().Trim();//
                            e.Fax = dt.Rows[0]["Fax"].ToString().Trim();//
                            e.Email = dt.Rows[0]["Email"].ToString().Trim();//
                            e.Addr = dt.Rows[0]["Addr"].ToString().Trim();//
                            e.ManagerCname = dt.Rows[0]["ManagerCname"].ToString().Trim();//
                            e.ManagerJobTitle = dt.Rows[0]["ManagerJobTitle"].ToString().Trim();//
                            e.ManagerTel = dt.Rows[0]["ManagerTel"].ToString().Trim();//
                            e.ManagerPhone = dt.Rows[0]["ManagerPhone"].ToString().Trim();//
                            e.ManagerFax = dt.Rows[0]["ManagerFax"].ToString().Trim();//
                            e.ManagerCityName = dt.Rows[0]["ManagerCityName"].ToString().Trim();//
                            e.ManagerEmail = dt.Rows[0]["ManagerEmail"].ToString().Trim();//
                            e.ManagerAddr = dt.Rows[0]["ManagerAddr"].ToString().Trim();//
                            e.I_Flag = dt.Rows[0]["I_Flag"].ToString().Trim();//
                            e.chk_flag = dt.Rows[0]["chk_flag"].ToString().Trim();//該機關底下是否已經有定稿的資料
                            e.user = LogInfo.mGuid;//
                            e.M_ID = dt.Rows[0]["M_ID"].ToString().Trim();//
                            e.M_Name = dt.Rows[0]["M_Name"].ToString().Trim();//
                            e.M_Office = dt.Rows[0]["M_Office"].ToString().Trim();//
                            e.M_competence = LogInfo.competence;//
                            sday = Convert.ToDateTime(dtDate.Rows[0]["PD_StartDate"].ToString()).ToString("yyyy/MM/dd");
                            eday = Convert.ToDateTime(dtDate.Rows[0]["PD_EndDate"].ToString()).ToString("yyyy/MM/dd");
                            e.all_dates = sday;//
                            e.all_datee = eday;//
                            eList.Add(e);
                            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                            string ans = objSerializer.Serialize(eList);  //new
                            context.Response.ContentType = "application/json";
                            context.Response.Write(ans);
                        }
                        else {
                            context.Response.Write("nodate");
                        }

                    }
                    else {
                        //沒填過過資料
                        //計畫起訖日期
                        //如果是承辦人 必須先判斷所屬機關有沒有設定期程
                        if (LogInfo.competence == "01")
                        {
                            string strchkflag = "";
                            pj._I_City = LogInfo.city;
                            DataTable dt_chkflag = pj.getFlagByCity();
                            if (dt_chkflag.Rows.Count > 0)
                            {
                                strchkflag = dt_chkflag.Rows[0]["I_Guid"].ToString().Trim();
                            }
                            pjd._PD_Type = LogInfo.city;
                            DataTable dtDate = pjd.SelectList();
                            mb._M_ID = LogInfo.id;
                            DataTable dtMem = mb.getMemberById();
                            if (dtDate.Rows.Count > 0 && dtDate.Rows[0]["PD_StartDate"] != null && dtDate.Rows[0]["PD_StartDate"].ToString().Trim() != "" && dtDate.Rows[0]["PD_EndDate"] != null && dtDate.Rows[0]["PD_EndDate"].ToString().Trim() != "")
                            {
                                List<pjdTooL> eList = new List<pjdTooL>();
                                pjdTooL e = new pjdTooL();
                                sday = Convert.ToDateTime(dtDate.Rows[0]["PD_StartDate"].ToString()).ToString("yyyy/MM/dd");
                                eday = Convert.ToDateTime(dtDate.Rows[0]["PD_EndDate"].ToString()).ToString("yyyy/MM/dd");
                                e.I_People = LogInfo.mGuid;//
                                e.user = LogInfo.mGuid;//
                                e.I_Office = LogInfo.office;//
                                DataTable ctcn = codeDB.getCn("02", LogInfo.city);
                                e.CityName = ctcn.Rows[0]["C_Item_cn"].ToString().Trim();//
                                e.M_ID = LogInfo.id;
                                e.PD_StartDate = sday;//
                                e.PD_EndDate = eday;//
                                e.chk_flag = strchkflag;
                                e.M_competence = LogInfo.competence;
                                eList.Add(e);
                                System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                                string ans = objSerializer.Serialize(eList);  //new
                                context.Response.ContentType = "application/json";
                                context.Response.Write(ans);
                            }
                            else
                            {
                                context.Response.Write("nodate");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    context.Response.Write("error");
                }
                break;
            //修改基本資料
            case "mod_project":
                try
                {
                    string saveType = string.IsNullOrEmpty(context.Request.Form["saveType"]) ? "" : context.Request.Form["saveType"].ToString().Trim();
                    string str_I_ID = string.IsNullOrEmpty(context.Request.Form["str_I_ID"]) ? "" : context.Request.Form["str_I_ID"].ToString().Trim();
                    string str_I_Guid = string.IsNullOrEmpty(context.Request.Form["str_I_Guid"]) ? "" : context.Request.Form["str_I_Guid"].ToString().Trim();
                    string str_1_sdate = string.IsNullOrEmpty(context.Request.Form["str_1_sdate"]) ? "" : context.Request.Form["str_1_sdate"].ToString().Trim();
                    string str_1_edate = string.IsNullOrEmpty(context.Request.Form["str_1_edate"]) ? "" : context.Request.Form["str_1_edate"].ToString().Trim();
                    string str_2_sdate = string.IsNullOrEmpty(context.Request.Form["str_2_sdate"]) ? "" : context.Request.Form["str_2_sdate"].ToString().Trim();
                    string str_2_edate = string.IsNullOrEmpty(context.Request.Form["str_2_edate"]) ? "" : context.Request.Form["str_2_edate"].ToString().Trim();
                    string str_3_sdate = string.IsNullOrEmpty(context.Request.Form["str_3_sdate"]) ? "" : context.Request.Form["str_3_sdate"].ToString().Trim();
                    string str_3_edate = string.IsNullOrEmpty(context.Request.Form["str_3_edate"]) ? "" : context.Request.Form["str_3_edate"].ToString().Trim();
                    string str_money_item1_1 = string.IsNullOrEmpty(context.Request.Form["str_money_item1_1"]) ? "0" : context.Request.Form["str_money_item1_1"].ToString().Trim();
                    string str_money_item1_2 = string.IsNullOrEmpty(context.Request.Form["str_money_item1_2"]) ? "0" : context.Request.Form["str_money_item1_2"].ToString().Trim();
                    string str_money_item1_3 = string.IsNullOrEmpty(context.Request.Form["str_money_item1_3"]) ? "0" : context.Request.Form["str_money_item1_3"].ToString().Trim();
                    string str_money_item1_sum = string.IsNullOrEmpty(context.Request.Form["str_money_item1_sum"]) ? "0" : context.Request.Form["str_money_item1_sum"].ToString().Trim();
                    string str_money_item2_1 = string.IsNullOrEmpty(context.Request.Form["str_money_item2_1"]) ? "0" : context.Request.Form["str_money_item2_1"].ToString().Trim();
                    string str_money_item2_2 = string.IsNullOrEmpty(context.Request.Form["str_money_item2_2"]) ? "0" : context.Request.Form["str_money_item2_2"].ToString().Trim();
                    string str_money_item2_3 = string.IsNullOrEmpty(context.Request.Form["str_money_item2_3"]) ? "0" : context.Request.Form["str_money_item2_3"].ToString().Trim();
                    string str_money_item2_sum = string.IsNullOrEmpty(context.Request.Form["str_money_item2_sum"]) ? "0" : context.Request.Form["str_money_item2_sum"].ToString().Trim();
                    string str_money_item3_1 = string.IsNullOrEmpty(context.Request.Form["str_money_item3_1"]) ? "0" : context.Request.Form["str_money_item3_1"].ToString().Trim();
                    string str_money_item3_2 = string.IsNullOrEmpty(context.Request.Form["str_money_item3_2"]) ? "0" : context.Request.Form["str_money_item3_2"].ToString().Trim();
                    string str_money_item3_3 = string.IsNullOrEmpty(context.Request.Form["str_money_item3_3"]) ? "0" : context.Request.Form["str_money_item3_3"].ToString().Trim();
                    string str_money_item3_sum = string.IsNullOrEmpty(context.Request.Form["str_money_item3_sum"]) ? "0" : context.Request.Form["str_money_item3_sum"].ToString().Trim();
                    string str_money_item4_1 = string.IsNullOrEmpty(context.Request.Form["str_money_item4_1"]) ? "0" : context.Request.Form["str_money_item4_1"].ToString().Trim();
                    string str_money_item4_2 = string.IsNullOrEmpty(context.Request.Form["str_money_item4_2"]) ? "0" : context.Request.Form["str_money_item4_2"].ToString().Trim();
                    string str_money_item4_3 = string.IsNullOrEmpty(context.Request.Form["str_money_item4_3"]) ? "0" : context.Request.Form["str_money_item4_3"].ToString().Trim();
                    string str_money_item4_sum = string.IsNullOrEmpty(context.Request.Form["str_money_item4_sum"]) ? "0" : context.Request.Form["str_money_item4_sum"].ToString().Trim();
                    
                    string str_Other_Oneself = string.IsNullOrEmpty(context.Request.Form["str_Other_Oneself"]) ? "" : context.Request.Form["str_Other_Oneself"].ToString().Trim();
                    string str_Other_Oneself_Money = string.IsNullOrEmpty(context.Request.Form["str_Other_Oneself_Money"]) ? "0" : context.Request.Form["str_Other_Oneself_Money"].ToString().Trim();
                    string str_Other_Other = string.IsNullOrEmpty(context.Request.Form["str_Other_Other"]) ? "" : context.Request.Form["str_Other_Other"].ToString().Trim();
                    string str_Other_Other_name = string.IsNullOrEmpty(context.Request.Form["str_Other_Other_name"]) ? "" : context.Request.Form["str_Other_Other_name"].ToString().Trim();
                    string str_Other_Other_Money = string.IsNullOrEmpty(context.Request.Form["str_Other_Other_Money"]) ? "0" : context.Request.Form["str_Other_Other_Money"].ToString().Trim();
                    string str_Target = string.IsNullOrEmpty(context.Request.Form["str_Target"]) ? "" : context.Request.Form["str_Target"].ToString().Trim();
                    string str_summary = string.IsNullOrEmpty(context.Request.Form["str_summary"]) ? "" : context.Request.Form["str_summary"].ToString().Trim();
                    string str_Finish_item1_1 = string.IsNullOrEmpty(context.Request.Form["str_Finish_item1_1"]) ? "0" : context.Request.Form["str_Finish_item1_1"].ToString().Trim();
                    string str_Finish_item1_2 = string.IsNullOrEmpty(context.Request.Form["str_Finish_item1_2"]) ? "0" : context.Request.Form["str_Finish_item1_2"].ToString().Trim();
                    string str_Finish_item1_3 = string.IsNullOrEmpty(context.Request.Form["str_Finish_item1_3"]) ? "0" : context.Request.Form["str_Finish_item1_3"].ToString().Trim();
                    string str_Finish_item1_all = string.IsNullOrEmpty(context.Request.Form["str_Finish_item1_all"]) ? "0" : context.Request.Form["str_Finish_item1_all"].ToString().Trim();
                    string str_Finish_item2_1 = string.IsNullOrEmpty(context.Request.Form["str_Finish_item2_1"]) ? "0" : context.Request.Form["str_Finish_item2_1"].ToString().Trim();
                    string str_Finish_item2_2 = string.IsNullOrEmpty(context.Request.Form["str_Finish_item2_2"]) ? "0" : context.Request.Form["str_Finish_item2_2"].ToString().Trim();
                    string str_Finish_item2_3 = string.IsNullOrEmpty(context.Request.Form["str_Finish_item2_3"]) ? "0" : context.Request.Form["str_Finish_item2_3"].ToString().Trim();
                    string str_Finish_item2_all = string.IsNullOrEmpty(context.Request.Form["str_Finish_item2_all"]) ? "0" : context.Request.Form["str_Finish_item2_all"].ToString().Trim();
                    string str_Finish_item3_1 = string.IsNullOrEmpty(context.Request.Form["str_Finish_item3_1"]) ? "0" : context.Request.Form["str_Finish_item3_1"].ToString().Trim();
                    string str_Finish_item3_2 = string.IsNullOrEmpty(context.Request.Form["str_Finish_item3_2"]) ? "0" : context.Request.Form["str_Finish_item3_2"].ToString().Trim();
                    string str_Finish_item3_3 = string.IsNullOrEmpty(context.Request.Form["str_Finish_item3_3"]) ? "0" : context.Request.Form["str_Finish_item3_3"].ToString().Trim();
                    string str_Finish_item3_all = string.IsNullOrEmpty(context.Request.Form["str_Finish_item3_all"]) ? "0" : context.Request.Form["str_Finish_item3_all"].ToString().Trim();
                    string str_Finish_item4_1 = string.IsNullOrEmpty(context.Request.Form["str_Finish_item4_1"]) ? "0" : context.Request.Form["str_Finish_item4_1"].ToString().Trim();
                    string str_Finish_item4_2 = string.IsNullOrEmpty(context.Request.Form["str_Finish_item4_2"]) ? "0" : context.Request.Form["str_Finish_item4_2"].ToString().Trim();
                    string str_Finish_item4_3 = string.IsNullOrEmpty(context.Request.Form["str_Finish_item4_3"]) ? "0" : context.Request.Form["str_Finish_item4_3"].ToString().Trim();
                    string str_Finish_item4_all = string.IsNullOrEmpty(context.Request.Form["str_Finish_item4_all"]) ? "0" : context.Request.Form["str_Finish_item4_all"].ToString().Trim();
                    string str_Finish_item5_1 = string.IsNullOrEmpty(context.Request.Form["str_Finish_item5_1"]) ? "0" : context.Request.Form["str_Finish_item5_1"].ToString().Trim();
                    string str_Finish_item5_2 = string.IsNullOrEmpty(context.Request.Form["str_Finish_item5_2"]) ? "0" : context.Request.Form["str_Finish_item5_2"].ToString().Trim();
                    string str_Finish_item5_3 = string.IsNullOrEmpty(context.Request.Form["str_Finish_item5_3"]) ? "0" : context.Request.Form["str_Finish_item5_3"].ToString().Trim();
                    string str_Finish_item5_all = string.IsNullOrEmpty(context.Request.Form["str_Finish_item5_all"]) ? "0" : context.Request.Form["str_Finish_item5_all"].ToString().Trim();
                    string str_M_ID = string.IsNullOrEmpty(context.Request.Form["str_M_ID"]) ? "" : context.Request.Form["str_M_ID"].ToString().Trim();
                    string str_savetype = string.IsNullOrEmpty(context.Request.Form["str_savetype"]) ? "" : context.Request.Form["str_savetype"].ToString().Trim();
                    //管理者挑選的承辦人ID
                    //string str_selPeopleID = string.IsNullOrEmpty(context.Request.Form["str_selPeople"]) ? "" : context.Request.Form["str_selPeople"].ToString().Trim();
                    string str_tmpguid = string.IsNullOrEmpty(context.Request.Form["str_tmpguid"]) ? "" : context.Request.Form["str_tmpguid"].ToString().Trim();

                    pj._I_GUID = str_tmpguid;
                    DataTable dt_chkpj = pj.getProjectInfo();
                    //pj._M_ID = str_selPeopleID;//只有管理者修改才有
                    pj._I_ID = str_I_ID;
                    pj._I_GUID = str_tmpguid;
                    pj._I_1_Sdate = str_1_sdate;
                    pj._I_1_Edate = str_1_edate;
                    pj._I_2_Sdate = str_2_sdate;
                    pj._I_2_Edate = str_2_edate;
                    pj._I_3_Sdate = str_3_sdate;
                    pj._I_3_Edate = str_3_edate;
                    pj._I_Money_item1_1 =str_money_item1_1;
                    pj._I_Money_item1_2 = str_money_item1_2;
                    pj._I_Money_item1_3 = str_money_item1_3;
                    pj._I_Money_item1_all = str_money_item1_sum;
                    pj._I_Money_item2_1 = str_money_item2_1;
                    pj._I_Money_item2_2 = str_money_item2_2;
                    pj._I_Money_item2_3 = str_money_item2_3;
                    pj._I_Money_item2_all = str_money_item2_sum;
                    pj._I_Money_item3_1 = str_money_item3_1;
                    pj._I_Money_item3_2 = str_money_item3_2;
                    pj._I_Money_item3_3 = str_money_item3_3;
                    pj._I_Money_item3_all = str_money_item3_sum;
                    pj._I_Money_item4_1 = str_money_item4_1;
                    pj._I_Money_item4_2 = str_money_item4_2;
                    pj._I_Money_item4_3 = str_money_item4_3;
                    pj._I_Money_item4_all = str_money_item4_sum;

                    pj._I_Other_Oneself = str_Other_Oneself;
                    pj._I_Other_Oneself_Money = Convert.ToDecimal(str_Other_Oneself_Money);
                    pj._I_Other_Other = str_Other_Other;
                    pj._I_Other_Other_name = str_Other_Other_name;
                    pj._I_Other_Other_Money = Convert.ToDecimal(str_Other_Other_Money);
                    pj._I_Target = str_Target;
                    pj._I_Summary = str_summary;
                    pj._I_Finish_item1_1 = Convert.ToDecimal(str_Finish_item1_1);
                    pj._I_Finish_item1_2 = Convert.ToDecimal(str_Finish_item1_2);
                    pj._I_Finish_item1_3 = Convert.ToDecimal(str_Finish_item1_3);
                    pj._I_Finish_item1_all = Convert.ToDecimal(str_Finish_item1_all);
                    pj._I_Finish_item2_1 = Convert.ToDecimal(str_Finish_item2_1);
                    pj._I_Finish_item2_2 = Convert.ToDecimal(str_Finish_item2_2);
                    pj._I_Finish_item2_3 = Convert.ToDecimal(str_Finish_item2_3);
                    pj._I_Finish_item2_all = Convert.ToDecimal(str_Finish_item2_all);
                    pj._I_Finish_item3_1 = Convert.ToDecimal(str_Finish_item3_1);
                    pj._I_Finish_item3_2 = Convert.ToDecimal(str_Finish_item3_2);
                    pj._I_Finish_item3_3 = Convert.ToDecimal(str_Finish_item3_3);
                    pj._I_Finish_item3_all = Convert.ToDecimal(str_Finish_item3_all);
                    pj._I_Finish_item4_1 = Convert.ToDecimal(str_Finish_item4_1);
                    pj._I_Finish_item4_2 = Convert.ToDecimal(str_Finish_item4_2);
                    pj._I_Finish_item4_3 = Convert.ToDecimal(str_Finish_item4_3);
                    pj._I_Finish_item4_all = Convert.ToDecimal(str_Finish_item4_all);
                    pj._I_Finish_item5_1 = Convert.ToDecimal(str_Finish_item5_1);
                    pj._I_Finish_item5_2 = Convert.ToDecimal(str_Finish_item5_2);
                    pj._I_Finish_item5_3 = Convert.ToDecimal(str_Finish_item5_3);
                    pj._I_Finish_item5_all = Convert.ToDecimal(str_Finish_item5_all);
                    pj._I_ModId = LogInfo.mGuid;
                    //20171201決定把管理者代填拿掉
                    pj._I_People = LogInfo.mGuid;
                    if (dt_chkpj.Rows.Count>0)
                    {
                        //修改
                        pj.modProjectInfo();

                    }
                    else {
                        //新增
                        pj.addProjectInfo();
                    }
                    //自動存檔不進LOG
                    if (str_savetype != "auto")
                    {
                        l_db._L_Type = "07";
                        l_db._L_Person = LogInfo.mGuid;
                        l_db._L_IP = Common.GetIP4Address();
                        l_db._L_ModItemGuid = str_tmpguid;
                        l_db._L_Desc = "基本資料";
                        l_db.addLog();
                    }

                    context.Response.Write("OK");
                }
                catch (Exception ex) {
                    context.Response.Write("Error：" + ex.Message.Replace("'", "\""));
                }
                break;
            //定稿
            case "submit_project":
                pj._I_People = LogInfo.mGuid;
                pj._I_City = LogInfo.city;
                DataTable dtchkpj = pj.selectCheckProject();
                if (dtchkpj.Rows.Count > 0)
                {//已經有定稿資料
                    context.Response.Write("have");
                }
                else {
                    pj.submitProject();
                    context.Response.Write("OK");
                }
                break;

            //複製計畫
            case "copy_projectdata":
                string copy_fromid = string.IsNullOrEmpty(context.Request.Form["copy_fromid"]) ? "" : context.Request.Form["copy_fromid"].ToString().Trim();
                string return_mid = "";
                pj._I_People = LogInfo.mGuid;
                pj._M_Competence = LogInfo.competence;//權限
                DataTable dt_chk_pj = pj.getProjectInfo();
                if (dt_chk_pj.Rows.Count>0) {//表示有填過資料 用UPDATE
                    pj._I_ID = copy_fromid;
                    pj.copyProjectByIDUpdate();
                }else {//沒填過資料 用insert
                    pj._I_ID = copy_fromid;
                    pj.copyProjectByIDInsert();
                    if (LogInfo.competence == "SA")
                    {
                        DataTable dtiid = pj.selCopyIIDForAdmin();
                        if (dtiid.Rows.Count > 0)
                        {
                            return_mid = dtiid.Rows[0]["I_ID"].ToString().Trim();
                        }
                    }
                    else {
                        return_mid = LogInfo.id;
                    }
                }
                context.Response.Write("OK,"+LogInfo.competence+","+return_mid+"");
                break;
            //撈期程資料
            case "load_date":
                string str_p_city = string.IsNullOrEmpty(context.Request.Form["p_city"]) ? "" : context.Request.Form["p_city"].ToString().Trim();
                pjd._PD_Type = str_p_city;
                DataTable dtm = pjd.SelectList();
                if (dtm.Rows.Count > 0)
                {
                    List<pjdTooL> eList = new List<pjdTooL>();
                    pjdTooL e = new pjdTooL();
                    string sday = "";
                    string eday = "";
                    if (dtm.Rows.Count > 0 && dtm.Rows[0]["PD_StartDate"] != null && dtm.Rows[0]["PD_StartDate"].ToString().Trim() != "" && dtm.Rows[0]["PD_EndDate"] != null && dtm.Rows[0]["PD_EndDate"].ToString().Trim() != "")
                    {
                        sday = Convert.ToDateTime(dtm.Rows[0]["PD_StartDate"].ToString()).ToString("yyyy/MM/dd");
                        eday = Convert.ToDateTime(dtm.Rows[0]["PD_EndDate"].ToString()).ToString("yyyy/MM/dd");
                        e.PD_StartDate = sday;//
                        e.PD_EndDate = eday;//
                        eList.Add(e);
                        System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                        string ans = objSerializer.Serialize(eList);  //new
                        context.Response.ContentType = "application/json";
                        context.Response.Write(ans);
                    } else {
                        context.Response.Write("nodate");
                    }
                }
                else {
                    context.Response.Write("nodate");
                }
                break;
        }


    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}