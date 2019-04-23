<%@ WebHandler Language="C#" Class="ProjectList" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Data;
using System.Collections.Generic;
using System.Xml;

public class ProjectList : IHttpHandler,IRequiresSessionState
{
    public class peojectTooL
    {
        public string I_ID { get; set; }//
        public string I_Guid { get; set; }//
        public string I_1_Sdate { get; set; }//
        public string I_3_Edate { get; set; }//
        public string I_City { get; set; }//
        public string I_Office { get; set; }//
        public string I_People { get; set; }//
        public string I_Createdate { get; set; }//
        public string I_Modifydate { get; set; }//
        public string cityName { get; set; }//
        public string personNmae { get; set; }//
        public string user { get; set; }//guid
        public string userid { get; set; }//id
        public string I_Flag { get; set; }//
        public string M_ID { get; set; }//
        public string M_competence { get; set; }//
        public string chk_flag { get; set; }//
    }
    public void ProcessRequest (HttpContext context)
    {
        ProjectInfo_DB pj = new ProjectInfo_DB();
        ProjectDate_DB pjd = new ProjectDate_DB();
        if (LogInfo.mGuid != "")
        {
            try
            {
                string str_func = string.IsNullOrEmpty(context.Request.Form["func"]) ? "" : context.Request.Form["func"].ToString().Trim();
                switch (str_func)
                {
                    //撈計畫列表
                    case "load_projectlist":
                        string cityno = LogInfo.city;
                        string mguid = LogInfo.mGuid;
                        string str_keyword = string.IsNullOrEmpty(context.Request.Form["str_keyword"]) ? "" : context.Request.Form["str_keyword"].ToString().Trim();
                        pj._I_City = cityno.Trim();
                        pj._I_People = mguid.Trim();
                        pj._M_Competence = LogInfo.competence;
                        pj._str_keyword = str_keyword;
                        DataTable dt = pj.selectProjectList();
                        if (dt.Rows.Count > 0)
                        {
                            List<peojectTooL> eList = new List<peojectTooL>();
                            for (int i=0;i<dt.Rows.Count;i++) {
                                peojectTooL e = new peojectTooL();
                                e.I_ID = dt.Rows[i]["I_ID"].ToString().Trim();//id
                                e.I_Guid = dt.Rows[i]["I_Guid"].ToString().Trim();//Guid
                                e.I_1_Sdate = dt.Rows[i]["I_1_Sdate"].ToString().Trim();//
                                e.I_3_Edate = dt.Rows[i]["I_3_Edate"].ToString().Trim();//
                                e.I_City = dt.Rows[i]["I_City"].ToString().Trim().Replace("/", "-");//
                                e.I_Office = dt.Rows[i]["I_Office"].ToString().Trim().Replace("/", "-");//
                                e.I_People = dt.Rows[i]["I_People"].ToString().Trim();//
                                e.I_Createdate = Convert.ToDateTime(dt.Rows[i]["I_Createdate"].ToString()).ToString("yyyy/MM/dd");//
                                e.I_Modifydate = Convert.ToDateTime(dt.Rows[i]["I_Modifydate"].ToString()).ToString("yyyy/MM/dd");//
                                e.cityName = dt.Rows[i]["cityName"].ToString().Trim();//
                                e.personNmae = dt.Rows[i]["personNmae"].ToString().Trim();//
                                e.user = LogInfo.mGuid;//
                                e.userid = LogInfo.id;//登入者ID
                                e.I_Flag = dt.Rows[i]["I_Flag"].ToString().Trim();//該計畫是否定稿
                                e.M_ID = dt.Rows[i]["M_ID"].ToString().Trim();//該計畫承辦人ID
                                e.M_competence = LogInfo.competence;//登入者權限
                                e.chk_flag = dt.Rows[i]["chk_flag"].ToString().Trim();//同機關底下是否有人定稿 無:null 有:I_Guid(定稿的計畫GUID)
                                eList.Add(e);
                            }
                            System.Web.Script.Serialization.JavaScriptSerializer objSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                            string ans = objSerializer.Serialize(eList);  //new
                            context.Response.ContentType = "application/json";
                            context.Response.Write(ans);
                        }
                        else {
                            context.Response.Write("nodata," + LogInfo.id + "");
                        }
                        break;

                    //撈計畫by person guid
                    case "load_projectbyperson":
                        pj._I_People = LogInfo.mGuid;
                        DataTable dtpj = pj.getProjectInfo();
                        if (dtpj.Rows.Count > 0)
                        {
                            //有填過資料
                            context.Response.Write("have");
                        }
                        else {
                            //沒填過資料  承辦人
                            if (LogInfo.competence == "01")
                            {
                                string sday = "";
                                string eday = "";
                                //計畫起訖日期
                                pjd._PD_Type = LogInfo.city;
                                DataTable dtDate = pjd.SelectList();
                                if (dtDate.Rows.Count>0 && dtDate.Rows[0]["PD_StartDate"] != null && dtDate.Rows[0]["PD_StartDate"].ToString().Trim()!="" && dtDate.Rows[0]["PD_EndDate"] != null && dtDate.Rows[0]["PD_EndDate"].ToString().Trim()!="") {
                                    sday = Convert.ToDateTime(dtDate.Rows[0]["PD_StartDate"].ToString()).ToString("yyyy/MM/dd");
                                    eday = Convert.ToDateTime(dtDate.Rows[0]["PD_EndDate"].ToString()).ToString("yyyy/MM/dd");
                                    //pj._I_1_Sdate = sday;
                                    //pj._I_3_Edate = eday;
                                    //pj._I_People = LogInfo.mGuid;
                                    //pj._I_City = LogInfo.city;
                                    //pj._I_Office = LogInfo.office;
                                    //pj.addProjectInfo();
                                    context.Response.Write("ok");
                                }else {
                                    context.Response.Write("nodate");
                                }

                            }else
                            {
                                //主管 or 管理者
                            }

                        }
                        break;


                }
            }
            catch (Exception ex)
            {
                context.Response.Write("error");
            }

        }
        else
        {
            context.Response.Write("timeout");
        }
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}