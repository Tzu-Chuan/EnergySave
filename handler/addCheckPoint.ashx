<%@ WebHandler Language="C#" Class="addCheckPoint" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Xml;
using System.Collections.Generic;

public class addCheckPoint : IHttpHandler,IRequiresSessionState {
    CheckPoint_DB cp_db = new CheckPoint_DB();
    Member_DB m_db = new Member_DB();
    Log_DB l_db = new Log_DB();
    public void ProcessRequest(HttpContext context)
    {
        try
        {
            if (LogInfo.mGuid == "")
            {
                context.Response.Write("reLogin");
                return;
            }

            string person_id = (context.Request["person_id"] != null) ? context.Request["person_id"].ToString() : "";
            string period = (context.Request["period"] != null) ? context.Request["period"].ToString() : "";
            string category = (context.Request["category"] != null) ? context.Request["category"].ToString() : "";
            string parentid = (context.Request["parentid"] != null) ? context.Request["parentid"].ToString() : "";

            //PushItem
            string[] del_pguid = null;
            string[] pgid = null;
            string[] pushitem = null;
            string[] exfinish = null;

            //CheckPoint
            string[] del_cpguid = null;
            string[] cpgid = null;
            string[] cp_no = null;
            string[] cp_year = null;
            string[] cp_month = null;
            //string[] cp_desc = null;
            string[] lastitem = null;

            switch (category)
            {
                //節電基礎工作
                case "01":
                    del_pguid = (context.Request["del_pguid"] != null) ? context.Request["del_pguid"].ToString().Split(',') : null;
                    pgid = (context.Request["pgid"] != null) ? context.Request["pgid"].ToString().Split(',') : null;
                    pushitem = (context.Request["pushitem"] != null) ? context.Request["pushitem"].ToString().Split(',') : null;

                    del_cpguid = (context.Request["del_cpguid"] != null) ? context.Request["del_cpguid"].ToString().Split(',') : null;
                    cpgid = (context.Request["cpgid"] != null) ? context.Request["cpgid"].ToString().Split(',') : null;
                    cp_no = (context.Request["cp_no"] != null) ? context.Request["cp_no"].ToString().Split(',') : null;
                    cp_year = (context.Request["cp_year"] != null) ? context.Request["cp_year"].ToString().Split(',') : null;
                    cp_month = (context.Request["cp_month"] != null) ? context.Request["cp_month"].ToString().Split(',') : null;
                    //cp_desc = (context.Request["cp_desc"] != null) ? context.Request["cp_desc"].ToString().Split(',') : null;
                    lastitem = (context.Request["lastitem"] != null) ? context.Request["lastitem"].ToString().Split(',') : null;
                    break;
                //因地制宜
                case "02":
                    del_pguid = (context.Request["p_del_pguid"] != null) ? context.Request["p_del_pguid"].ToString().Split(',') : null;
                    pgid = (context.Request["p_pgid"] != null) ? context.Request["p_pgid"].ToString().Split(',') : null;
                    pushitem = (context.Request["p_pushitem"] != null) ? context.Request["p_pushitem"].ToString().Split(',') : null;

                    del_cpguid = (context.Request["p_del_cpguid"] != null) ? context.Request["p_del_cpguid"].ToString().Split(',') : null;
                    cpgid = (context.Request["p_cpgid"] != null) ? context.Request["p_cpgid"].ToString().Split(',') : null;
                    cp_no = (context.Request["p_cp_no"] != null) ? context.Request["p_cp_no"].ToString().Split(',') : null;
                    cp_year = (context.Request["p_cp_year"] != null) ? context.Request["p_cp_year"].ToString().Split(',') : null;
                    cp_month = (context.Request["p_cp_month"] != null) ? context.Request["p_cp_month"].ToString().Split(',') : null;
                    //cp_desc = (context.Request["p_cp_desc"] != null) ? context.Request["p_cp_desc"].ToString().Split(',') : null;
                    lastitem = (context.Request["p_lastitem"] != null) ? context.Request["p_lastitem"].ToString().Split(',') : null;
                    break;
                //設備汰換與智慧用電
                case "03":
                    del_pguid = (context.Request["s_del_pguid"] != null) ? context.Request["s_del_pguid"].ToString().Split(',') : null;
                    pgid = (context.Request["s_pgid"] != null) ? context.Request["s_pgid"].ToString().Split(',') : null;
                    pushitem = (context.Request["s_pushitem"] != null) ? context.Request["s_pushitem"].ToString().Split(',') : null;

                    del_cpguid = (context.Request["s_del_cpguid"] != null) ? context.Request["s_del_cpguid"].ToString().Split(',') : null;
                    cpgid = (context.Request["s_cpgid"] != null) ? context.Request["s_cpgid"].ToString().Split(',') : null;
                    cp_no = (context.Request["s_cp_no"] != null) ? context.Request["s_cp_no"].ToString().Split(',') : null;
                    cp_year = (context.Request["s_cp_year"] != null) ? context.Request["s_cp_year"].ToString().Split(',') : null;
                    cp_month = (context.Request["s_cp_month"] != null) ? context.Request["s_cp_month"].ToString().Split(',') : null;
                    //cp_desc = (context.Request["s_cp_desc"] != null) ? context.Request["s_cp_desc"].ToString().Split(',') : null;
                    lastitem = (context.Request["s_lastitem"] != null) ? context.Request["s_lastitem"].ToString().Split(',') : null;
                    break;
                //擴大補助
                case "04":
                    del_pguid = (context.Request["a_del_pguid"] != null) ? context.Request["a_del_pguid"].ToString().Split(',') : null;
                    pgid = (context.Request["a_pgid"] != null) ? context.Request["a_pgid"].ToString().Split(',') : null;
                    pushitem = (context.Request["a_pushitem"] != null) ? context.Request["a_pushitem"].ToString().Split(',') : null;
                    exfinish = (context.Request["exFinish"] != null) ? context.Request["exFinish"].ToString().Split(',') : null;

                    del_cpguid = (context.Request["a_del_cpguid"] != null) ? context.Request["a_del_cpguid"].ToString().Split(',') : null;
                    cpgid = (context.Request["a_cpgid"] != null) ? context.Request["a_cpgid"].ToString().Split(',') : null;
                    cp_no = (context.Request["a_cp_no"] != null) ? context.Request["a_cp_no"].ToString().Split(',') : null;
                    cp_year = (context.Request["a_cp_year"] != null) ? context.Request["a_cp_year"].ToString().Split(',') : null;
                    cp_month = (context.Request["a_cp_month"] != null) ? context.Request["a_cp_month"].ToString().Split(',') : null;
                    //cp_desc = (context.Request["a_cp_desc"] != null) ? context.Request["a_cp_desc"].ToString().Split(',') : null;
                    lastitem = (context.Request["a_lastitem"] != null) ? context.Request["a_lastitem"].ToString().Split(',') : null;
                    break;
            }

            //刪除 PushItem
            foreach (string delstr in del_pguid)
            {
                cp_db._P_Guid = delstr;
                cp_db._P_ModId = LogInfo.mGuid;
                cp_db.deletePushItem();
            }

            //刪除 CheckPoint
            foreach (string delstr in del_cpguid)
            {
                cp_db._CP_Guid = delstr;
                cp_db._CP_ModId = LogInfo.mGuid;
                cp_db.deleteCheckPoint();
            }


            //先解出XML放到Array再一併處理
            List<string> noAry = new List<string>();
            List<string> descAry = new List<string>();
            string tmpXML = (context.Request["tmpXML"] != null) ? context.Server.UrlDecode(context.Request["tmpXML"]) : "<?xml version='1.0' encoding='utf-8'?><root></root>";
            XmlDocument tmpXDoc = new XmlDocument();
            tmpXDoc.LoadXml(tmpXML);
            XmlNodeList xNode = tmpXDoc.SelectNodes("/root/cpitem");
            for (int i = 0; i < xNode.Count; i++)
            {
                noAry.Add(xNode[i].SelectSingleNode("no").InnerText);
                descAry.Add(xNode[i].SelectSingleNode("desc").InnerText);
            }

            m_db._M_ID = person_id;
            parentid = m_db.getProgectGuidByPersonId();

            int BreakNo = 0;
            for (int pnum = 0; pnum < pushitem.Length; pnum++)
            {
                if (pushitem[pnum].Trim() == "")
                    continue;

                string piGuid = (pgid[pnum].Trim() == "") ? Guid.NewGuid().ToString("N") : pgid[pnum].Trim();
                cp_db._P_Guid = piGuid;
                cp_db._P_ParentId = parentid;
                cp_db._P_Period = period;
                cp_db._P_Type = category;
                cp_db._P_ItemName = pushitem[pnum].Trim();
                if (category == "04")
                    cp_db._P_ExFinish = exfinish[pnum];
                cp_db._P_CreateId = LogInfo.mGuid;
                cp_db._P_ModId = LogInfo.mGuid;
                cp_db._P_Sort = (pnum + 1);
                cp_db.addPushItem();

                for (int i = BreakNo; i < noAry.Count; i++)
                {
                    cp_db._CP_Guid = (cpgid[i].Trim() == "") ? Guid.NewGuid().ToString("N") : cpgid[i].Trim();
                    cp_db._CP_ParentId = piGuid;
                    cp_db._CP_ProjectId = parentid;
                    cp_db._CP_Point = noAry[i].Trim();
                    cp_db._CP_ReserveYear = cp_year[i].Trim();
                    cp_db._CP_ReserveMonth = cp_month[i].Trim();
                    cp_db._CP_Desc = descAry[i].Trim();
                    cp_db._CP_CreateId = LogInfo.mGuid;
                    cp_db._CP_ModId = LogInfo.mGuid;
                    cp_db.addCheckPoint();

                    if (i == 0)
                    {
                        //Log
                        l_db._L_Type = "07";
                        l_db._L_Person = LogInfo.mGuid;
                        l_db._L_IP = Common.GetIP4Address();
                        l_db._L_ModItemGuid = parentid;
                        l_db._L_Desc = "查核點";
                        l_db.addLog();
                    }

                    BreakNo += 1;
                    //該筆若為PushItem最後一項則跳離迴圈
                    if (lastitem[i] == "Y")
                        break;
                }
            }

            context.Response.Write("<script type='text/JavaScript'>parent.feedback('succeed','" + category + "');</script>");
        }
        catch (Exception ex)
        {
            context.Response.Write("<script type='text/JavaScript'>parent.feedback('Error：" + ex.Message.Replace("'", "\"") + "')</script>");
        }
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}