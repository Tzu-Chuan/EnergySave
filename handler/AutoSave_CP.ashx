<%@ WebHandler Language="C#" Class="AutoSave_CP" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Xml;
using System.Collections.Generic;

public class AutoSave_CP : IHttpHandler,IRequiresSessionState {
    CheckPoint_DB cp_db = new CheckPoint_DB();
    Member_DB m_db = new Member_DB();
    public void ProcessRequest(HttpContext context) {
        try
        {
            string LoginPerson = "SessionTimeout-AutoSave";
            if (LogInfo.mGuid != "")
                LoginPerson = LogInfo.mGuid;

            string person_id = (context.Request["person_id"] != null) ? context.Request["person_id"].ToString() : "";
            string period = (context.Request["period"] != null) ? context.Request["period"].ToString() : "";
            string parentid = (context.Request["parentid"] != null) ? context.Request["parentid"].ToString() : "";
            int BreakNo = 0;

            //get Project ID
            m_db._M_ID = person_id;
            parentid = m_db.getProgectGuidByPersonId();

            //節電基礎工作
            string[] bw_del_pguid = (context.Request["del_pguid"] != null) ? context.Request["del_pguid"].ToString().Split(',') : null;
            string[] bw_pgid = (context.Request["pgid"] != null) ? context.Request["pgid"].ToString().Split(',') : null;
            string[] bw_pushitem = (context.Request["pushitem"] != null) ? context.Request["pushitem"].ToString().Split(',') : null;

            string[] bw_del_cpguid = (context.Request["del_cpguid"] != null) ? context.Request["del_cpguid"].ToString().Split(',') : null;
            string[] bw_cpgid = (context.Request["cpgid"] != null) ? context.Request["cpgid"].ToString().Split(',') : null;
            string[] bw_cp_no = (context.Request["cp_no"] != null) ? context.Request["cp_no"].ToString().Split(',') : null;
            string[] bw_cp_year = (context.Request["cp_year"] != null) ? context.Request["cp_year"].ToString().Split(',') : null;
            string[] bw_cp_month = (context.Request["cp_month"] != null) ? context.Request["cp_month"].ToString().Split(',') : null;
            //string[] bw_cp_desc = (context.Request["cp_desc"] != null) ? context.Request["cp_desc"].ToString().Split(',') : null;
            string[] bw_lastitem = (context.Request["lastitem"] != null) ? context.Request["lastitem"].ToString().Split(',') : null;

            //刪除 PushItem
            foreach (string delstr in bw_del_pguid)
            {
                cp_db._P_Guid = delstr;
                cp_db._P_ModId = LoginPerson;
                cp_db.deletePushItem();
            }

            //刪除 CheckPoint
            foreach (string delstr in bw_del_cpguid)
            {
                cp_db._CP_Guid = delstr;
                cp_db._CP_ModId = LoginPerson;
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

            for (int pnum = 0; pnum < bw_pushitem.Length; pnum++)
            {
                if (bw_pushitem[pnum].Trim() == "")
                    continue;

                string piGuid = (bw_pgid[pnum].Trim() == "") ? Guid.NewGuid().ToString("N") : bw_pgid[pnum].Trim();
                cp_db._P_Guid = piGuid;
                cp_db._P_ParentId = parentid;
                cp_db._P_Period = period;
                cp_db._P_Type = "01";
                cp_db._P_ItemName = bw_pushitem[pnum].Trim();
                cp_db._P_CreateId = LoginPerson;
                cp_db._P_ModId = LoginPerson;
                cp_db.addPushItem();

                for (int i = BreakNo; i < noAry.Count; i++)
                {
                    cp_db._CP_Guid = (bw_cpgid[i].Trim() == "") ? Guid.NewGuid().ToString("N") : bw_cpgid[i].Trim();
                    cp_db._CP_ParentId = piGuid;
                    cp_db._CP_ProjectId = parentid;
                    cp_db._CP_Point = noAry[i].Trim();
                    cp_db._CP_ReserveYear = bw_cp_year[i].Trim();
                    cp_db._CP_ReserveMonth = bw_cp_month[i].Trim();
                    cp_db._CP_Desc = descAry[i].Trim();
                    cp_db._CP_CreateId = LogInfo.mGuid;
                    cp_db._CP_ModId = LogInfo.mGuid;
                    cp_db.addCheckPoint();

                    BreakNo += 1;
                    //該筆若為PushItem最後一項則跳離迴圈
                    if (bw_lastitem[i] == "Y")
                        break;
                }
            }

            //因地制宜
            string[] p_del_pguid = (context.Request["p_del_pguid"] != null) ? context.Request["p_del_pguid"].ToString().Split(',') : null;
            string[] p_pgid = (context.Request["p_pgid"] != null) ? context.Request["p_pgid"].ToString().Split(',') : null;
            string[] p_pushitem = (context.Request["p_pushitem"] != null) ? context.Request["p_pushitem"].ToString().Split(',') : null;

            string[] p_del_cpguid = (context.Request["p_del_cpguid"] != null) ? context.Request["p_del_cpguid"].ToString().Split(',') : null;
            string[] p_cpgid = (context.Request["p_cpgid"] != null) ? context.Request["p_cpgid"].ToString().Split(',') : null;
            string[] p_cp_no = (context.Request["p_cp_no"] != null) ? context.Request["p_cp_no"].ToString().Split(',') : null;
            string[] p_cp_year = (context.Request["p_cp_year"] != null) ? context.Request["p_cp_year"].ToString().Split(',') : null;
            string[] p_cp_month = (context.Request["p_cp_month"] != null) ? context.Request["p_cp_month"].ToString().Split(',') : null;
            string[] p_cp_desc = (context.Request["p_cp_desc"] != null) ? context.Request["p_cp_desc"].ToString().Split(',') : null;
            string[] p_lastitem = (context.Request["p_lastitem"] != null) ? context.Request["p_lastitem"].ToString().Split(',') : null;

            //刪除 PushItem
            foreach (string delstr in p_del_pguid)
            {
                cp_db._P_Guid = delstr;
                cp_db._P_ModId = LoginPerson;
                cp_db.deletePushItem();
            }

            //刪除 CheckPoint
            foreach (string delstr in p_del_cpguid)
            {
                cp_db._CP_Guid = delstr;
                cp_db._CP_ModId = LoginPerson;
                cp_db.deleteCheckPoint();
            }

            BreakNo = 0;
            for (int pnum = 0; pnum < p_pushitem.Length; pnum++)
            {
                if (p_pushitem[pnum].Trim() == "")
                    continue;

                string piGuid = (p_pgid[pnum].Trim() == "") ? Guid.NewGuid().ToString("N") : p_pgid[pnum].Trim();
                cp_db._P_Guid = piGuid;
                cp_db._P_ParentId = parentid;
                cp_db._P_Period = period;
                cp_db._P_Type = "02";
                cp_db._P_ItemName = p_pushitem[pnum].Trim();
                cp_db._P_CreateId = LogInfo.mGuid;
                cp_db._P_ModId = LogInfo.mGuid;
                cp_db.addPushItem();

                for (int i = BreakNo; i < p_cp_no.Length; i++)
                {
                    cp_db._CP_Guid = (p_cpgid[i].Trim() == "") ? Guid.NewGuid().ToString("N") : p_cpgid[i].Trim();
                    cp_db._CP_ParentId = piGuid;
                    cp_db._CP_ProjectId = parentid;
                    cp_db._CP_Point = p_cp_no[i].Trim();
                    cp_db._CP_ReserveYear = p_cp_year[i].Trim();
                    cp_db._CP_ReserveMonth = p_cp_month[i].Trim();
                    cp_db._CP_Desc = p_cp_desc[i].Trim();
                    cp_db._CP_CreateId = LogInfo.mGuid;
                    cp_db._CP_ModId = LogInfo.mGuid;
                    cp_db.addCheckPoint();

                    BreakNo += 1;
                    //該筆若為PushItem最後一項則跳離迴圈
                    if (p_lastitem[i] == "Y")
                        break;
                }
            }


            //設備汰換與智慧用電
            string[] s_del_pguid = (context.Request["s_del_pguid"] != null) ? context.Request["s_del_pguid"].ToString().Split(',') : null;
            string[] s_pgid = (context.Request["s_pgid"] != null) ? context.Request["s_pgid"].ToString().Split(',') : null;
            string[] s_pushitem = (context.Request["s_pushitem"] != null) ? context.Request["s_pushitem"].ToString().Split(',') : null;

            string[] s_del_cpguid = (context.Request["s_del_cpguid"] != null) ? context.Request["s_del_cpguid"].ToString().Split(',') : null;
            string[] s_cpgid = (context.Request["s_cpgid"] != null) ? context.Request["s_cpgid"].ToString().Split(',') : null;
            string[] s_cp_no = (context.Request["s_cp_no"] != null) ? context.Request["s_cp_no"].ToString().Split(',') : null;
            string[] s_cp_year = (context.Request["s_cp_year"] != null) ? context.Request["s_cp_year"].ToString().Split(',') : null;
            string[] s_cp_month = (context.Request["s_cp_month"] != null) ? context.Request["s_cp_month"].ToString().Split(',') : null;
            string[] s_cp_desc = (context.Request["s_cp_desc"] != null) ? context.Request["s_cp_desc"].ToString().Split(',') : null;
            string[] s_lastitem = (context.Request["s_lastitem"] != null) ? context.Request["s_lastitem"].ToString().Split(',') : null;

            //刪除 PushItem
            foreach (string delstr in s_del_pguid)
            {
                cp_db._P_Guid = delstr;
                cp_db._P_ModId = LoginPerson;
                cp_db.deletePushItem();
            }

            //刪除 CheckPoint
            foreach (string delstr in s_del_cpguid)
            {
                cp_db._CP_Guid = delstr;
                cp_db._CP_ModId = LoginPerson;
                cp_db.deleteCheckPoint();
            }

            BreakNo = 0;
            for (int pnum = 0; pnum < s_pushitem.Length; pnum++)
            {
                if (s_pushitem[pnum].Trim() == "")
                    continue;

                string piGuid = (s_pgid[pnum].Trim() == "") ? Guid.NewGuid().ToString("N") : s_pgid[pnum].Trim();
                cp_db._P_Guid = piGuid;
                cp_db._P_ParentId = parentid;
                cp_db._P_Period = period;
                cp_db._P_Type = "03";
                cp_db._P_ItemName = s_pushitem[pnum].Trim();
                cp_db._P_CreateId = LogInfo.mGuid;
                cp_db._P_ModId = LogInfo.mGuid;
                cp_db.addPushItem();

                for (int i = BreakNo; i < s_cp_no.Length; i++)
                {
                    cp_db._CP_Guid = (s_cpgid[i].Trim() == "") ? Guid.NewGuid().ToString("N") : s_cpgid[i].Trim();
                    cp_db._CP_ParentId = piGuid;
                    cp_db._CP_ProjectId = parentid;
                    cp_db._CP_Point = s_cp_no[i].Trim();
                    cp_db._CP_ReserveYear = s_cp_year[i].Trim();
                    cp_db._CP_ReserveMonth = s_cp_month[i].Trim();
                    cp_db._CP_Desc = s_cp_desc[i].Trim();
                    cp_db._CP_CreateId = LogInfo.mGuid;
                    cp_db._CP_ModId = LogInfo.mGuid;
                    cp_db.addCheckPoint();

                    BreakNo += 1;
                    //該筆若為PushItem最後一項則跳離迴圈
                    if (s_lastitem[i] == "Y")
                        break;
                }
            }

            context.Response.Write("<script type='text/JavaScript'>parent.autofeedback('succeed');</script>");
        }
        catch (Exception ex) { context.Response.Write("<script type='text/JavaScript'>parent.autofeedback('Error：" + ex.Message.Replace("'", "\"") + "')</script>"); }
    }

    public bool IsReusable {
        get {
            return false;
        }
    }

}