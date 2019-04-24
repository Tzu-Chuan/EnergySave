<%@ WebHandler Language="C#" Class="addSeason" %>

using System;
using System.Web;
using System.Web.SessionState;
using System.Data;
using System.Xml;
using System.Collections.Generic;

public class addSeason : IHttpHandler,IRequiresSessionState {
    ReportSeasonV2_DB rs_db = new ReportSeasonV2_DB();
    PushItemDesc_DB pd_db = new PushItemDesc_DB();
    ReportCheck_DB rc_db = new ReportCheck_DB();
    CheckPoint_DB cp_db = new CheckPoint_DB();
    Member_DB m_db = new Member_DB();
    Log_DB l_db = new Log_DB();
    CodeTable_DB ct_db = new CodeTable_DB();
    ExpandFinish_DB ef_db = new ExpandFinish_DB();
    public void ProcessRequest (HttpContext context) {
        ///-----------------------------------------------------
        ///功    能: 季報詳細資料
        ///說明:
        /// * Request["subbtn"]: 是否送審
        /// * Request["tmpguid"]: 季報Guid
        /// * Request["tmpMonth"]: 月份總計
        /// * Request["year"]: 年
        /// * Request["season"]: 季
        /// * Request["stage"]: 期
        /// * Request["RS_CostDesc"]: 預算狀態
        /// * Request["piGuid"]: 推動項目Guid
        /// 實支數
        /// * Request["RS_Type01Real"]: 節電基礎工作
        /// * Request["RS_Type02Real"]: 因地制宜
        /// * Request["RS_Type03Real"]: 設備汰換及智慧用電
        /// * Request["RS_Type04Real"]: 擴大補助
        /// 推動項目
        /// * Request["exGuid"]: Guid
        /// * Request["exRealFinish"]: 推動項目累計完成數
        /// 查核點
        /// * Request["cpGuid"]: Guid
        /// * Request["CP_RealProcess"]: 查核點累計實際進度
        /// * Request["PD_Summary"]: 辦理情形
        /// * Request["PD_BackwardDesc"]: 進度差異說明
        /// 設備汰換及智慧用電-累計完成數
        /// * Request["RS_03Type01C"]: 無風管冷氣
        /// * Request["RS_03Type02C"]: 老舊辦公室照明
        /// * Request["RS_03Type03C"]: 室內停車場智慧照明
        /// * Request["RS_03Type04C"]: 中型能管系統
        /// * Request["RS_03Type05C"]: 大型能管系統
        ///-----------------------------------------------------
        XmlDocument xDoc = new XmlDocument();

        if (LogInfo.mGuid == "")
        {
            xDoc = ExceptionUtil.GetErrorMassageDocument("請重新登入");
            context.Response.Write("<script  type='text/javascript'>parent.feedback('" + xDoc.OuterXml + "'); parent.location.href='../webpage/login.aspx';</script>");
            return;
        }

        try
        {
            string ReviewStatus = (context.Request["subbtn"] != null) ? context.Request["subbtn"].ToString().Trim() : "";
            string rs_guid = (context.Request["tmpguid"] != null) ? context.Request["tmpguid"].ToString().Trim() : "";
            string TotalMonth = (context.Request["tmpMonth"] != null) ? context.Request["tmpMonth"].ToString().Trim() : "";
            string year = (context.Request["year"] != null) ? context.Request["year"].ToString().Trim() : "";
            string season = (context.Request["season"] != null) ? context.Request["season"].ToString().Trim() : "";
            string stage = (context.Request["stage"] != null) ? context.Request["stage"].ToString().Trim() : "";
            string RS_CostDesc = (context.Request["RS_CostDesc"] != null) ? context.Request["RS_CostDesc"].ToString().Trim() : "";
            string[] cpGuid = (context.Request["cpGuid"] != null) ? context.Request["cpGuid"].ToString().Trim().Split(',') : null;
            string[] CP_RealProcess = (context.Request["CP_RealProcess"] != null) ? context.Request["CP_RealProcess"].ToString().Trim().Split(',') : null;
            string RS_Type01Real = (context.Request["RS_Type01Real"] != null) ? context.Request["RS_Type01Real"].ToString().Trim() : "";
            string RS_Type02Real = (context.Request["RS_Type02Real"] != null) ? context.Request["RS_Type02Real"].ToString().Trim() : "";
            string RS_Type03Real = (context.Request["RS_Type03Real"] != null) ? context.Request["RS_Type03Real"].ToString().Trim() : "";
            string RS_Type04Real = (context.Request["RS_Type04Real"] != null) ? context.Request["RS_Type04Real"].ToString().Trim() : "";
            string RS_03Type01C = (context.Request["RS_03Type01C"] != null) ? context.Request["RS_03Type01C"].ToString().Trim() : "";
            string RS_03Type02C = (context.Request["RS_03Type02C"] != null) ? context.Request["RS_03Type02C"].ToString().Trim() : "";
            string RS_03Type03C = (context.Request["RS_03Type03C"] != null) ? context.Request["RS_03Type03C"].ToString().Trim() : "";
            string RS_03Type04C = (context.Request["RS_03Type04C"] != null) ? context.Request["RS_03Type04C"].ToString().Trim() : "";
            string RS_03Type05C = (context.Request["RS_03Type05C"] != null) ? context.Request["RS_03Type05C"].ToString().Trim() : "";
            /// 推動項目預計完成數
            string[] exGuid = (context.Request["exGuid"] != null) ? context.Request["exGuid"].ToString().Trim().Split(',') : null;
            string[] exRealFinish = (context.Request["exRealFinish"] != null) ? context.Request["exRealFinish"].ToString().Trim().Split(',') : null;
            /// 辦理情形&進度差異說明
            string[] piGuid = (context.Request["pi_guid"] != null) ? context.Request["pi_guid"].ToString().Trim().Split(',') : null;
            string[] PD_Summary = (context.Request["PD_Summary"] != null) ? context.Request["PD_Summary"].ToString().Trim().Split(',') : null;
            string[] PD_BackwardDesc = (context.Request["PD_BackwardDesc"] != null) ? context.Request["PD_BackwardDesc"].ToString().Trim().Split(',') : null;

            m_db._M_ID = LogInfo.id;
            string ProjectGuid = m_db.getProgectGuidByPersonId();

            //辦理情形&進度差異說明
            if (piGuid != null)
            {
                for (int i = 0; i < piGuid.Length; i++)
                {
                    pd_db._PD_Guid = Guid.NewGuid().ToString("N");
                    pd_db._PD_PushitemGuid = piGuid[i];
                    pd_db._PD_ProjectGuid = ProjectGuid;
                    pd_db._PD_Stage = stage;
                    pd_db._PD_Year = year;
                    pd_db._PD_Season = season;
                    pd_db._PD_Summary = PD_Summary[i];
                    pd_db._PD_BackwardDesc = PD_BackwardDesc[i];
                    pd_db.setPushitemDesc();
                }
            }

            //先解出XML放到Array再一併處理
            //List<string> summaryAry = new List<string>();
            //List<string> backwardAry = new List<string>();
            //string tmpXML = (context.Request["tmpXML"] != null) ? context.Server.UrlDecode(context.Request["tmpXML"]) : "<?xml version='1.0' encoding='utf-8'?><root></root>";
            //XmlDocument tmpXDoc = new XmlDocument();
            //tmpXDoc.LoadXml(tmpXML);
            //XmlNodeList xNode = tmpXDoc.SelectNodes("/root/cpitem");
            //for (int i = 0; i < xNode.Count; i++)
            //{
            //    summaryAry.Add(xNode[i].SelectSingleNode("summary").InnerText);
            //    backwardAry.Add(xNode[i].SelectSingleNode("backward").InnerText);
            //}

            //擴大補助累計完成數
            if (exGuid != null)
            {
                for (int i = 0; i < exGuid.Length; i++)
                {
                    ef_db._EF_ReportId = rs_guid;
                    ef_db._EF_PushitemId = exGuid[i];
                    ef_db._EF_Finish = exRealFinish[i];
                    ef_db.SaveExFinish();
                }
            }


            //update 查核點 (累計實際進度)
            if (cpGuid != null)
            {
                for (int i = 0; i < cpGuid.Length; i++)
                {
                    cp_db._CP_Guid = cpGuid[i];
                    cp_db._CP_RealProcess = CP_RealProcess[i];
                    //cp_db._CP_Summary = summaryAry[i];
                    //cp_db._CP_BackwardDesc = backwardAry[i];
                    cp_db.updateSeasonInfo();
                }
            }

            ///基本資料
            DataSet ds = rs_db.getSeasonDetail(LogInfo.mGuid, year, season, stage);
            ///預定進度&實際進度
            DataSet ds2 = rs_db.getSeasonProcess(LogInfo.mGuid, year, season, stage);
            /// 辦理情形&差異進度說明
            pd_db._PD_Year = year;
            pd_db._PD_Season= season;
            pd_db._PD_Stage = stage;
            DataTable pd_dt = pd_db.GetPiDesc(LogInfo.mGuid);

            rs_db._RS_Guid = rs_guid;
            rs_db._RS_PorjectGuid = ProjectGuid;
            rs_db._RS_Year = year;
            rs_db._RS_Season = season;
            rs_db._RS_Stage = stage;
            rs_db._RS_StartDay = toROC_Date(ds.Tables[0].Rows[0]["I_" + stage + "_Sdate"].ToString().Trim());
            rs_db._RS_EndDay = toROC_Date(ds.Tables[0].Rows[0]["I_" + stage + "_Edate"].ToString().Trim());
            rs_db._RS_TotalMonth = CountMonth(ds.Tables[0].Rows[0]["I_" + stage + "_Sdate"].ToString(), ds.Tables[0].Rows[0]["I_" + stage + "_Edate"].ToString().Trim());
            rs_db._RS_CostDesc = RS_CostDesc;
            rs_db._RS_Type01Money = ds.Tables[0].Rows[0]["I_Money_item1_" + stage].ToString().Trim();
            rs_db._RS_Type01Real = (RS_Type01Real != "") ? RS_Type01Real : "0";
            rs_db._RS_Type01RealRate = CountMoneyRealRatio(ds.Tables[0].Rows[0]["I_Money_item1_" + stage].ToString().Trim(), RS_Type01Real);
            rs_db._RS_Type02Money = ds.Tables[0].Rows[0]["I_Money_item2_" + stage].ToString().Trim();
            rs_db._RS_Type02Real = (RS_Type02Real != "") ? RS_Type02Real : "0";
            rs_db._RS_Type02RealRate = CountMoneyRealRatio(ds.Tables[0].Rows[0]["I_Money_item2_" + stage].ToString().Trim(), RS_Type02Real);
            rs_db._RS_Type03Money = ds.Tables[0].Rows[0]["I_Money_item3_" + stage].ToString().Trim();
            rs_db._RS_Type03Real = (RS_Type03Real != "") ? RS_Type03Real : "0";
            rs_db._RS_Type03RealRate = CountMoneyRealRatio(ds.Tables[0].Rows[0]["I_Money_item3_" + stage].ToString().Trim(), RS_Type03Real);
            rs_db._RS_Type04Money = ds.Tables[0].Rows[0]["I_Money_item4_" + stage].ToString().Trim();
            rs_db._RS_Type04Real = (RS_Type04Real != "") ? RS_Type04Real : "0";
            rs_db._RS_Type04RealRate = CountMoneyRealRatio(ds.Tables[0].Rows[0]["I_Money_item4_" + stage].ToString().Trim(), RS_Type04Real);
            rs_db._RS_AllSchedule = Math.Round(double.Parse(ds2.Tables[0].Rows[0]["sumValAll"].ToString().Trim()), 2).ToString();
            rs_db._RS_AllRealSchedule = Math.Round(double.Parse(ds2.Tables[0].Rows[0]["sumRealValAll"].ToString().Trim()), 2).ToString();
            rs_db._RS_01Schedule = Math.Round(double.Parse(ds2.Tables[0].Rows[0]["sumVal01"].ToString().Trim()), 2).ToString();
            rs_db._RS_01RealSchedule = Math.Round(double.Parse(ds2.Tables[0].Rows[0]["sumRealVal01"].ToString().Trim()), 2).ToString();
            rs_db._RS_02Schedule = Math.Round(double.Parse(ds2.Tables[0].Rows[0]["sumVal02"].ToString().Trim()), 2).ToString();
            rs_db._RS_02RealSchedule = Math.Round(double.Parse(ds2.Tables[0].Rows[0]["sumRealVal02"].ToString().Trim()), 2).ToString();
            rs_db._RS_03Schedule = Math.Round(double.Parse(ds2.Tables[0].Rows[0]["sumVal03"].ToString().Trim()), 2).ToString();
            rs_db._RS_03RealSchedule = Math.Round(double.Parse(ds2.Tables[0].Rows[0]["sumRealVal03"].ToString().Trim()), 2).ToString();
            rs_db._RS_04Schedule = Math.Round(double.Parse(ds2.Tables[0].Rows[0]["sumVal04"].ToString().Trim()), 2).ToString();
            rs_db._RS_04RealSchedule = Math.Round(double.Parse(ds2.Tables[0].Rows[0]["sumRealVal04"].ToString().Trim()), 2).ToString();
            rs_db._RS_CheckPointData = CheckPointXML(ds.Tables[2]);
            rs_db._RS_PushItemDesc = PD_XML(pd_dt);
            rs_db._RS_03Type01S = ds.Tables[0].Rows[0]["I_Finish_item1_" + stage].ToString().Trim();
            rs_db._RS_03Type01C = (RS_03Type01C != "") ? RS_03Type01C : "0";
            rs_db._RS_03Type02S = ds.Tables[0].Rows[0]["I_Finish_item2_" + stage].ToString().Trim();
            rs_db._RS_03Type02C = (RS_03Type02C != "") ? RS_03Type02C : "0";
            rs_db._RS_03Type03S = ds.Tables[0].Rows[0]["I_Finish_item3_" + stage].ToString().Trim();
            rs_db._RS_03Type03C = (RS_03Type03C != "") ? RS_03Type03C : "0";
            rs_db._RS_03Type04S = ds.Tables[0].Rows[0]["I_Finish_item4_" + stage].ToString().Trim();
            rs_db._RS_03Type04C = (RS_03Type04C != "") ? RS_03Type04C : "0";
            rs_db._RS_03Type05S = ds.Tables[0].Rows[0]["I_Finish_item5_" + stage].ToString().Trim();
            rs_db._RS_03Type05C = (RS_03Type05C != "") ? RS_03Type05C : "0";
            rs_db._RS_CreateId = LogInfo.mGuid;
            rs_db._RS_ModId = LogInfo.mGuid;

            string logstr = string.Empty;
            DataTable checkDt = rs_db.getSeasonInfo(LogInfo.mGuid, year, season, stage);
            if (checkDt.Rows.Count > 0)
            {
                rs_db.modSeason();
                logstr = "修改季報 - " + year + "年第" + stage + "期第" + season + "季";
            }
            else
            {
                rs_db.addSeason();
                logstr = "建立季報 - " + year + "年第" + stage + "期第" + season + "季";
            }


            /// 季報異動 Log
            l_db._L_Type = "09";
            l_db._L_Person = LogInfo.mGuid;
            l_db._L_IP = Common.GetIP4Address();
            l_db._L_ModItemGuid = rs_guid;
            l_db._L_Desc = logstr;
            l_db.addLog();

            /// 送審
            if (ReviewStatus == "Y")
            {
                rc_db._RC_Guid = Guid.NewGuid().ToString("N").ToLower();
                rc_db._RC_ReportGuid = rs_guid;
                rc_db._RC_PeopleGuid = LogInfo.mGuid;
                rc_db._RC_Stage = stage;
                rc_db._RC_Year = (Int32.Parse(year) + 1911).ToString();
                rc_db._RC_Season = season;
                rc_db._RC_CreateId = LogInfo.mGuid;
                //防止連點按鈕
                DataTable chkDt = rc_db.CheckReportExist();
                if (Int32.Parse(chkDt.Rows[0]["Total"].ToString()) == 0)
                    rc_db.addSeason();

                /// 季報送審 Log
                l_db._L_Type = "05";
                l_db._L_Person = LogInfo.mGuid;
                l_db._L_IP = Common.GetIP4Address();
                l_db._L_ModItemGuid = rs_guid;
                l_db._L_Desc = "季報 - " + year + "年第" + stage + "期第" + season + "季";
                l_db.addLog();

                xDoc.LoadXml("<?xml version='1.0' encoding='utf-8'?><root><Response>SubReview</Response></root>");
            }
            else
            {
                xDoc.LoadXml("<?xml version='1.0' encoding='utf-8'?><root><Response>Save</Response></root>");
            }
        }
        catch (Exception ex)
        {
            xDoc = ExceptionUtil.GetExceptionDocument(ex);
        }

        context.Response.Write("<script  type='text/javascript'>parent.feedback('" + xDoc.OuterXml + "');</script>");
    }


    private string CheckPointXML(DataTable dt)
    {
        string rVal = string.Empty;
        DataView dv = dt.DefaultView;
        if (dv.Count > 0)
        {
            XmlDocument doc = new XmlDocument();
            /// 根節點
            XmlElement cpList = doc.CreateElement("cpList");
            doc.AppendChild(cpList);
            XmlElement PushItem = doc.DocumentElement;
            for (int i = 0; i < dv.Count; i++)
            {
                if (i == 0)
                {
                    /// Node - PushItem 
                    PushItem = doc.CreateElement("PushItem");
                    PushItem.SetAttribute("P_Guid", dv[i]["P_Guid"].ToString());
                    PushItem.SetAttribute("P_Type", dv[i]["P_Type"].ToString());
                    PushItem.SetAttribute("P_Period", dv[i]["P_Period"].ToString());
                    string CodeStr = string.Empty;
                    switch (dv[i]["P_Type"].ToString())
                    {
                        case "03":
                            CodeStr = getDeviceReplaceCode(dv[i]["P_ItemName"].ToString());
                            break;
                        case "04":
                            CodeStr = getExType_Cn(dv[i]["P_ItemName"].ToString());
                            break;
                        default:
                            CodeStr = dv[i]["P_ItemName"].ToString();
                            break;
                    }
                    PushItem.SetAttribute("P_ItemName", CodeStr);
                    PushItem.SetAttribute("P_WorkRatio", dv[i]["P_WorkRatio"].ToString());
                    PushItem.SetAttribute("P_ExFinish", dv[i]["P_ExFinish"].ToString());
                    cpList.AppendChild(PushItem);
                }
                else if (dv[i - 1]["P_Guid"].ToString() != dv[i]["P_Guid"].ToString())
                {
                    /// Node - PushItem 
                    PushItem = doc.CreateElement("PushItem");
                    PushItem.SetAttribute("P_Guid", dv[i]["P_Guid"].ToString());
                    PushItem.SetAttribute("P_Type", dv[i]["P_Type"].ToString());
                    PushItem.SetAttribute("P_Period", dv[i]["P_Period"].ToString());
                    string CodeStr = string.Empty;
                    switch (dv[i]["P_Type"].ToString())
                    {
                        default:
                            CodeStr = dv[i]["P_ItemName"].ToString();
                            break;
                        case "03":
                            CodeStr = getDeviceReplaceCode(dv[i]["P_ItemName"].ToString());
                            break;
                        case "04":
                            CodeStr = getExType_Cn(dv[i]["P_ItemName"].ToString());
                            break;
                    }
                    PushItem.SetAttribute("P_ItemName", CodeStr);
                    PushItem.SetAttribute("P_WorkRatio", dv[i]["P_WorkRatio"].ToString());
                    PushItem.SetAttribute("P_ExFinish", dv[i]["P_ExFinish"].ToString());
                    cpList.AppendChild(PushItem);
                }

                /// Node - CheckPoint 
                XmlElement cp = doc.CreateElement("CheckPoint");
                /// CheckPoint各欄位
                XmlElement gid = doc.CreateElement("CP_Guid");
                gid.InnerText = VerificationString(dv[i]["CP_Guid"].ToString());
                cp.AppendChild(gid);
                XmlElement point = doc.CreateElement("CP_Point");
                point.InnerText = VerificationString(dv[i]["CP_Point"].ToString());
                cp.AppendChild(point);
                XmlElement year = doc.CreateElement("CP_ReserveYear");
                year.InnerText = VerificationString(dv[i]["CP_ReserveYear"].ToString());
                cp.AppendChild(year);
                XmlElement month = doc.CreateElement("CP_ReserveMonth");
                month.InnerText = VerificationString(dv[i]["CP_ReserveMonth"].ToString());
                cp.AppendChild(month);
                XmlElement desc = doc.CreateElement("CP_Desc");
                desc.InnerText = VerificationString(dv[i]["CP_Desc"].ToString());
                cp.AppendChild(desc);
                XmlElement process = doc.CreateElement("CP_Process");
                process.InnerText = VerificationString(dv[i]["CP_Process"].ToString());
                cp.AppendChild(process);
                XmlElement realprocess = doc.CreateElement("CP_RealProcess");
                realprocess.InnerText = VerificationString(dv[i]["CP_RealProcess"].ToString());
                cp.AppendChild(realprocess);
                //XmlElement sumary = doc.CreateElement("CP_Summary");
                //sumary.InnerText = VerificationString(dv[i]["CP_Summary"].ToString());
                //cp.AppendChild(sumary);
                //XmlElement bdesc = doc.CreateElement("CP_BackwardDesc");
                //bdesc.InnerText = VerificationString(dv[i]["CP_BackwardDesc"].ToString());
                //cp.AppendChild(bdesc);
                PushItem.AppendChild(cp);

            }
            rVal = doc.OuterXml.ToString();
        }
        return rVal;
    }

    private string PD_XML(DataTable dt)
    {
        string rVal = string.Empty;
        DataView dv = dt.DefaultView;
        if (dv.Count > 0)
        {
            XmlDocument doc = new XmlDocument();
            /// 根節點
            XmlElement pdList = doc.CreateElement("pdList");
            doc.AppendChild(pdList);
            XmlElement pdItem = doc.DocumentElement;
            for (int i = 0; i < dv.Count; i++)
            {
                pdItem = doc.CreateElement("pd_item");
                pdItem.SetAttribute("PD_ID", dv[i]["PD_ID"].ToString());
                pdItem.SetAttribute("PD_Guid", dv[i]["PD_Guid"].ToString());
                pdItem.SetAttribute("PD_PushitemGuid", dv[i]["PD_PushitemGuid"].ToString());
                pdItem.SetAttribute("PD_ProjectGuid", dv[i]["PD_ProjectGuid"].ToString());
                pdItem.SetAttribute("PD_Stage", dv[i]["PD_Stage"].ToString());
                pdItem.SetAttribute("PD_Year", dv[i]["PD_Year"].ToString());
                pdItem.SetAttribute("PD_Season", dv[i]["PD_Season"].ToString());
                pdItem.SetAttribute("PD_Summary", dv[i]["PD_Summary"].ToString());
                pdItem.SetAttribute("PD_BackwardDesc", dv[i]["PD_BackwardDesc"].ToString());
                pdItem.SetAttribute("PD_CreateDate", DateTime.Parse(dv[i]["PD_CreateDate"].ToString()).ToString("yyyy/MM/dd"));
                pdItem.SetAttribute("PD_ModDate", DateTime.Parse(dv[i]["PD_ModDate"].ToString()).ToString("yyyy/MM/dd"));
                pdItem.SetAttribute("PD_Status", dv[i]["PD_Status"].ToString());
                pdList.AppendChild(pdItem);
            }
            rVal = doc.OuterXml.ToString();
        }
        return rVal;
    }

    private string VerificationString(string str)
    {
        string rVal = string.Empty;
        if (string.IsNullOrEmpty(str))
            rVal = " ";
        else
            rVal = str;
        return rVal;
    }

    private string getDeviceReplaceCode(string item)
    {
        string rVal = string.Empty;
        CheckPoint_DB cp_db = new CheckPoint_DB();
        DataTable dt = cp_db.getPushItemName(item);
        if (dt.Rows.Count > 0)
            rVal = dt.Rows[0]["C_Item_cn"].ToString();
        return rVal;
    }

    private string getExType_Cn(string item)
    {
        string str = string.Empty;
        DataTable dt = ct_db.getCn("09", item);
        if (dt.Rows.Count > 0)
            str = dt.Rows[0]["C_Item_cn"].ToString();
        return str;
    }

    //日期轉民國年
    private string toROC_Date(string dateStr)
    {
        string TaiwanDate = string.Empty;
        var ROC_Calendar = new System.Globalization.TaiwanCalendar();
        TaiwanDate = ROC_Calendar.GetYear(DateTime.Parse(dateStr)).ToString().PadLeft(3, '0') + "/" +
                DateTime.Parse(dateStr).Month.ToString().PadLeft(2, '0') + "/" +
                DateTime.Parse(dateStr).Day.ToString().PadLeft(2, '0');
        return TaiwanDate;
    }

    //計算月份總計
    private string CountMonth(string StartDay,string EndDay)
    {
        string total = string.Empty;
        var ROC_Calendar = new System.Globalization.TaiwanCalendar();
        int startYear = ROC_Calendar.GetYear(DateTime.Parse(StartDay));
        int endYear = ROC_Calendar.GetYear(DateTime.Parse(EndDay));
        int countYear = endYear - startYear;
        int countMonth = (DateTime.Parse(EndDay).Month - DateTime.Parse(StartDay).Month) + 1;
        total = ((countYear * 12) + countMonth).ToString();
        return total;
    }

    //計算月份總計
    private string CountMoneyRealRatio(string Money, string RealMoney)
    {
        string ratio = string.Empty;
        string tmpReal = (RealMoney == "") ? "0" : RealMoney;
        string tmpMoney = (Money == "") ? "0" : Money;
        double m = double.Parse(tmpMoney);
        double rm = double.Parse(tmpReal);
        ratio = Math.Round(((rm / m) * 100), 0).ToString();
        return ratio;
    }


    public bool IsReusable {
        get {
            return false;
        }
    }

}