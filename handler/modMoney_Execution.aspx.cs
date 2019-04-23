using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class handler_modMoney_Execution : System.Web.UI.Page
{
    MoneyExecute_DB me_db = new MoneyExecute_DB();
    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            if (LogInfo.mGuid == "" || LogInfo.city == "")
            {
                Response.Write("reLogin");
                return;
            }
            //PR_CaseName 標案名稱 PR_CaseMoney 標案：發包金額(B) PR_SelfMoney 自辦：金額(C)
            //PR_Stage 期數 PR_City 執行機關  PR_PlanTitle 計畫名稱 PR_Office 主責局處 PR_Method 處理方式 PR_Money 金額 PR_Steps 涉及措施
            string PR_ID = (Request["PR_ID"] != null) ? Request["PR_ID"].ToString().Trim() : "";
            string PR_Stage = (Request["PR_Stage"] != null) ? Request["PR_Stage"].ToString().Trim() : "";
            string PR_City = LogInfo.city;
            string PR_PlanTitle = (Request["PR_PlanTitle"] != null) ? Request["PR_PlanTitle"].ToString().Trim() : "";
            string PR_Office = (Request["PR_Office"] != null) ? Request["PR_Office"].ToString().Trim() : "";
            string PR_Money = (Request["PR_Money"] != null) ? Request["PR_Money"].ToString().Trim() : "";
            string PR_CaseName = (Request["PR_CaseName"] != null) ? Request["PR_CaseName"].ToString().Trim() : "";
            string PR_CaseMoney = (Request["PR_CaseMoney"] != null) ? Request["PR_CaseMoney"].ToString().Trim() : "";
            string PR_SelfMoney = (Request["PR_SelfMoney"] != null) ? Request["PR_SelfMoney"].ToString().Trim() : "";
            string PR_Steps = (Request["PR_Steps"] != null) ? Request["PR_Steps"].ToString().Trim() : "";
            string save_flag = (Request["save_flag"] != null) ? Request["save_flag"].ToString().Trim() : "";

            me_db._PR_ID = PR_ID;
            me_db._PR_Stage = PR_Stage;
            me_db._PR_City = PR_City;
            me_db._PR_PlanTitle = PR_PlanTitle;
            me_db._PR_Office = PR_Office;
            me_db._PR_Money = PR_Money;
            me_db._PR_CaseName = PR_CaseName;
            me_db._PR_CaseMoney = PR_CaseMoney;
            me_db._PR_SelfMoney = PR_SelfMoney;
            me_db._PR_Steps = PR_Steps;
            me_db._PR_CreateId = LogInfo.mGuid;
            me_db._PR_ModId = LogInfo.mGuid;
            me_db._PR_ModDate = DateTime.Now;

            if (save_flag=="new") {
                me_db.addMoney();
            }
            if (save_flag == "mod") {
                me_db.modMoney();
            }
            Response.Write("success");
        }
        catch (Exception ex)
        {
            Response.Write("Error:" + ex.Message);
        }
    }
}