<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="total_grandtotal.aspx.cs" Inherits="WebPage_total_grandtotal" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            //匯出按鈕
            $("#btn_export").click(function () {
                if (funcChk($("#ddlReportType").val(), $("#ddlStage").val(), $("#sel_Sdate").val(), $("#sel_Edate").val())) {
                    if ($("#ddlReportType").val() == "01") {
                        window.location = "../handler/ExportGrandTotal.aspx?s=" + encodeURIComponent($("#ddlStage").val()) + "&sd=" + encodeURIComponent($("#sel_Sdate").val()) + "&ed=" + encodeURIComponent($("#sel_Edate").val()) + "";
                    }
                    if ($("#ddlReportType").val() == "02") {
                        window.location = "../handler/ExportGrandTotalEx.aspx?s=" + encodeURIComponent($("#ddlStage").val()) + "&sd=" + encodeURIComponent($("#sel_Sdate").val()) + "&ed=" + encodeURIComponent($("#sel_Edate").val()) + "";
                    }
                    
                }
                
            });

            //'檢查查詢欄位
            function funcChk(strType, strStage, strSdate, strEdate) {
                var rt = true;
                var strErr = "";
                if (strType == "") {
                    rt = false;
                    strErr += "請選擇[月報類別]\n";
                }
                if (strStage == "") {
                    rt = false;
                    strErr += "請選擇[期別]\n";
                }
                if (strSdate == "") {
                    rt = false;
                    strErr += "請輸入[查詢年月-開始年月]\n";
                }
                if (strEdate == "") {
                    rt = false;
                    strErr += "請輸入[查詢年月-結束年月]\n";
                }
                if (strSdate != "" && strEdate != "") {
                    var reg = /^[1-9]\d*$/;
                    if (reg.test(strSdate) && reg.test(strEdate) && strSdate.length == 6 && strEdate.length == 6) {
                        var intSdate = parseInt(strSdate);
                        var intEdate = parseInt(strEdate);
                        if (intSdate > intEdate) {
                            rt = false;
                            strErr += "查詢月份[開始年月]不能大於[結束年月]\n";
                        }
                    } else {
                        rt = false;
                        strErr += "請輸入正確的[查詢年月]起訖 ex:201901 \n";
                    }
                }
                if (strErr!="") {
                    alert(strErr);
                }
                return rt;
            }
        });

        
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">月報統計表</span></div>
            <div class="right">附加圖表 / 管理員總覽表 / 月報統計表</div>
        </div>

        月報類別：
        <select id="ddlReportType" class="inputex" style="margin-top: 10px;">
            <option value="01">設備汰換</option>
            <option value="02">擴大補助</option>
        </select>
        期別：
        <select id="ddlStage" class="inputex" style="margin-top: 10px;">
            <option value="1" selected="selected">第一期</option>
            <option value="2">第二期</option>
            <option value="3">第三期</option>
        </select>&nbsp;&nbsp;
        查詢年月：
        <input type="text" class="width15 inputex" id="sel_Sdate" placeholder="ex:201901" value="" maxlength="6" />
        ~ 
        <input type="text" class="width15 inputex" id="sel_Edate" placeholder="ex:201902" value="" maxlength="6" />
        &nbsp;&nbsp;<a href="javascript:void(0);" class="genbtn" id="btn_export">匯出Excel</a><br />
        <br />
        <div class="stripeMe">
        </div>
    </div>
    <br />
</asp:Content>

