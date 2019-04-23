<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="ReportMonthList.aspx.cs" Inherits="WebPage_ReportMonthList" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            getData(0);

            $(document).on("click", "#searchbtn", function () {
                getData(0);
            });

            //TypeStr 查詢月報類別change事件
            $("#TypeStr").change(function () {
                getData(0);
            });
            //新增月報
            $(document).on("click", "#subbtn", function () {
                //$("#nType").val("02");
                //$("#nYear").val("107");
                //$("#nMonth").val("07");
                //$("#nStage").val("1");
                if ($("#nType").val() == "" || $("#nYear").val() == "" || $("#nMonth").val() == "" || $("#nStage").val() == "") {
                    alert("尚有選項未選擇");
                    return false;
                }
                else {
                    //設備汰換
                    if ($("#nType").val()=="01") {
                        location.href = "ReportMonth.aspx?year=" + $("#nYear").val() + "&month=" + $("#nMonth").val() + "&stage=" + $("#nStage").val() + "&rmtype=" + $("#nType").val();
                    }
                    //擴大補助
                    if ($("#nType").val()=="02") {
                        location.href = "ReportMonthEx.aspx?year=" + $("#nYear").val() + "&month=" + $("#nMonth").val() + "&stage=" + $("#nStage").val() + "&rmtype=" + $("#nType").val();
                    }
                }
                    
            });
        });//js end

        function getData(p) {
             $.ajax({
                 type: "POST",
                 async: false, //在沒有返回值之前,不會執行下一步動作
                 url: "../handler/getMonthList.aspx",
                 data: {
                     PageNo: p,
                     mtype: $("#TypeStr").val(),
                     year: $("#YearStr").val(),
                     month: $("#MonthStr").val(),
                     stage: $("#StageStr").val()
                 },
                 error: function (xhr) {
                     alert("Error " + xhr.status);
                     console.log(xhr.responseText);
                 },
                 success: function (data) {
                     if ($(data).find("Error").length > 0) {
                         alert($(data).find("Error").attr("Message"));
                     }
                     else {
                         $("#tablist tbody").empty();
                         var tabstr = '',yearEn='';
                         if ($(data).find("data_item").length > 0) {
                             $(data).find("data_item").each(function (i) {
                                 if ($(this).children("RM_Year").text().trim() != "") {
                                     yearEn = (parseInt($(this).children("RM_Year").text().trim()) - 1911).toString();
                                 }
                                 tabstr += '<tr>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("itemNo").text().trim() + '</td>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("MTypeName").text().trim() + '</td>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + yearEn + '</td>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("RM_Month").text().trim() + '</td>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("RM_Stage").text().trim() + '</td>';
                                 if ($(this).children("RC_CheckType").text().trim() == "" && $(this).children("RC_Status").text().trim() == "")
                                     tabstr += '<td align="center" nowrap="nowrap">草稿</td>';
                                 else if ($(this).children("RC_CheckType").text().trim() == "" && $(this).children("RC_Status").text().trim() == "A")
                                     tabstr += '<td align="center" nowrap="nowrap">送審中</td>';
                                 else
                                     tabstr += '<td align="center" nowrap="nowrap">審核通過</td>';
                                 //tabstr += '<td align="center" nowrap="nowrap">' + $.datepicker.formatDate('yy/mm/dd', new Date($(this).children("RM_CreateDate").text().trim())) + '</td>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + $.datepicker.formatDate('yy/mm/dd', new Date($(this).children("RM_CreateDate").text().trim())) + '</td>';
                                 tabstr += '<td align="center">';
                                 if ($(this).children("RC_CheckType").text().trim() == "" && $(this).children("RC_Status").text().trim() == "") {
                                     if ($(this).children("MTypeNum").text().trim()=="01") {//設備汰換
                                         tabstr += '<a class="genbtn" href="ReportMonth.aspx?year=' + $(this).children("RM_Year").text().trim() + '&month=' + $(this).children("RM_Month").text().trim() + '&stage=' + $(this).children("RM_Stage").text().trim() + '&rmtype=' + $(this).children("MTypeNum").text().trim() + '">編輯</a>';
                                     }
                                     if ($(this).children("MTypeNum").text().trim()=="02") {//擴大補助
                                         tabstr += '<a class="genbtn" href="ReportMonthEx.aspx?year=' + $(this).children("RM_Year").text().trim() + '&month=' + $(this).children("RM_Month").text().trim() + '&stage=' + $(this).children("RM_Stage").text().trim() + '&rmtype=' + $(this).children("MTypeNum").text().trim() + '">編輯</a>';
                                     }
                                 }
                                 else if ($(this).children("RC_CheckType").text().trim() == "" && $(this).children("RC_Status").text().trim() == "A")
                                 {
                                     if ($(this).children("MTypeNum").text().trim() =="01") {//設備汰換
                                         tabstr += '<a class="genbtn" href="DetailReportMonth.aspx?v=' + $(this).children("enGuid").text().trim() + '">查看</a>';
                                     }
                                     if ($(this).children("MTypeNum").text().trim() =="02") {//擴大補助
                                         tabstr += '<a class="genbtn" href="DetailReportMonthEx.aspx?v=' + $(this).children("enGuid").text().trim() + '">查看</a>';
                                     }
                                     
                                 }
                                 tabstr += '</td></tr>';
                             });
                         }
                         else
                             tabstr += "<tr><td colspan='8'>查詢無資料</td></tr>";
                         $("#tablist tbody").append(tabstr);
                         $(".stripeMe tr").mouseover(function () { $(this).addClass("spe"); }).mouseout(function () { $(this).removeClass("spe"); });
                         $(".stripeMe table:not(td > table) > tbody > tr:not('.spe'):even").addClass("alt");
                         CreatePage(p, $("total", data).text());
                     }
                 }
             });
        }
        /// 新增視窗
        function doShowDialog() {
            /// dialog setting
            $("#nType").val("");
            $("#NewBlock").dialog({
                title: "新增月報",
                autoOpen: false,
                width: 320,
                height: 290,
                closeOnEscape: true,
                position: { my: "center", at: "center", of: window },
                modal: true,
                resizable: false,
                close: function (event, ui) {
                    $(".dialogInput").val("");
                }
            });

            $("#NewBlock").dialog("open");
        }
        
        //分頁設定
        //ListNum: 每頁顯示資料筆數
        //PageNum: 分頁頁籤顯示數
        PageOption.Selector = "#pageblock";
        PageOption.ListNum = 10;
        PageOption.PageNum = 10;
        PageOption.PrevStep = false;
        PageOption.NextStep = false;
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">月報管理</span></div>
            <div class="right"></div>
        </div>

        <div class="margin10T">
            月報類別：<select id="TypeStr" class="inputex margin10B">
                <option value="" selected="selected">全部</option>
                <option value="01">設備汰換</option>
                <option value="02">擴大補助</option>
            </select><br />
            年：<select id="YearStr" class="inputex" style="margin-right: 10px;">
                <option value="">--請選擇--</option>
                <option value="107">107</option>
                <option value="108">108</option>
                <option value="109">109</option>
            </select>
            月：<select id="MonthStr" class="inputex" style="margin-right: 10px;">
                <option value="">--請選擇--</option>
                <option value="01">1月</option>
                <option value="02">2月</option>
                <option value="03">3月</option>
                <option value="04">4月</option>
                <option value="05">5月</option>
                <option value="06">6月</option>
                <option value="07">7月</option>
                <option value="08">8月</option>
                <option value="09">9月</option>
                <option value="10">10月</option>
                <option value="11">11月</option>
                <option value="12">12月</option>
            </select>
            期：<select id="StageStr" class="inputex"style="margin-right: 10px;">
                <option value="">--請選擇--</option>
                <option value="1">第一期</option>
                <option value="2">第二期</option>
                <option value="3">第三期</option>
            </select>
            <input type="button" id="searchbtn" value="查詢" class="genbtn" />
        </div>
        <div style="text-align:right;">
            <input class="genbtn" type="button" id="addbtn" value="新增月報" onclick="doShowDialog()" />
        </div>

        <div class="stripeMe margin10T font-normal">
            <table id="tablist" width="100%" border="0" cellspacing="0" cellpadding="0">
                <thead>
                    <tr>
                        <th nowrap="nowrap" style="width:50px;">項次</th>
                        <th nowrap="nowrap" style="width:16%;">類別</th>
                        <th nowrap="nowrap" style="width:10%;">年</th>
                        <th nowrap="nowrap" style="width:10%;">月</th>
                        <th nowrap="nowrap" style="width:10%;">期</th>
                        <th nowrap="nowrap" style="width:16%;">狀態</th>
                        <th nowrap="nowrap" style="width:16%;">建立日期</th>
                        <th nowrap="nowrap" style="width:16%;">功能</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
            <!--擴大補助 列表div-->
            <div id="pageblock" class="margin20T textcenter"></div>
        </div>
    </div>

    <!--新增彈出的dialog-->
    <div id="NewBlock" style="display:none; text-align:center;">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td style="text-align:right;">月報類別：</td>
                    <td style="text-align:left;">
                        <select id="nType" class="inputex margin10TB dialogInput">
                            <option value="" selected="selected">--請選擇--</option>
                            <option value="01">設備汰換</option>
                            <option value="02">擴大補助</option>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td style="text-align:right;">年：</td>
                    <td style="text-align:left;">
                        <select id="nYear" class="inputex margin10TB dialogInput">
                            <option value="">--請選擇--</option>
                            <option value="2018">107</option>
                            <option value="2019">108</option>
                            <option value="2020">109</option>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td style="text-align:right;">月：</td>
                    <td style="text-align:left;">
                        <select id="nMonth" class="inputex margin10TB dialogInput">
                            <option value="">--請選擇--</option>
                            <option value="01">1月</option>
                            <option value="02">2月</option>
                            <option value="03">3月</option>
                            <option value="04">4月</option>
                            <option value="05">5月</option>
                            <option value="06">6月</option>
                            <option value="07">7月</option>
                            <option value="08">8月</option>
                            <option value="09">9月</option>
                            <option value="10">10月</option>
                            <option value="11">11月</option>
                            <option value="12">12月</option>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td style="text-align:right;">期：</td>
                    <td style="text-align:left;">
                        <select id="nStage" class="inputex margin10TB dialogInput">
                            <option value="">--請選擇--</option>
                            <option value="1">第一期</option>
                            <option value="2">第二期</option>
                            <option value="3">第三期</option>
                        </select>
                    </td>
                </tr>
            </table>
            <input type="button" class="genbtn" id="subbtn" value="確定" style="margin-top:10px;" />
    </div>
</asp:Content>

