<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="SeasonList.aspx.cs" Inherits="WebPage_SeasonList" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            getData(0);

            $(document).on("click", "#searchbtn", function () {
                getData(0);
            });

            $(document).on("click", "#subbtn", function () {
                if ($("#nYear").val() == "" || $("#nSeason").val() == "" || $("#nStage").val() == "") {
                    alert("尚有選項未選擇");
                    return false;
                }
                else
                    location.href = "ReportSeason.aspx?year=" + $("#nYear").val() + "&season=" + $("#nSeason").val() + "&stage=" + $("#nStage").val();
            });
            //刪除 草稿
            $(document).on("click", "a[name='delBtn']", function () {
                delData($(this).attr("id"))
            });
        });//js end

        function getData(p) {
             $.ajax({
                 type: "POST",
                 async: false, //在沒有返回值之前,不會執行下一步動作
                 url: "../handler/getSeasonList.aspx",
                 data: {
                     PageNo: p,
                     year: $("#YearStr").val(),
                     season: $("#SeasonStr").val(),
                     stage: $("#StageStr").val()
                 },
                 error: function (xhr) {
                     alert(xhr.responseText);
                 },
                 success: function (data) {
                     if ($(data).find("Error").length > 0) {
                         alert($(data).find("Error").attr("Message"));
                     }
                     else {
                         $("#tablist tbody").empty();
                         var tabstr = '';
                         if ($(data).find("data_item").length > 0) {
                             $(data).find("data_item").each(function (i) {
                                 tabstr += '<tr>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("itemNo").text().trim() + '</td>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("RS_Year").text().trim() + '</td>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("RS_Season").text().trim() + '</td>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("RS_Stage").text().trim() + '</td>';
                                 if ($(this).children("RC_CheckType").text().trim() == "" && $(this).children("RC_Status").text().trim() == "")
                                     tabstr += '<td align="center" nowrap="nowrap">草稿</td>';
                                 else if ($(this).children("RC_CheckType").text().trim() == "" && $(this).children("RC_Status").text().trim() == "A")
                                     tabstr += '<td align="center" nowrap="nowrap">送審中</td>';
                                 else
                                     tabstr += '<td align="center" nowrap="nowrap">審核通過</td>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + $.datepicker.formatDate('yy/mm/dd', new Date($(this).children("RS_CreateDate").text().trim())) + '</td>';
                                 tabstr += '<td align="center">';
                                 if ($(this).children("RC_CheckType").text().trim() == "" && $(this).children("RC_Status").text().trim() == "") {
                                     tabstr += '<a class="genbtn" href="ReportSeason.aspx?year=' + $(this).children("RS_Year").text().trim() + '&season=' + $(this).children("RS_Season").text().trim() + '&stage=' + $(this).children("RS_Stage").text().trim() + '">編輯</a>';
                                     tabstr += '<a class="genbtn" name="delBtn" href="javascript:void(0);" id="' + $(this).children("RS_Guid").text().trim() + '" >刪除</a>';
                                 }
                                 else if ($(this).children("RC_CheckType").text().trim() == "" && $(this).children("RC_Status").text().trim() == "A")
                                     tabstr += '<a class="genbtn" href="DetailReportSeason.aspx?v='+$(this).children("RS_ID").text().trim()+'">查看</a>';
                                 tabstr += '</td></tr>';
                             });
                         }
                         else
                             tabstr += "<tr><td colspan='7'>查詢無資料</td></tr>";
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
            $("#NewBlock").dialog({
                title: "新增季報",
                autoOpen: false,
                width: 300,
                height: 250,
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

        //刪除 草稿
        function delData(strRguid) {
             $.ajax({
                 type: "POST",
                 async: false, //在沒有返回值之前,不會執行下一步動作
                 url: "../handler/delReporSeason.aspx",
                 data: {
                     strReportGuid: strRguid
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
                         alert("刪除成功");
                         getData(0);
                     }
                 }
             });
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
            <div class="left"><span class="filetitle font-size5">季報管理</span></div>
            <div class="right"></div>
        </div>

        <div class="margin10T">
            年：<select id="YearStr" class="inputex" style="margin-right: 10px;">
                <option value="">--請選擇--</option>
                <option value="107">107</option>
                <option value="108">108</option>
                <option value="109">109</option>
            </select>
            季：<select id="SeasonStr" class="inputex" style="margin-right: 10px;">
                <option value="">--請選擇--</option>
                <option value="1">第一季</option>
                <option value="2">第二季</option>
                <option value="3">第三季</option>
                <option value="4">第四季</option>
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
            <input class="genbtn" type="button" id="addbtn" value="新增季報" onclick="doShowDialog()" />
        </div>
        <div class="stripeMe margin10T font-normal">
            <table id="tablist" width="100%" border="0" cellspacing="0" cellpadding="0">
                <thead>
                    <tr>
                        <th nowrap="nowrap" style="width:50px;">項次</th>
                        <th nowrap="nowrap" style="width:16%;">年</th>
                        <th nowrap="nowrap" style="width:16%;">季</th>
                        <th nowrap="nowrap" style="width:16%;">期</th>
                        <th nowrap="nowrap" style="width:16%;">狀態</th>
                        <th nowrap="nowrap" style="width:16%;">建立日期</th>
                        <th nowrap="nowrap" style="width:16%;">功能</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
            <div id="pageblock" class="margin20T textcenter"></div>
        </div>
    </div>

    <div id="NewBlock" style="display:none; text-align:center;">
            年：
            <select id="nYear" class="inputex margin10TB dialogInput">
                <option value="">--請選擇--</option>
                <option value="107">107</option>
                <option value="108">108</option>
                <option value="109">109</option>
            </select><br />
            季：
            <select id="nSeason" class="inputex margin10B dialogInput">
                <option value="">--請選擇--</option>
                <option value="1">第一季</option>
                <option value="2">第二季</option>
                <option value="3">第三季</option>
                <option value="4">第四季</option>
            </select><br />
            期：
            <select id="nStage" class="inputex margin20B dialogInput">
                <option value="">--請選擇--</option>
                <option value="1">第一期</option>
                <option value="2">第二期</option>
                <option value="3">第三期</option>
            </select><br />
            <input type="button" class="genbtn" id="subbtn" value="確定" />
    </div>
</asp:Content>

