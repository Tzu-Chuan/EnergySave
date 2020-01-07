<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="HistorySeason.aspx.cs" Inherits="WebPage_HistorySeason" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            getddl("02", "#ddlCity")
            $("#ddlCity").val($("#city").val());
            if ($("#SA").val() == "Y") {
                $("#cityspan").show();
                $("#divExport").show();
            }
                

            getData(0);

            //限制只能輸入數字
            $(document).on("keyup", ".num", function () {
                if (/[^0-9]/g.test(this.value)) {
                    this.value = this.value.replace(/[^0-9]/g, '');
                }
            });

            $(document).on("click", "#searchbtn", function () {
                if ($("#startday").val() != "" && $("#endday").val()) {
                    var sday = new Date($("#startday").val());
                    var eday = new Date($("#endday").val());
                    if (eday < sday) {
                        alert("起始日不可大於結束日");
                        return;
                    }
                }

                getData(0);
            });

            ////20190801新增 匯出列表全部資料
            $(document).on("click", "#exportExcelbtn", function () {
                window.location = "../handler/ExportHistorySeasonList.aspx?s=" + $("#exportStage").val() + "";
            });
        });

        function getData(p) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getHistory_S.ashx",
                data: {
                    CurrentPage: p,
                    sday: $("#startday").val(),
                    eday: $("#endday").val(),
                    year: $("#yearstr").val(),
                    season: $("#seasonstr").val(),
                    stage: $("#stagestr").val(),
                    city: $("#ddlCity").val(),
                    SearchStr: $("#SearchStr").val()
                },
                error: function (xhr) {
                    alert(xhr.responseText);
                },
                success: function (data) {
                    if (data.indexOf("Error") > -1)
                        alert(data);
                    else {
                        if (data == "reLogin") {
                            alert("請重新登入");
                            window.location = "Login.aspx";
                            return;
                        }

                        if (data != null) {
                            data = $.parseXML(data);
                            $("#tablist tbody").empty();
                            var tabstr = '';
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function (i) {
                                    tabstr += '<tr>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + (i + 1) + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("City").text().trim() + '</td>';//20190801新增縣市欄位
                                    var year = parseInt($(this).children("RC_Year").text().trim()) - 1911;
                                    tabstr += '<td align="center" nowrap="nowrap">' + year + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + parseInt($(this).children("RC_Season").text().trim()) + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("RC_Stage").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("MbName").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("RC_CheckDate").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("AdName").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap"><a href="DetailReportSeason.aspx?v=' + $(this).children("RS_ID").text().trim() + '" class="genbtn" target="_blank">明細</a></td>';
                                    tabstr += '</tr>';
                                });
                            }
                            else
                                tabstr += "<tr><td colspan='9'>查詢無資料</td></tr>";
                            $("#tablist tbody").append(tabstr);
                            $(".stripeMe tr").mouseover(function () { $(this).addClass("spe"); }).mouseout(function () { $(this).removeClass("spe"); });
                            $(".stripeMe table:not(td > table) > tbody > tr:not('.spe'):even").addClass("alt");
                            CreatePage(p, $("total", data).text());
                        }
                    }
                }
            });
        }

        function getddl(gno, tagName) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/GetDDL.ashx",
                data: {
                    Group: gno
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if (data == "error") {
                        alert("GetDDL Error");
                        return false;
                    }

                    if (data != null) {
                        data = $.parseXML(data);
                        $(tagName).empty();
                        var ddlstr = '<option value="">--請選擇--</option>';
                        if ($(data).find("code").length > 0) {
                            $(data).find("code").each(function (i) {
                                ddlstr += '<option value="' + $(this).attr("v") + '">' + $(this).attr("desc") + '</option>';
                            });
                        }
                        $(tagName).append(ddlstr);
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
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <input id="SA" type="hidden" value="<%= showcity %>" />
    <input id="city" type="hidden" value="<%= city %>" />
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">季報歷史資料</span></div>
            <div class="right">歷史資料 / 季報</div>
        </div>

        <div style="margin-top: 5px;">送審日期：<input type="text" id="startday" class="inputex Jdatepicker" />&nbsp;~&nbsp;<input type="text" id="endday" class="inputex Jdatepicker" /></div>
        <div style="margin-top: 5px;">
            <span id="cityspan" style="display: none;">執行機關：<select id="ddlCity" class="inputex"></select></span>
            年：
            <select id="yearstr" class="inputex">
                <option value="">--請選擇--</option>
                <option value="107">107</option>
                <option value="108">108</option>
                <option value="109">109</option>
            </select>
            季：
            <select id="seasonstr" class="inputex">
                <option value="">--請選擇--</option>
                <option value="1">第一季</option>
                <option value="2">第二季</option>
                <option value="3">第三季</option>
                <option value="4">第四季</option>
            </select>
            期：
            <select id="stagestr" class="inputex">
                <option value="">--請選擇--</option>
                <option value="1">第一期</option>
                <option value="2">第二期</option>
                <option value="3">第三期</option>
            </select>
        </div>
        <div style="margin-top: 5px;">
            關鍵字：<input type="text" id="SearchStr" class="inputex" />&nbsp;
            <input type="button" id="searchbtn" value="查詢" class="genbtn" />
        </div>
        <div style="margin-top:5px;display:none;" id="divExport">
            <select id="exportStage" class="inputex">
                <option value="1" >第一期</option>
                <option value="2" >第二期</option>
                <option value="3" >第三期</option>
            </select>
            <input type="button" id="exportExcelbtn" value="匯出資料" class="genbtn" /><!--20190801新增匯出列表全部資料-->
        </div>
        <br />
        <div class="stripeMe margin5T font-normal">
            <table id="tablist" width="100%" border="0" cellspacing="0" cellpadding="0">
                <thead>
                    <tr>
                        <th nowrap="nowrap" style="width: 50px;">項次</th>
                        <th nowrap="nowrap">縣市</th><!--20190801新增縣市欄位-->
                        <th nowrap="nowrap">年</th>
                        <th nowrap="nowrap">季</th>
                        <th nowrap="nowrap">期</th>
                        <th nowrap="nowrap">承辦人</th>
                        <th nowrap="nowrap" style="width: 300px;">審核日期</th>
                        <th nowrap="nowrap">審核主管</th>
                        <th nowrap="nowrap">詳細資料</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
            <div id="pageblock" class="margin20T textcenter"></div>
        </div>
    </div>
</asp:Content>

