<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="HistoryMonth.aspx.cs" Inherits="WebPage_HistoryMonth" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
     <script type="text/javascript">
         $(document).ready(function () {
             if ($("#SA").val() == "Y") {
                 $("#divExport").show();
            }
             getddl("02", "#ddlCity");
             getddl("02", "#ddlCityExport");
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
                 window.location = "../handler/ExportHistoryMonthList.aspx?s=" + $("#exportStage").val() + "&city="+$("#ddlCityExport").val()+"";
            });
        });

        function getData(p) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getHistory_M.ashx",
                data: {
                    CurrentPage: p,
                    sday: $("#startday").val(),
                    eday: $("#endday").val(),
                    stage: $("#stagestr").val(),
                    year: $("#yearstr").val(),
                    month: $("#monthstr").val(),
                    city: $("#ddlCity").val(),
                    SearchStr: $("#SearchStr").val(),
                    reporttype: $("#reportTypestr").val()
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
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
                                    if ($(this).children("RC_ReportType").text().trim() == "01")
                                        tabstr += '<td align="center" nowrap="nowrap">設備汰換</td>';
                                    else if ($(this).children("RC_ReportType").text().trim() == "03")
                                        tabstr += '<td align="center" nowrap="nowrap">擴大補助</td>';
                                    else
                                        tabstr += '<td align="center" nowrap="nowrap"></td>';

                                    var year = parseInt($(this).children("RC_Year").text().trim()) - 1911;
                                    tabstr += '<td align="center" nowrap="nowrap">' + year + '</td>';

                                    
                                    tabstr += '<td align="center" nowrap="nowrap">' + parseInt($(this).children("RC_Month").text().trim()) + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("MbName").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("RC_CheckDate").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("AdName").text().trim() + '</td>';
                                    if($(this).children("RC_ReportType").text().trim() == "01")
                                        tabstr += '<td align="center" nowrap="nowrap"><a href="DetailReportMonth.aspx?v=' + $(this).children("enGuid").text().trim() + '" class="genbtn">明細</a></td>';
                                    if($(this).children("RC_ReportType").text().trim() == "03")
                                        tabstr += '<td align="center" nowrap="nowrap"><a href="DetailReportMonthEx.aspx?v=' + $(this).children("enGuid").text().trim() + '" class="genbtn">明細</a></td>';
                                    tabstr += '</tr>';
                                });
                            }
                            else
                                tabstr += "<tr><td colspan='8'>查詢無資料</td></tr>";
                            $("#tablist tbody").append(tabstr);
                            if ($("comp", data).text() == "SA")
                                $("#cityspan").show();
                            $(".stripeMe tr").mouseover(function () { $(this).addClass("spe"); }).mouseout(function () { $(this).removeClass("spe"); });
                            $(".stripeMe table:not(td > table) > tbody > tr:not('.spe'):even").addClass("alt");
                            PageFun(p, $("total", data).text());
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

        var nowPage = 0; //當前頁
        var listNum = 10; //每頁顯示個數
        var PagesLen; //總頁數 
        var PageNum = 4; //下方顯示分頁數(PageNum+1個)
        function PageFun(PageNow, TotalItem) {
            //Math.ceil -> 無條件進位
            PagesLen = Math.ceil(TotalItem / listNum);
            if (PagesLen <= 1)
                $("#changpage").hide();
            else {
                $("#changpage").show();
                upPage(PageNow, TotalItem);
            }
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <input id="SA" type="hidden" value="<%= showcity %>" />
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">月報歷史資料</span></div>
            <div class="right">歷史資料 / 月報</div>
        </div>

        <div style="margin-top:5px;">
            月報類別：
            <select id="reportTypestr" class="inputex">
                <option value="">--請選擇--</option>
                <option value="01">設備汰換</option>
                <option value="03">擴大補助</option>
            </select>
        </div>
        <div style="margin-top:5px;">審核日期：<input type="text" id="startday" class="inputex Jdatepicker" />&nbsp;~&nbsp;<input type="text" id="endday" class="inputex Jdatepicker" /></div>
        <div style="margin-top:5px;">
            <span id="cityspan" style="display:none;">執行機關：<select id="ddlCity"></select></span>
            期：
            <select id="stagestr" class="inputex">
                <option value="">--請選擇--</option>
                <option value="1">第一期</option>
                <option value="2">第二期</option>
                <option value="3">第三期</option>
            </select>
            年：<input type="text" id="yearstr" class="inputex num" style="width:50px;" maxlength="3" />
            月：
            <select id="monthstr" class="inputex">
                <option value="">--請選擇--</option>
                <option value="01">1</option>
                <option value="02">2</option>
                <option value="03">3</option>
                <option value="04">4</option>
                <option value="05">5</option>
                <option value="06">6</option>
                <option value="07">7</option>
                <option value="08">8</option>
                <option value="09">9</option>
                <option value="10">10</option>
                <option value="11">11</option>
                <option value="12">12</option>
            </select>
        </div>
	    <div>
            關鍵字：<input type="text" id="SearchStr" class="inputex" />&nbsp;
            <input type="button" id="searchbtn" value="查詢" class="genbtn" />
	    </div>
        <div style="margin-top:5px;display:none;" id="divExport">
            <select id="exportStage" class="inputex">
                <option value="1" >第一期</option>
                <option value="2" >第二期</option>
                <option value="3" >第三期</option>
            </select>
            <select id="ddlCityExport"></select>
            <input type="button" id="exportExcelbtn" value="匯出資料" class="genbtn" /><!--20190801新增匯出列表全部資料-->
        </div>
        <br />
        <div class="stripeMe margin5T font-normal">
            <table id="tablist" width="100%" border="0" cellspacing="0" cellpadding="0">
                <thead>
                    <tr>
                        <th nowrap="nowrap" style="width:50px;">項次</th>
                        <th nowrap="nowrap">縣市</th><!--20190801新增縣市欄位-->
                        <th nowrap="nowrap">類別</th>
                        <th nowrap="nowrap">年</th>
                        <th nowrap="nowrap">月</th>
                        <th nowrap="nowrap">承辦人</th>
                        <th nowrap="nowrap" style="width:300px;">審核日期</th>
                        <th nowrap="nowrap">審核主管</th>
                        <th nowrap="nowrap">詳細資料</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
            <div id="changpage" class="margin20T textcenter"></div>
        </div>
    </div>
</asp:Content>

