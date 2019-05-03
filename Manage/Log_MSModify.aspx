<%@ Page Title="" Language="C#" MasterPageFile="~/Manage/Admin.master" AutoEventWireup="true" CodeFile="Log_MSModify.aspx.cs" Inherits="Manage_Log_MSModify" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            getData(0);

            //datepicker
            $("#startday,#endday").datepicker({
                changeMonth: true,
                changeYear: true,
                dateFormat: 'yy/mm/dd',
                dayNamesMin: ["日", "一", "二", "三", "四", "五", "六"],
                monthNamesShort: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"],
                yearRange: '-100:+100'
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

            $(document).on("change", "input[name='MSstatus']", function () {
                getData(0);
            });
        });

        function getData(p) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getLogMSModify.ashx",
                data: {
                    CurrentPage: p,
                    type: $("input[name='MSstatus']:checked").val(),
                    sday: $("#startday").val(),
                    eday: $("#endday").val(),
                    SearchStr: $("#SearchStr").val()
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if (data.indexOf("Error") > -1)
                        alert(data);
                    else {
                        if (data != null) {
                            data = $.parseXML(data);
                            $("#tablist tbody").empty();
                            var tabstr = '';
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function (i) {
                                    tabstr += '<tr>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + (i + 1) + '</td>';
                                    var tmpType = ($(this).children("L_Type").text().trim() == "08") ? $(this).children("LType").text().trim().substr(0, 2) + "-" + $(this).children("MTpye").text().trim() : $(this).children("LType").text().trim().substr(0, 2);
                                    tabstr += '<td align="center" nowrap="nowrap">' + tmpType + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("City").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("M_Office").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("M_Name").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("L_IP").text().trim() + '</td>';
                                    if ($(this).children("L_Type").text().trim() == "08")
                                        tabstr += '<td align="center" nowrap="nowrap">' + (parseInt($(this).children("MYear").text().trim()) - 1911).toString() + '年 - ' + $(this).children("MMonth").text().trim() + '月</td>';
                                    else
                                        tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("SYear").text().trim() + '年 - 第' + $(this).children("SSeason").text().trim() + '季</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + datetimeformat($(this).children("L_ModDate").text().trim()) + '</td>';
                                    tabstr += '</tr>';
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
                }
            });
        } 

        function datetimeformat(v) {
            var date = new Date(v);
            var yyyy = date.getFullYear();
            var MM = date.getMonth() + 1;
            var dd = date.getDate();
            var time = date.toLocaleTimeString();
            return yyyy + "/" + MM + "/" + dd + " " + time;
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
	        <div class="left"><span class="filetitle font-size5">月 / 季報更動紀錄</span></div><!-- left -->
            <div class="right">紀錄查詢 > 月 / 季報更動紀錄</div><!-- right -->
        </div><!-- twocol -->

        <div style="margin-top:10px;">
            <input type="radio" value="" name="MSstatus" checked="checked" />&nbsp;全部&nbsp;
            <input type="radio" value="08" name="MSstatus"/>&nbsp;月報&nbsp;
            <input type="radio" value="09" name="MSstatus" />&nbsp;季報
	    </div>
        <div style="margin-top:10px;">日期：<input type="text" id="startday" class="inputex" />&nbsp;~&nbsp;<input type="text" id="endday" class="inputex" /></div>
	    <div style="margin-top:5px;">關鍵字：<input type="text" id="SearchStr" class="inputex" />&nbsp;<input type="button" id="searchbtn" value="查詢" class="genbtn" /></div>
        <br />
        <div class="stripeMe margin5T font-normal">
            <table id="tablist" width="100%" border="0" cellspacing="0" cellpadding="0">
                <thead>
                    <tr>
                        <th nowrap="nowrap" style="width:40px;">項次</th>
                        <th nowrap="nowrap">類別</th>
                        <th nowrap="nowrap">執行機關</th>
                        <th nowrap="nowrap">承辦局處</th>
                        <th nowrap="nowrap">更動人</th>
                        <th nowrap="nowrap">IP</th>
                        <th nowrap="nowrap">月報：年-月<br />季報：年-季</th>
                        <th nowrap="nowrap">建立日期</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
            <div id="pageblock" class="margin20T textcenter"></div>
        </div>
    </div>
</asp:Content>

