<%@ Page Title="" Language="C#" MasterPageFile="~/Manage/Admin.master" AutoEventWireup="true" CodeFile="Log_ModPw.aspx.cs" Inherits="Manage_Log_ModPw" %>

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
        });

        function getData(p) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getLogModPw.ashx",
                data: {
                    CurrentPage: p,
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
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("ModPerson").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("L_IP").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("L_Desc").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("LoginPerson").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + datetimeformat($(this).children("L_ModDate").text().trim()) + '</td>';
                                    tabstr += '</tr>';
                                });
                            }
                            else
                                tabstr += "<tr><td colspan='8'>查詢無資料</td></tr>";
                            $("#tablist tbody").append(tabstr);
                            $(".stripeMe tr").mouseover(function () { $(this).addClass("spe"); }).mouseout(function () { $(this).removeClass("spe"); });
                            $(".stripeMe table:not(td > table) > tbody > tr:not('.spe'):even").addClass("alt");
                            PageFun(p, $("total", data).text());
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
    <div class="container">
        <div class="twocol filetitlewrapper">
	        <div class="left"><span class="filetitle font-size5">密碼修改紀錄</span></div><!-- left -->
            <div class="right">紀錄查詢 > 密碼修改紀錄</div><!-- right -->
        </div><!-- twocol -->

        <div style="margin-top:10px;">日期：<input type="text" id="startday" class="inputex" />&nbsp;~&nbsp;<input type="text" id="endday" class="inputex" /></div>
	    <div style="margin-top:5px;">關鍵字：<input type="text" id="SearchStr" class="inputex" />&nbsp;<input type="button" id="searchbtn" value="查詢" class="genbtn" /></div>
        <br />
        <div class="stripeMe margin5T font-normal">
            <table id="tablist" width="100%" border="0" cellspacing="0" cellpadding="0">
                <thead>
                    <tr>
                        <th nowrap="nowrap" style="width:40px;">項次</th>
                        <th nowrap="nowrap">異動成員</th>
                        <th nowrap="nowrap">IP</th>
                        <th nowrap="nowrap">描述</th>
                        <th nowrap="nowrap">修改者</th>
                        <th nowrap="nowrap">建立日期</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
            <div id="changpage" class="margin20T textcenter"></div>
        </div>
    </div>
</asp:Content>

