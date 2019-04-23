<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="ProjectList.aspx.cs" Inherits="WebPage_ProjectList" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            getData(0);

            //新增 button
            $(document).on("click", "#addbtn", function () {
                location.href = "ProjectInfo.aspx?v=" + $("#<%= tmpid.ClientID %>").val();
            });

            //修改 button
            $(document).on("click", "input[name='modbtn']", function () {
                var mid = $(this).attr("aid");
                location.href = "ProjectInfo.aspx?v=" + mid;
            });

            //明細 button
            $(document).on("click", "input[name='previewbtn']", function () {
                var mid = $(this).attr("aid");
                location.href = "Preview.aspx?v=" + mid;
            });

            //明細 button
            $(document).on("keypress", "#SearchStr", function (e) {
                if ((e.keyCode == 13) || (e.key == "Enter") || (e.code == "Enter"))
                    getData(0);
            });
        });

        function getData(p) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getProjectList.ashx",
                data: {
                    CurrentPage: p,
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
                                    tabstr += '<tr aid="' + $(this).children("I_ID").text().trim() + '">';
                                    tabstr += '<td align="center" nowrap="nowrap">' + (i + 1) + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("City").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("M_Office").text().trim() + '</td>';
                                    var pdate = $.datepicker.formatDate('yy/mm/dd', new Date($(this).children("PD_StartDate").text().trim())) + ' - ' + $.datepicker.formatDate('yy/mm/dd', new Date($(this).children("PD_EndDate").text().trim()));
                                    tabstr += '<td align="center" nowrap="nowrap">' + pdate + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("M_Name").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $.datepicker.formatDate('yy/mm/dd', new Date($(this).children("I_Modifydate").text().trim())) + '</td>';
                                    //狀態
                                    if ($(this).children("I_Flag").text().trim() == "Y")
                                        tabstr += '<td align="center" nowrap="nowrap" style="color:red;">定稿</td>';
                                    else
                                        tabstr += '<td align="center" nowrap="nowrap">草稿</td>';
                                    //動作
                                    var mbtn = '<input class="genbtn" name="modbtn" type="button" aid="' + $(this).children("M_ID").text().trim() + '" value="修改" />';
                                    var pbtn = '<input class="genbtn" name="previewbtn" type="button" aid="' + $(this).children("M_ID").text().trim() + '" value="明細" />';
                                    switch ($("comp", data).text()) {
                                        case "SA":
                                            if (parseInt($(this).children("CityFlag").text().trim()) > 0) {
                                                if ($(this).children("I_Flag").text().trim() == "Y")
                                                    tabstr += '<td align="center" nowrap="nowrap">' + mbtn + '&nbsp;' + pbtn + '</td >';
                                                else
                                                    tabstr += '<td align="center" nowrap="nowrap">' + pbtn + '</td >';
                                            }
                                            else
                                                tabstr += '<td align="center" nowrap="nowrap">' + pbtn + '</td >';
                                            break;
                                        case "01":
                                            if ($("cityflag", data).text() == "Y")
                                                tabstr += '<td align="center" nowrap="nowrap">' + pbtn + '</td >';
                                            else {
                                                if ($("myGuid", data).text() == $(this).children("I_People").text().trim())
                                                    tabstr += '<td align="center" nowrap="nowrap">' + mbtn + '</td >';
                                                else
                                                    tabstr += '<td align="center" nowrap="nowrap">' + pbtn + '</td >';
                                            }

                                            //新增按鈕
                                            if ($("checkinfo", data).text() != "Y" && $("cityflag", data).text() != "Y")
                                                $("#addbtn").show();
                                            break;
                                        case "02":
                                            tabstr += '<td align="center" nowrap="nowrap">' + pbtn + '</td >';
                                            break;
                                    }
                                    tabstr += '</tr>';
                                });
                            }
                            else {
                                tabstr += "<tr><td colspan='8'>查詢無資料</td></tr>";

                                //新增按鈕
                                if ($("checkinfo", data).text() != "Y")
                                    $("#addbtn").show();
                            }
                            $("#tablist tbody").append(tabstr);
                            $(".stripeMe tr").mouseover(function () { $(this).addClass("spe"); }).mouseout(function () { $(this).removeClass("spe"); });
                            $(".stripeMe table:not(td > table) > tbody > tr:not('.spe'):even").addClass("alt");
                            CreatePage(p, $("total", data).text());
                        }
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
    <input id="tmpid" runat="server" type="hidden" />
    <div class="container">
        <div class="twocol filetitlewrapper">
	        <div class="left"><span class="filetitle font-size5">計畫列表</span></div><!-- left -->
            <div class="right">計畫列表</div><!-- right -->
        </div><!-- twocol -->
        <br />
        <div class="twocol">
	        <div class="left">關鍵字：<input type="text" class="inputex" id="SearchStr" /><input type="text" style="display:none;"/>&nbsp;<input type="button" class="genbtn" value="查詢" onclick="getData(0)" /></div><!-- left -->
            <div class="right"><input type="button" class="genbtn" value="新增計畫基本資料" id="addbtn" style="display:none;" /></div><!-- right -->
        </div><!-- twocol -->
        <br />
        <div class="stripeMe margin5T font-normal">
            <table width="100%" border="0" cellspacing="0" cellpadding="0" id="tablist">
                <thead>
                    <tr>
                        <th nowrap="nowrap">項次</th>
                        <th nowrap="nowrap">執行機關</th>
                        <th nowrap="nowrap">承辦局處</th>
                        <th nowrap="nowrap">執行期程</th>
                        <th nowrap="nowrap">承辦人</th>
                        <th nowrap="nowrap">更新日期</th>
                        <th nowrap="nowrap">狀態</th>
                        <th nowrap="nowrap">動作</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
            <div id="pageblock" class="margin20T textcenter"></div>
        </div>
    </div>
</asp:Content>

