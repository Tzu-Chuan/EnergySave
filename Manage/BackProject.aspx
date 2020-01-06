<%@ Page Title="" Language="C#" MasterPageFile="~/Manage/Admin.master" AutoEventWireup="true" CodeFile="BackProject.aspx.cs" Inherits="Manage_BackProject" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            getData(0);
            
            //明細 button
            $(document).on("click", "input[name='backBtn']", function (e) {
                if (confirm("是否確定將計畫退回草稿狀態?")) {
                    backProject($(this).attr("aid"));
                    //alert($(this).attr("aid"));
                }
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
                                    //tabstr += '<td align="center" nowrap="nowrap">' + pdate + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("M_Name").text().trim() + '</td>';
                                    //tabstr += '<td align="center" nowrap="nowrap">' + $.datepicker.formatDate('yy/mm/dd', new Date($(this).children("I_Modifydate").text().trim())) + '</td>';
                                    //狀態
                                    if ($(this).children("I_Flag").text().trim() == "Y")
                                        tabstr += '<td align="center" nowrap="nowrap" style="color:red;">定稿</td>';
                                    else
                                        tabstr += '<td align="center" nowrap="nowrap">草稿</td>';
                                    //動作
                                    var mbtn = '<input class="genbtn" name="backBtn" type="button" aid="' + $(this).children("M_ID").text().trim() + '" value="退回" />';

                                    if (parseInt($(this).children("CityFlag").text().trim()) > 0) {
                                        if ($(this).children("I_Flag").text().trim() == "Y")
                                            tabstr += '<td align="center" nowrap="nowrap"><input class="genbtn" name="backBtn" type="button" aid="' + $(this).children("I_ID").text().trim() + '" value="退回" /></td >';
                                    }else
                                        tabstr += '<td></td>';
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

        function backProject(pjid) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/backProject.aspx",
                data: {
                   aid:pjid
                },
                error: function (xhr) {
                    alert(xhr.responseText);
                },
                success: function (data) {
                    if ($(data).find("Error").length > 0) {
                        alert($(data).find("Error").attr("Message"));
                    }
                    else {
                        if (data == "reLogin") {
                            alert("請重新登入");
                            window.location = "Login.aspx";
                            return;
                        }
                        alert("退回草稿成功");
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
    <input id="tmpid" runat="server" type="hidden" />
    <div class="container">
        <div class="twocol filetitlewrapper">
	        <div class="left"><span class="filetitle font-size5">計畫書退回</span></div><!-- left -->
            <div class="right">計畫書退回</div><!-- right -->
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
                        <%--<th nowrap="nowrap">執行期程</th>--%>
                        <th nowrap="nowrap">承辦人</th>
                        <%--<th nowrap="nowrap">更新日期</th>--%>
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

