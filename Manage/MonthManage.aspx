<%@ Page Title="" Language="C#" MasterPageFile="~/Manage/Admin.master" AutoEventWireup="true" CodeFile="MonthManage.aspx.cs" Inherits="Manage_MonthManage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script>
        $(document).ready(function () {
            getddl("02", "#ddlCity");
            getData(0);

            $(document).on("change", "#ddlCity", function () {
                getData(0);
            });


            $(document).on("click", "input[name='delbtn']", function () {
                if (confirm("確定抽單?")) {
                    $.ajax({
                        type: "POST",
                        async: false, //在沒有返回值之前,不會執行下一步動作
                        url: "../handler/CancelReport.aspx",
                        data: {
                            type: "Month",
                            id: $(this).attr("aid")
                        },
                        error: function (xhr) {
                            alert(xhr.responseText);
                        },
                        success: function (data) {
                            if ($(data).find("Error").length > 0) {
                                alert($(data).find("Error").attr("Message"));
                            }
                            else {
                                alert("抽單完成");
                                getData(0);
                            }
                        }
                    });
                }
            });
        });// end js

         function getData(p) {
             $.ajax({
                 type: "POST",
                 async: false, //在沒有返回值之前,不會執行下一步動作
                 url: "../handler/getMonthManageList.aspx",
                 data: {
                     PageNo: p,
                     city: $("#ddlCity").val()
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
                         var tabstr = '';
                         if ($(data).find("data_item").length > 0) {
                             $(data).find("data_item").each(function (i) {
                                 tabstr += '<tr>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + (i + 1) + '</td>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("City").text().trim() + '</td>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("I_Office").text().trim() + '</td>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + (parseInt($(this).children("RC_Year").text().trim()) - 1911) + '</td>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("RC_Month").text().trim() + '</td>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("M_Name").text().trim() + '</td>';
                                 tabstr += '<td align="center" nowrap="nowrap">' + $.datepicker.formatDate('yy/mm/dd', new Date($(this).children("RC_CreateDate").text().trim())) + '</td>';
                                 tabstr += '<td align="center">';
                                 tabstr += '<input type="button" class="genbtn" name= "delbtn" aid= "' + $(this).children("RC_Guid").text().trim() + '" value="抽單" />';
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
                        var ddlstr = '<option value="">全部</option>';
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
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
	        <div class="left"><span class="filetitle font-size5">月報管理</span></div><!-- left -->
            <div class="right"></div><!-- right -->
        </div><!-- twocol -->

        <div class="twocol" style="margin-top:20px;">
	        <div class="left">執行機關：<select id="ddlCity" class="inputex"></select></div><!-- left -->
            <div class="right"></div><!-- right -->
        </div><!-- twocol -->
        <br />
        <div class="stripeMe margin5T font-normal">
            <table id="tablist" width="100%" border="0" cellspacing="0" cellpadding="0">
                <thead>
                    <tr>
                        <th nowrap="nowrap" style="width:40px;">項次</th>
                        <th nowrap="nowrap">執行機關</th>
                        <th nowrap="nowrap">承辦局處</th>
                        <th nowrap="nowrap">年</th>
                        <th nowrap="nowrap">月</th>
                        <th nowrap="nowrap">承辦人</th>
                        <th nowrap="nowrap" style="width:150px;">送審日期</th>
                        <th nowrap="nowrap" style="width:150px;">動作</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
            <div id="pageblock" class="margin20T textcenter"></div>
        </div>
    </div>
</asp:Content>

