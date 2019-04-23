<%@ Page Title="" Language="C#" MasterPageFile="~/Manage/Admin.master" AutoEventWireup="true" CodeFile="setProjectDate.aspx.cs" Inherits="Manage_setProjectDate" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            getddl("02", "#ddlCity");
            getDate();
            
            //datepicker
            $("#startday,#endday").datepicker({
                changeMonth: true,
                changeYear: true,
                dateFormat: 'yy/mm/dd',
                dayNamesMin: ["日", "一", "二", "三", "四", "五", "六"],
                monthNamesShort: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"],
                yearRange: '-100:+100'
            });

            //儲存 button
            $(document).on("click", "#savebtn", function () {
                if ($("#ddlCity").val() == "") {
                    alert("請選擇執行機關");
                    return;
                }

                if ($("#startday").val() != "" && $("#endday").val()) {
                    var sday = new Date($("#startday").val());
                    var eday = new Date($("#endday").val());
                    if (eday < sday) {
                        alert("起始日不可大於結束日");
                        return;
                    }
                }

                $.ajax({
                    type: "POST",
                    async: true, //在沒有返回值之前,不會執行下一步動作
                    url: "../handler/setProjectDay.ashx",
                    data: {
                        type: $("#ddlCity").val(),
                        sday: $("#startday").val(),
                        eday: $("#endday").val()
                    },
                    error: function (xhr) {
                        alert("Error " + xhr.status);
                        console.log(xhr.responseText);
                        $.unblockUI();
                    },
                    beforeSend: function () {
                        $.blockUI({ message: '資料儲存中...' });
                    },
                    success: function (data) {
                        if (data.indexOf("Error") > -1) {
                            alert(data);
                            $.unblockUI();
                        }
                        else {
                            if (data == "reLogin") {
                                alert("請重新登入");
                                window.location = "../WebPage/Login.aspx";
                                return;
                            }

                            if (data == "succeed") {
                                getDate();
                            }
                        }
                    }
                });
            });

            //執行機關 change
            $(document).on("change", "#ddlCity", function () {
                $.ajax({
                    type: "POST",
                    async: false, //在沒有返回值之前,不會執行下一步動作
                    url: "../handler/getProjectDay.ashx",
                    data: {
                        type: $("#ddlCity").val()
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
                                if ($(data).find("data_item").length > 0) {
                                    $(data).find("data_item").each(function (i) {
                                        $("#startday").val(dateFormat($(this).children("PD_StartDate").text().trim()));
                                        $("#endday").val(dateFormat($(this).children("PD_EndDate").text().trim()));
                                    });
                                }
                            }
                        }
                    }
                });
            });
        });//js end

        function getDate() {
            $.ajax({
                type: "POST",
                async: true, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getProjectDay.ashx",
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                    $.unblockUI();
                },
                success: function (data) {
                    if (data.indexOf("Error") > -1) {
                        alert(data);
                        $.unblockUI();
                    }
                    else {
                        if (data != null) {
                            data = $.parseXML(data);
                            $("#tablist").find("tbody").empty();
                            var tabstr = '<tbody>';
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function (i) {
                                    tabstr += '<tr>';
                                    tabstr += '<td align="center">' + $(this).children("PD_Name").text().trim() + '</td>';
                                    tabstr += '<td align="center">' + dateFormat($(this).children("PD_StartDate").text().trim()) + '</td>';
                                    tabstr += '<td align="center">' + dateFormat($(this).children("PD_EndDate").text().trim()) + '</td>';
                                    tabstr += '</tr>';
                                });
                            }
                            tabstr += '</tbody>';
                            $("#tablist").append(tabstr);
                            $("#tablist").show();
                            $.unblockUI();
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

        function dateFormat(day) {
            var tmpday = day;
            if (tmpday.indexOf("1900") < 0)
                tmpday = $.datepicker.formatDate('yy/mm/dd', new Date(tmpday));
            else
                tmpday = "";
            return tmpday;
        }
</script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
	        <div class="left"><span class="filetitle font-size5">專案期程</span></div><!-- left -->
            <div class="right"></div><!-- right -->
        </div><!-- twocol -->
        <div style="margin-top:20px;">
            執行機關：<select id="ddlCity" class="inputex"></select>
        </div>
        <div style="margin-top:10px;">
            起始日期：<input type="text" id="startday" class="inputex pdate" maxlength="10" />&nbsp;&nbsp;
            結束日期：<input type="text" id="endday" class="inputex pdate" maxlength="10" />&nbsp;&nbsp;
            <input type="button" id="savebtn" class="genbtn" value="儲存" />
        </div>
        <br />
         <div class="stripeMe margin5T font-normal">
             各機關期程列表
            <table id="tablist" width="60%" border="0" cellspacing="0" cellpadding="0" style="display:none;">
                <thead>
                    <tr>
                        <th>執行機關</th>
                        <th>起始日期</th>
                        <th>結束日期</th>
                    </tr>
                </thead>
            </table>
        </div>
    </div>
</asp:Content>

