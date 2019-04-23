<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="Preview.aspx.cs" Inherits="WebPage_Preview" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            if ($.getParamValue('sub') == "1")
                $("#backbtn").attr("href", "Progress.aspx?v=" + $.getParamValue('v'));
            else
                $("#backbtn").attr("href", "ProjectList.aspx");

            getInfoData();
            getCheckPoint("01");
            getCheckPoint("02");
            getCheckPoint("03");
            getCheckPoint("04");
            getProgressData();

            $(document).on("change", "#cp_period", function () {
                getCheckPoint("01");
                getCheckPoint("02");
                getCheckPoint("03");
                getCheckPoint("04");
                getProgressData(); //判斷session是否timeout
            });

            $(document).on("change", "#p_period", function () {
                getProgressData();
            });

            $(document).on("click", "#backbtn", function () {
                location.href = "Progress.aspx?v=" + $.getParamValue('v');
            });

            $(document).on("click", "#subbtn", function () {
                if ($("#nullstatus").val() != "") {
                    alert($("#nullstatus").val());
                    return;
                }

                if (confirm('確認送出？')) {
                    $.ajax({
                        type: "POST",
                        async: true, //在沒有返回值之前,不會執行下一步動作
                        url: "../handler/subInfo.ashx",
                        data: {
                            id: $.getParamValue('v')
                        },
                        error: function (xhr) {
                            alert("Error " + xhr.status);
                            console.log(xhr.responseText);
                        },
                        success: function (data) {
                            if (data.indexOf("Error") > -1) {
                                alert(data);
                            }
                            else {
                                if (data == "reLogin") {
                                    alert("請重新登入");
                                    window.location = "Login.aspx";
                                    return;
                                }

                                if (data == "succeed") {
                                    alert("完成");
                                    location.href = "ProjectList.aspx";
                                }
                            }
                        }
                    });
                }
            });

            //匯出 button
            $(document).on("click", "#exbtn", function () {
                openExport();
                // window.location = "../Handler/ExportInfo.aspx?v=" + $.getParamValue('v');
            });

            //確認匯出
            $(document).on("click", "#expbtn", function () {
                var type = $("input[name='extype']:checked").val();
                window.open("../Handler/ExportInfo.aspx?v=" + $.getParamValue('v') + "&tp=" + type);
                $.fancybox.close();
            });
        });

        function openExport() {
            $.fancybox({
                href: "#exblock",
                title: "",
                //closeBtn: false,
                minWidth: "200",
                minHeight: "100",
                wrapCSS: 'fancybox-custom',
                openEffect: 'elastic',
                closeEffect: 'elastic',
                helpers: {
                    title: {
                        type: 'inside'
                    },
                    overlay: {
                        css: {
                            'background': 'gary'
                        },
                        locked: false   //開始fancybox時,背景是否回top
                        //closeClick: false //點背景關閉 fancybox
                    }
                }
            });
        }

        function getInfoData() {
            $.ajax({
                type: "POST",
                async: true, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getPreview.aspx",
                data: {
                    id: $.getParamValue('v')
                },
                error: function (xhr) {
                    alert(xhr.responseText);
                },
                success: function (data) {
                    if ($(data).find("Error").length > 0) {
                        alert($(data).find("Error").attr("Message"));
                    }
                    else {
                        if (data != null) {
                            if ($(data).find("unVisiable").text() == "Y")
                                $("#backbtn").show();
                            else {
                                $("#subbtn").show();
                                $("#backbtn").show();
                            }
                            var tabstr = '';
                            //基本資料
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function (i) {
                                    $("#I_City").html($(this).children("CityName").text().trim());
                                    $("#I_Office").html($(this).children("I_Office").text().trim());
                                    $("#startdate").html(transROCdate($(this).children("PD_StartDate").text().trim()));
                                    $("#enddate").html(transROCdate($(this).children("PD_EndDate").text().trim()));
                                    $("#I_1_Sdate").html(transROCdate($(this).children("I_1_Sdate").text().trim()));
                                    $("#I_1_Edate").html(transROCdate($(this).children("I_1_Edate").text().trim()));
                                    $("#I_2_Sdate").html(transROCdate($(this).children("I_2_Sdate").text().trim()));
                                    $("#I_2_Edate").html(transROCdate($(this).children("I_2_Edate").text().trim()));
                                    $("#I_3_Sdate").html(transROCdate($(this).children("I_3_Sdate").text().trim()));
                                    $("#I_3_Edate").html(transROCdate($(this).children("I_3_Edate").text().trim()));
                                    //計畫執行總月數
                                    $("#totalmonth").append(monthDiff(new Date($(this).children("I_1_Sdate").text().trim()), new Date($(this).children("I_3_Edate").text().trim())));
                                    tabstr = '<tr>';
                                    tabstr += '<td>節電基礎工作</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item1_1").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item1_2").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item1_3").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item1_all").text().trim() + '</td>';
                                    tabstr += '</tr>';
                                    tabstr += '<tr>';
                                    tabstr += '<td>因地制宜</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item2_1").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item2_2").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item2_3").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item2_all").text().trim() + '</td>';
                                    tabstr += '</tr>';
                                    tabstr += '<tr>';
                                    tabstr += '<td nowrap="nowrap">設備汰換及智慧用電</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item3_1").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item3_2").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item3_3").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item3_all").text().trim() + '</td>';
                                    tabstr += '</tr>';
                                    tabstr += '<tr>';
                                    tabstr += '<td>擴大補助</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item4_1").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item4_2").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item4_3").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Money_item4_all").text().trim() + '</td>';
                                    tabstr += '</tr>';
                                    var item1total = returnFloat($(this).children("I_Money_item1_1").text().trim()) + returnFloat($(this).children("I_Money_item2_1").text().trim()) + returnFloat($(this).children("I_Money_item3_1").text().trim()) + returnFloat($(this).children("I_Money_item4_1").text().trim());
                                    var item2total = returnFloat($(this).children("I_Money_item1_2").text().trim()) + returnFloat($(this).children("I_Money_item2_2").text().trim()) + returnFloat($(this).children("I_Money_item3_2").text().trim()) + returnFloat($(this).children("I_Money_item4_2").text().trim());
                                    var item3total = returnFloat($(this).children("I_Money_item1_3").text().trim()) + returnFloat($(this).children("I_Money_item2_3").text().trim()) + returnFloat($(this).children("I_Money_item3_3").text().trim()) + returnFloat($(this).children("I_Money_item4_3").text().trim());
                                    tabstr += '<tr class="spe">';
                                    tabstr += '<td nowrap="nowrap">總計</td>';
                                    tabstr += '<td align="right">' + item1total.toFixed(3) + '</td>';
                                    tabstr += '<td align="right">' + item2total.toFixed(3) + '</td>';
                                    tabstr += '<td align="right">' + item3total.toFixed(3) + '</td>';
                                    tabstr += '<td align="right">' + (parseFloat(item1total) + parseFloat(item2total) + parseFloat(item3total)).toFixed(3) + '</td>';
                                    tabstr += '</tr>';
                                    $("#moneyTab tbody").empty();
                                    $("#moneyTab tbody").append(tabstr);
                                    if ($(this).children("I_Other_Oneself").text().trim() == "Y") {
                                        $("#self").show();
                                        $("#I_Other_Oneself_Money").html($(this).children("I_Other_Oneself_Money").text().trim());
                                    }
                                    if ($(this).children("I_Other_Other").text().trim() == "Y") {
                                        $("#other").show();
                                        $("#I_Other_Other_name").html($(this).children("I_Other_Other_name").text().trim());
                                        $("#I_Other_Other_Money").html($(this).children("I_Other_Other_Money").text().trim());
                                    }
                                    $("#I_Target").html($(this).children("I_Target").text().trim().replace(/\n/g, "<br>"));
                                    $("#I_Summary").html($(this).children("I_Summary").text().trim().replace(/\n/g, "<br>"));

                                    
                                    tabstr = '<tr>';
                                    tabstr += '<td nowrap="nowrap">無風管冷氣(kW)<br />註：每台冷氣約4kW</td >';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item1_1").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item1_2").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item1_3").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item1_all").text().trim() + '</td>';
                                    tabstr += '</tr>';
                                    tabstr += '<tr>';
                                    tabstr += '<td nowrap="nowrap">老舊辦公室照明(具)</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item2_1").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item2_2").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item2_3").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item2_all").text().trim() + '</td>';
                                    tabstr += '</tr>';
                                    tabstr += '<tr>';
                                    tabstr += '<td nowrap="nowrap">室內停車場智慧照明(盞)</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item3_1").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item3_2").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item3_3").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item3_all").text().trim() + '</td>';
                                    tabstr += '</tr>';
                                    tabstr += '<tr>';
                                    tabstr += '<td nowrap="nowrap">中型能管系統(套)</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item4_1").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item4_2").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item4_3").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item4_all").text().trim() + '</td>';
                                    tabstr += '</tr>';
                                    tabstr += '<tr>';
                                    tabstr += '<td nowrap="nowrap">大型能管系統(套)</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item5_1").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item5_2").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item5_3").text().trim() + '</td>';
                                    tabstr += '<td align="right">' + $(this).children("I_Finish_item5_all").text().trim() + '</td>';
                                    tabstr += '</tr>';
                                    $("#finishTab tbody").empty();
                                    $("#finishTab tbody").append(tabstr);

                                    //檢查未填資料
                                    if ($(this).children("I_1_Sdate").text().trim() == "" ||
                                        $(this).children("I_1_Edate").text().trim() == "" ||
                                        $(this).children("I_2_Sdate").text().trim() == "" ||
                                        $(this).children("I_2_Edate").text().trim() == "" ||
                                        $(this).children("I_3_Sdate").text().trim() == "" ||
                                        $(this).children("I_3_Edate").text().trim() == "" ||
                                        $(this).children("I_Target").text().trim() == "" ||
                                        $(this).children("I_Summary").text().trim() == ""
                                    ) {
                                        $("#nullstatus").val("【基本資料】未填寫完整，請再確認");
                                    }

                                    if ($(this).children("I_Flag").text().trim() == "Y")
                                        $("#exbtn").show();
                                });
                            }

                            // 擴大項目預計完成數
                            if ($(data).find("ex_item").length > 0) {
                                if ($(data).find("ex_item[P_Period='1']").length > 0) {
                                    tabstr = '<tr>';
                                    tabstr += '<th nowrap="nowrap">項目</th>';
                                    tabstr += '<th nowrap="nowrap">第1期</th>';
                                    tabstr += '</tr>';
                                    $(data).find("ex_item[P_Period='1']").each(function (i) {
                                        tabstr += '<tr>';
                                        tabstr += '<td nowrap="nowrap">' + $(this).attr("P_ItemName") + '</td>';
                                        tabstr += '<td align="right">' + $(this).attr("P_ExFinish") + '</td>';
                                        tabstr += '</tr>';
                                    });
                                }
                                if ($(data).find("ex_item[P_Period='2']").length > 0) {
                                    tabstr += '<tr>';
                                    tabstr += '<th nowrap="nowrap">項目</th>';
                                    tabstr += '<th nowrap="nowrap">第2期</th>';
                                    tabstr += '</tr>';
                                    $(data).find("ex_item[P_Period='2']").each(function (i) {
                                        tabstr += '<tr>';
                                        tabstr += '<td nowrap="nowrap">' + $(this).attr("P_ItemName") + '</td>';
                                        tabstr += '<td align="right">' + $(this).attr("P_ExFinish") + '</td>';
                                        tabstr += '</tr>';
                                    });
                                }
                                if ($(data).find("ex_item[P_Period='3']").length > 0) {
                                    tabstr += '<tr>';
                                    tabstr += '<th nowrap="nowrap">項目</th>';
                                    tabstr += '<th nowrap="nowrap">第3期</th>';
                                    tabstr += '</tr>';
                                    $(data).find("ex_item[P_Period='3']").each(function (i) {
                                        tabstr += '<tr>';
                                        tabstr += '<td nowrap="nowrap">' + $(this).attr("P_ItemName") + '</td>';
                                        tabstr += '<td align="right">' + $(this).attr("P_ExFinish") + '</td>';
                                        tabstr += '</tr>';
                                    });
                                }
                            }
                            else
                                tabstr = '<tr><td>查詢無資料</td></tr>';
                            $("#exFinishTab").empty();
                            $("#exFinishTab").append(tabstr);

                            //承辦人
                            if ($(data).find("mb_item").length > 0) {
                                $(data).find("mb_item").each(function (i) {
                                    $("#mb_Name").html($(this).children("M_Name").text().trim());
                                    $("#mb_JobTitle").html($(this).children("M_JobTitle").text().trim());
                                    $("#mb_Tel").html($(this).children("M_Tel").text().trim());
                                    $("#mb_Phone").html($(this).children("M_Phone").text().trim());
                                    $("#mb_Fax").html($(this).children("M_Fax").text().trim());
                                    $("#mb_Email").html($(this).children("M_Email").text().trim());
                                    $("#mb_Addr").html($(this).children("M_Addr").text().trim());

                                });
                            }

                            //承辦主管
                            if ($(data).find("mm_item").length > 0) {
                                $(data).find("mm_item").each(function (i) {
                                    $("#mm_Name").html($(this).children("M_Name").text().trim());
                                    $("#mm_JobTitle").html($(this).children("M_JobTitle").text().trim());
                                    $("#mm_Tel").html($(this).children("M_Tel").text().trim());
                                    $("#mm_Phone").html($(this).children("M_Phone").text().trim());
                                    $("#mm_Fax").html($(this).children("M_Fax").text().trim());
                                    $("#mm_Email").html($(this).children("M_Email").text().trim());
                                    $("#mm_Addr").html($(this).children("M_Addr").text().trim());
                                });
                            }

                        }
                    }
                }
            });
        }

        function getCheckPoint(tp) {
            $.ajax({
                type: "POST",
                async: true, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getCheckPoint.aspx",
                data: {
                    period: $("#cp_period").val(),
                    type: tp,
                    person_id: $.getParamValue('v')
                },
                beforeSend: function () {
                    $('body').loading({ theme: 'dark' });
                },
                complete: function () {
                    if (tp == "04")
                        $('body').loading("stop");
                },
                error: function (xhr) {
                    alert(xhr.responseText);
                },
                success: function (data) {
                    if ($(data).find("Error").length > 0) {
                        alert($(data).find("Error").attr("Message"));
                    }
                    else {
                        var TagName = "";
                        switch (tp) {
                            case "01":
                                TagName = "#basicworkTab";
                                InputCode = "";
                                break;
                            case "02":
                                TagName = "#placeTab";
                                InputCode = "p_";
                                break;
                            case "03":
                                TagName = "#smartTab";
                                InputCode = "s_";
                                break;
                            case "04":
                                TagName = "#allowanceTab";
                                InputCode = "a_";
                                break;
                        }

                        $(TagName + " tbody").empty();
                        var tabstr = '';
                        var finishStr = '';
                        if ($(data).find("PushItem[P_Type='" + tp + "']").length > 0) {
                            $(data).find("PushItem[P_Type='" + tp + "']").each(function (i) {
                                //First row
                                tabstr += '<tr>';
                                tabstr += '<td align="center" nowrap="nowrap" rowspan="' + $(this).children().length + '">';
                                switch (tp) {
                                    default:
                                        tabstr += $(this).attr("P_ItemName");
                                        break;
                                    case "03":
                                        tabstr += getCP_TypeCn($(this).attr("P_ItemNameCode"));
                                        break;
                                    case "04":
                                        tabstr += getExTypeCn($(this).attr("P_ItemNameCode"));
                                        break;
                                }
                                tabstr += '</td>';
                                tabstr += '<td align="center">' + $(this).children().find("CP_Point")[0].textContent + '</td>';
                                tabstr += '<td align="center">' + $(this).children().find("CP_ReserveYear")[0].textContent + ' 年 ' + $(this).children().find("CP_ReserveMonth")[0].textContent + ' 月</td>';
                                tabstr += '<td>' + $(this).children().find("CP_Desc")[0].textContent + '</td>';
                                tabstr += '</tr>';
                                $(this).children().each(function (i) {
                                    // 跳過第 1 筆 Row
                                    if (i != 0) {
                                        tabstr += '<tr>';
                                        tabstr += '<td align="center">' + $(this).children("CP_Point").text().trim() + '</td>';
                                        tabstr += '<td align="center">' + $(this).children("CP_ReserveYear").text().trim() + ' 年 ' + $(this).children("CP_ReserveMonth").text().trim() + ' 月</td>';
                                        tabstr += '<td>' + $(this).children("CP_Desc").text().trim() + '</td>';
                                        tabstr += '</tr>';
                                    }
                                });
                            });
                            $(TagName).append(tabstr);
                        }
                        else {
                            $(TagName).append('<tr><td colspan="4">查詢無資料</td></tr>');
                        }
                    }
                }
            });
        }

        function getProgressData() {
            $.ajax({
                type: "POST",
                async: true, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getProgress.aspx",
                data: {
                    period: $("#p_period").val(),
                    mid: $.getParamValue('v')
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                beforeSend: function (data) {
                    $("#bwTabLoad").show();
                    $("#pTabLoad").show();
                    $("#sTabLoad").show();
                    $("#bwTab").empty();
                    $("#pTab").empty();
                    $("#sTab").empty();
                    $("#exTab").empty();
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

                        if (data != null) {
                            //data = $.parseXML(data);
                            $("#bwTab").append("<tr><td>查詢無資料</td></tr>");
                            $("#pTab").append("<tr><td>查詢無資料</td></tr>");
                            $("#sTab").append("<tr><td>查詢無資料</td></tr>");
                            $("#exTab").append("<tr><td>查詢無資料</td></tr>");

                            var monthstr = '';
                            var monthAry = [];
                            var tmpI = 1;

                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function (i) {
                                    switch ($(this).children("P_Type").text().trim()) {
                                        //節電基礎工作
                                        case "01":
                                            //年月丟到陣列
                                            var mstr = $(this).children("CP_ReserveYear").text().trim() + '-' + $(this).children("CP_ReserveMonth").text().trim();
                                            if ($.inArray(mstr, monthAry) == -1)
                                                monthAry.push(mstr);

                                            //判斷是否為最後一筆row
                                            if ($(this).children("TypeTotal").text().trim() == tmpI) {
                                                var monthInYearNum = 0;
                                                //表頭月份
                                                var tmp1 = "";
                                                var xmlstr = "";
                                                monthAry.sort();
                                                for (var j = 0; j < monthAry.length; j++) {
                                                    var splitstr = monthAry[j].split('-');
                                                    //月份
                                                    var mtmp = (splitstr[1].substr(0, 1) == "0") ? splitstr[1].substr(1, 1) : splitstr[1];
                                                    monthstr += '<th nowrap="nowrap">' + mtmp + '月</th>';

                                                    //湊表頭年份的XML
                                                    if (splitstr[0] == tmp1 || j == 0) {
                                                        monthInYearNum += 1;
                                                        if (j == (monthAry.length - 1))
                                                            xmlstr += '<data year="' + splitstr[0] + '" cspan="' + monthInYearNum + '" />';
                                                    }
                                                    else {
                                                        xmlstr += '<data year="' + tmp1 + '" cspan="' + monthInYearNum + '" />';
                                                        monthInYearNum = 1;
                                                        if (j == (monthAry.length - 1))
                                                            xmlstr += '<data year="' + splitstr[0] + '" cspan="' + monthInYearNum + '" />';
                                                    }
                                                    //set tmp1
                                                    tmp1 = splitstr[0];
                                                }
                                                xmlstr = '<root>' + xmlstr + '</root>';

                                                //表頭年份
                                                var tabColumnCount = 0;
                                                xmlstr = $.parseXML(xmlstr);
                                                var yearstr = '<tr>';
                                                $(xmlstr).find("data").each(function (i) {
                                                    yearstr += '<th colspan="' + $(this).attr("cspan") + '">' + $(this).attr("year") + '年</th>';
                                                    tabColumnCount += parseInt($(this).attr("cspan"));
                                                });
                                                yearstr += '</tr>';

                                                var tabdetail = TableDetail(data, "data01", monthAry);

                                                var startday = "";
                                                var endday = "";
                                                $(data).find("priddate").each(function (i) {
                                                    //起始日
                                                    var year = new Date($(this).children("I_" + $("#p_period").val() + "_Sdate").text().trim()).getFullYear() - 1911;
                                                    var month = new Date($(this).children("I_" + $("#p_period").val() + "_Sdate").text().trim()).getMonth() + 1;
                                                    var day = new Date($(this).children("I_" + $("#p_period").val() + "_Sdate").text().trim()).getDate();
                                                    startday = year + "年" + month + "月" + day + "日";
                                                    //結束日
                                                    year = new Date($(this).children("I_" + $("#p_period").val() + "_Edate").text().trim()).getFullYear() - 1911;
                                                    month = new Date($(this).children("I_" + $("#p_period").val() + "_Edate").text().trim()).getMonth() + 1;
                                                    day = new Date($(this).children("I_" + $("#p_period").val() + "_Edate").text().trim()).getDate();
                                                    endday = year + "年" + month + "月" + day + "日";
                                                });

                                                var tabstr = '<tr>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">推動項目</th>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">工作比重(%)</th>';
                                                tabstr += '<th rowspan="3">年&nbsp;&nbsp;&nbsp;月<br>進度(%)</th>';
                                                tabstr += '<th colspan="' + monthAry.length + '">' + startday + '至' + endday + '</th></tr >';
                                                monthstr = '<tr>' + monthstr + '</tr>';
                                                tabstr += yearstr + monthstr + tabdetail;
                                                $("#bwTab").empty();
                                                $("#bwTab").append(tabstr);

                                                //各月份合計(湊完table再抓name做處理)
                                                monthTotal($(this).children("P_Type").text().trim(), monthAry.length);

                                                //清空全域變數
                                                monthstr = "";
                                                monthAry = [];
                                                tmpI = 0;
                                            }
                                            tmpI += 1;
                                            break;
                                        //因地制宜
                                        case "02":
                                            //年月丟到陣列
                                            var mstr = $(this).children("CP_ReserveYear").text().trim() + '-' + $(this).children("CP_ReserveMonth").text().trim();
                                            if ($.inArray(mstr, monthAry) == -1)
                                                monthAry.push(mstr);
                                            //判斷是否為最後一筆row
                                            if ($(this).children("TypeTotal").text().trim() == tmpI) {
                                                var monthInYearNum = 0;
                                                //表頭月份
                                                var tmp1 = "";
                                                var xmlstr = "";
                                                monthAry.sort();
                                                for (var j = 0; j < monthAry.length; j++) {
                                                    var splitstr = monthAry[j].split('-');
                                                    //月份
                                                    var mtmp = (splitstr[1].substr(0, 1) == "0") ? splitstr[1].substr(1, 1) : splitstr[1];
                                                    monthstr += '<th nowrap="nowrap">' + mtmp + '月</th>';

                                                    //湊表頭年份的XML
                                                    if (splitstr[0] == tmp1 || j == 0) {
                                                        monthInYearNum += 1;
                                                        if (j == (monthAry.length - 1))
                                                            xmlstr += '<data year="' + splitstr[0] + '" cspan="' + monthInYearNum + '" />';
                                                    }
                                                    else {
                                                        xmlstr += '<data year="' + tmp1 + '" cspan="' + monthInYearNum + '" />';
                                                        monthInYearNum = 1;
                                                        if (j == (monthAry.length - 1))
                                                            xmlstr += '<data year="' + splitstr[0] + '" cspan="' + monthInYearNum + '" />';
                                                    }
                                                    //set tmp1
                                                    tmp1 = splitstr[0];
                                                }
                                                xmlstr = '<root>' + xmlstr + '</root>';

                                                //表頭年份
                                                xmlstr = $.parseXML(xmlstr);
                                                var yearstr = '<tr>';
                                                $(xmlstr).find("data").each(function (i) {
                                                    yearstr += '<th colspan="' + $(this).attr("cspan") + '">' + $(this).attr("year") + '年</th>';
                                                });
                                                yearstr += '</tr>';

                                                var tabdetail = TableDetail(data, "data02", monthAry);

                                                var startday = "";
                                                var endday = "";
                                                $(data).find("priddate").each(function (i) {
                                                    //起始日
                                                    var year = new Date($(this).children("I_" + $("#p_period").val() + "_Sdate").text().trim()).getFullYear() - 1911;
                                                    var month = new Date($(this).children("I_" + $("#p_period").val() + "_Sdate").text().trim()).getMonth() + 1;
                                                    var day = new Date($(this).children("I_" + $("#p_period").val() + "_Sdate").text().trim()).getDate();
                                                    startday = year + "年" + month + "月" + day + "日";
                                                    //結束日
                                                    year = new Date($(this).children("I_" + $("#p_period").val() + "_Edate").text().trim()).getFullYear() - 1911;
                                                    month = new Date($(this).children("I_" + $("#p_period").val() + "_Edate").text().trim()).getMonth() + 1;
                                                    day = new Date($(this).children("I_" + $("#p_period").val() + "_Edate").text().trim()).getDate();
                                                    endday = year + "年" + month + "月" + day + "日";
                                                });

                                                var tabstr = '<tr>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">推動項目</th>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">工作比重(%)</th>';
                                                tabstr += '<th rowspan="3">年&nbsp;&nbsp;&nbsp;月<br>進度(%)</th>';
                                                tabstr += '<th colspan="' + monthAry.length + '">' + startday + '至' + endday + '</th></tr >';
                                                tabstr += yearstr + monthstr + tabdetail;
                                                $("#pTab").empty();
                                                $("#pTab").append(tabstr);

                                                //各月份合計(湊完table再抓name做處理)
                                                monthTotal($(this).children("P_Type").text().trim(), monthAry.length);

                                                //清空全域變數
                                                monthstr = "";
                                                monthAry = [];
                                                tmpI = 0;
                                            }
                                            tmpI += 1;
                                            break;
                                        //設備汰換及智慧用電
                                        case "03":
                                            //年月丟到陣列
                                            var mstr = $(this).children("CP_ReserveYear").text().trim() + '-' + $(this).children("CP_ReserveMonth").text().trim();
                                            if ($.inArray(mstr, monthAry) == -1)
                                                monthAry.push(mstr);
                                            //判斷是否為最後一筆row
                                            if ($(this).children("TypeTotal").text().trim() == tmpI) {
                                                var monthInYearNum = 0;
                                                //表頭月份
                                                var tmp1 = "";
                                                var xmlstr = "";
                                                monthAry.sort();
                                                for (var j = 0; j < monthAry.length; j++) {
                                                    var splitstr = monthAry[j].split('-');
                                                    //月份
                                                    var mtmp = (splitstr[1].substr(0, 1) == "0") ? splitstr[1].substr(1, 1) : splitstr[1];
                                                    monthstr += '<th nowrap="nowrap">' + mtmp + '月</th>';

                                                    //湊表頭年份的XML
                                                    if (splitstr[0] == tmp1 || j == 0) {
                                                        monthInYearNum += 1;
                                                        if (j == (monthAry.length - 1))
                                                            xmlstr += '<data year="' + splitstr[0] + '" cspan="' + monthInYearNum + '" />';
                                                    }
                                                    else {
                                                        xmlstr += '<data year="' + tmp1 + '" cspan="' + monthInYearNum + '" />';
                                                        monthInYearNum = 1;
                                                        if (j == (monthAry.length - 1))
                                                            xmlstr += '<data year="' + splitstr[0] + '" cspan="' + monthInYearNum + '" />';
                                                    }
                                                    //set tmp1
                                                    tmp1 = splitstr[0];
                                                }
                                                xmlstr = '<root>' + xmlstr + '</root>';

                                                //表頭年份
                                                xmlstr = $.parseXML(xmlstr);
                                                var yearstr = '<tr>';
                                                $(xmlstr).find("data").each(function (i) {
                                                    yearstr += '<th colspan="' + $(this).attr("cspan") + '">' + $(this).attr("year") + '年</th>';
                                                });
                                                yearstr += '</tr>';

                                                var tabdetail = TableDetail(data, "data03", monthAry);

                                                var startday = "";
                                                var endday = "";
                                                $(data).find("priddate").each(function (i) {
                                                    //起始日
                                                    var year = new Date($(this).children("I_" + $("#p_period").val() + "_Sdate").text().trim()).getFullYear() - 1911;
                                                    var month = new Date($(this).children("I_" + $("#p_period").val() + "_Sdate").text().trim()).getMonth() + 1;
                                                    var day = new Date($(this).children("I_" + $("#p_period").val() + "_Sdate").text().trim()).getDate();
                                                    startday = year + "年" + month + "月" + day + "日";
                                                    //結束日
                                                    year = new Date($(this).children("I_" + $("#p_period").val() + "_Edate").text().trim()).getFullYear() - 1911;
                                                    month = new Date($(this).children("I_" + $("#p_period").val() + "_Edate").text().trim()).getMonth() + 1;
                                                    day = new Date($(this).children("I_" + $("#p_period").val() + "_Edate").text().trim()).getDate();
                                                    endday = year + "年" + month + "月" + day + "日";
                                                });

                                                var tabstr = '<tr>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">推動項目</th>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">工作比重(%)</th>';
                                                tabstr += '<th rowspan="3">年&nbsp;&nbsp;&nbsp;月<br>進度(%)</th>';
                                                tabstr += '<th colspan="' + monthAry.length + '">' + startday + '至' + endday + '</th></tr >';
                                                tabstr += yearstr + monthstr + tabdetail;
                                                $("#sTab").empty();
                                                $("#sTab").append(tabstr);

                                                //各月份合計(湊完table再抓name做處理)
                                                monthTotal($(this).children("P_Type").text().trim(), monthAry.length);

                                                //清空全域變數
                                                monthstr = "";
                                                monthAry = [];
                                                tmpI = 0;
                                            }
                                            tmpI += 1;
                                            break;
                                        //擴大補助
                                        case "04":
                                            //年月丟到陣列
                                            var mstr = $(this).children("CP_ReserveYear").text().trim() + '-' + $(this).children("CP_ReserveMonth").text().trim();
                                            if ($.inArray(mstr, monthAry) == -1)
                                                monthAry.push(mstr);
                                            //判斷是否為最後一筆row
                                            if ($(this).children("TypeTotal").text().trim() == tmpI) {
                                                var monthInYearNum = 0;
                                                //表頭月份
                                                var tmp1 = "";
                                                var xmlstr = "";
                                                monthAry.sort();
                                                for (var j = 0; j < monthAry.length; j++) {
                                                    var splitstr = monthAry[j].split('-');
                                                    //月份
                                                    var mtmp = (splitstr[1].substr(0, 1) == "0") ? splitstr[1].substr(1, 1) : splitstr[1];
                                                    monthstr += '<th nowrap="nowrap">' + mtmp + '月</th>';

                                                    //湊表頭年份的XML
                                                    if (splitstr[0] == tmp1 || j == 0) {
                                                        monthInYearNum += 1;
                                                        if (j == (monthAry.length - 1))
                                                            xmlstr += '<data year="' + splitstr[0] + '" cspan="' + monthInYearNum + '" />';
                                                    }
                                                    else {
                                                        xmlstr += '<data year="' + tmp1 + '" cspan="' + monthInYearNum + '" />';
                                                        monthInYearNum = 1;
                                                        if (j == (monthAry.length - 1))
                                                            xmlstr += '<data year="' + splitstr[0] + '" cspan="' + monthInYearNum + '" />';
                                                    }
                                                    //set tmp1
                                                    tmp1 = splitstr[0];
                                                }
                                                xmlstr = '<root>' + xmlstr + '</root>';

                                                //表頭年份
                                                xmlstr = $.parseXML(xmlstr);
                                                var yearstr = '<tr>';
                                                $(xmlstr).find("data").each(function (i) {
                                                    yearstr += '<th colspan="' + $(this).attr("cspan") + '">' + $(this).attr("year") + '年</th>';
                                                });
                                                yearstr += '</tr>';

                                                var tabdetail = TableDetail(data, "data04", monthAry);

                                                var startday = "";
                                                var endday = "";
                                                $(data).find("priddate").each(function (i) {
                                                    //起始日
                                                    var year = new Date($(this).children("I_" + $("#p_period").val() + "_Sdate").text().trim()).getFullYear() - 1911;
                                                    var month = new Date($(this).children("I_" + $("#p_period").val() + "_Sdate").text().trim()).getMonth() + 1;
                                                    var day = new Date($(this).children("I_" + $("#p_period").val() + "_Sdate").text().trim()).getDate();
                                                    startday = year + "年" + month + "月" + day + "日";
                                                    //結束日
                                                    year = new Date($(this).children("I_" + $("#p_period").val() + "_Edate").text().trim()).getFullYear() - 1911;
                                                    month = new Date($(this).children("I_" + $("#p_period").val() + "_Edate").text().trim()).getMonth() + 1;
                                                    day = new Date($(this).children("I_" + $("#p_period").val() + "_Edate").text().trim()).getDate();
                                                    endday = year + "年" + month + "月" + day + "日";
                                                });

                                                var tabstr = '<tr>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">推動項目</th>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">工作比重(%)</th>';
                                                tabstr += '<th rowspan="3">年&nbsp;&nbsp;&nbsp;月<br>進度(%)</th>';
                                                tabstr += '<th colspan="' + monthAry.length + '">' + startday + '至' + endday + '</th></tr >';
                                                tabstr += yearstr + monthstr + tabdetail;
                                                $("#exTab").empty();
                                                $("#exTab").append(tabstr);

                                                //各月份合計(湊完table再抓name做處理)
                                                monthTotal($(this).children("P_Type").text().trim(), monthAry.length);

                                                //清空全域變數
                                                monthstr = "";
                                                monthAry = [];
                                                tmpI = 0;
                                            }
                                            tmpI += 1;
                                            break;
                                    }
                                });
                            }
                            $("#bwTabLoad").hide();
                            $("#pTabLoad").hide();
                            $("#sTabLoad").hide();
                            $("#exTabLoad").hide();

                        }//data != null end
                    }
                }
            });//ajax end
        }

        function TableDetail(xmlstr, nodename, monthAry) {
            var rnum = 0; //計算同推動項目群組目前是第幾筆
            var tabrow1 = ''; //Row 查核點
            var tabrow2 = ''; //Row 累計預定進度
            var tabrow3 = ''; //Row 查核點進度說明
            var rowstr = '';
            var tmpid = "";
            var typenum = nodename.substr(nodename.length - 2, 2); //Type 代碼
            var wr = { tmp: 0, sum01: 0, sum02: 0, sum03: 0, sum04: 0 } //工作比重
            $(xmlstr).find(nodename).each(function (k) {
                if (k == 0)
                    tmpid = $(this).children("CP_ParentId").text().trim();
                //推動項目群組第一筆進來時
                if ($(this).children("CP_ParentId").text().trim() != tmpid || k == 0) {
                    //跳下一個推動項目群組時
                    if (k != 0) {
                        //每筆row的後面未填滿時要補滿
                        if (rnum != monthAry.length) {
                            var mRange = monthAry.length - rnum
                            for (var a = 0; a < mRange; a++) {
                                tabrow1 += '<td valign="top"></td>';
                                tabrow2 += '<td></td>';
                            }
                        }
                        //結尾
                        tabrow1 += '</tr>';
                        tabrow2 += '</tr>';
                        tabrow3 += '</td></tr>';
                        rowstr += tabrow1 + tabrow2 + tabrow3;
                        //清空暫存
                        tabrow1 = "";
                        tabrow2 = "";
                        tabrow3 = "";
                        rnum = 0;
                    }
                    //工作比重加總
                    var wRatioVal = ($(this).children("P_WorkRatio").text().trim() != "") ? $(this).children("P_WorkRatio").text().trim() : 0;
                    wr.tmp += parseFloat(wRatioVal);
                    //檢查是否有未填資料
                    if ($(this).children("P_WorkRatio").text().trim() == "")
                        $("#nullstatus").val("【工作比重】未填寫完整，請再確認");
                    //第一筆開頭
                    var tmpiname = "";
                    switch (typenum) {
                        default:
                            tmpiname = $(this).children("P_ItemName").text().trim();
                            break;
                        case "03":
                            tmpiname = getCP_TypeCn($(this).children("P_ItemName").text().trim());
                            if ($(this).children("P_ItemName").text().trim() == "99")
                                tmpiname += "-" + $(this).children("CP_Desc").text().trim();
                            break;
                        case "04":
                            tmpiname = getExTypeCn($(this).children("P_ItemName").text().trim());
                            if ($(this).children("P_ItemName").text().trim() == "99")
                                tmpiname += "-" + $(this).children("CP_Desc").text().trim();
                            break;
                    }
                    tabrow1 += '<tr><td rowspan="3" align="center"><strong>' + tmpiname + '</strong></td>';
                    var tmpWorkRatio = ($(this).children("P_WorkRatio").text().trim() != "") ? $(this).children("P_WorkRatio").text().trim() : "0";
                    tabrow1 += '<td rowspan="3" align="center">' + tmpWorkRatio + '%</td><td align="center">查核點</td>';
                    tabrow2 += '<tr><td align="center">累計預定進度(%)</td >';
                    tabrow3 += '<tr><td nowrap="nowrap" align="center">查核點<br />進度說明</td><td colspan="' + monthAry.length + '">';
                    //因為資料庫與Table的X、Y軸相反
                    //為了好寫用上面整理好的月份陣列跑迴圈轉成同方向
                    for (var l = 0; l < monthAry.length; l++) {
                        var splitstr2 = monthAry[l].split('-');
                        //當對應到年跟月時
                        if ($(this).children("CP_ReserveYear").text().trim() == splitstr2[0] && $(this).children("CP_ReserveMonth").text().trim() == splitstr2[1]) {
                            tabrow1 += '<td valign="top" align="center">' + $(this).children("CP_Point").text().trim() + '</td>';
                            var tmpProc = ($(this).children("CP_Process").text().trim() == "") ? 0 : $(this).children("CP_Process").text().trim();
                            tabrow2 += '<td align="center"><span class="cpp" pv="pv' + typenum + l + '" item="' + $(this).children("CP_ParentId").text().trim().substr(0, 10) +'">' + tmpProc + '</span>%</td>';
                            tabrow3 += $(this).children("CP_Point").text().trim() + '&nbsp;&nbsp;' + $(this).children("CP_Desc").text().trim() + '<br />';
                            rnum += 1;
                            //檢查是否有未填資料
                            if ($(this).children("CP_Process").text().trim() == "")
                                $("#nullstatus").val("【累計預定進度】未填寫完整，請再確認");
                            break;
                        }
                        //沒對應到就補空值
                        else {
                            tabrow1 += '<td valign="top"></td>';
                            tabrow2 += '<td></td>';
                            rnum += 1;
                        }
                    }

                    //如果為最後一筆資料
                    if ($(xmlstr).find(nodename).length == (k + 1)) {
                        //每筆row的後面未填滿時要補滿
                        if (rnum != monthAry.length) {
                            var mRange = monthAry.length - rnum
                            for (var a = 0; a < mRange; a++) {
                                tabrow1 += '<td valign="top"></td>';
                                tabrow2 += '<td></td>';
                            }
                        }
                        //結尾
                        tabrow1 += '</tr>';
                        tabrow2 += '</tr>';
                        tabrow3 += '</td></tr>';
                        rowstr += tabrow1 + tabrow2 + tabrow3;

                        switch (typenum) {
                            case "01":
                                wr.sum01 = wr.tmp;
                                break;
                            case "02":
                                wr.sum02 = wr.tmp;
                                break;
                            case "03":
                                wr.sum03 = wr.tmp;
                                break;
                            case "04":
                                wr.sum04 = wr.tmp;
                                break;
                        }
                    }
                }
                //推動項目群組非第一筆進來時
                else {
                    for (var l = 0; l < monthAry.length; l++) {
                        var splitstr2 = monthAry[l].split('-');
                        //當對應到年跟月時
                        if ($(this).children("CP_ReserveYear").text().trim() == splitstr2[0] && $(this).children("CP_ReserveMonth").text().trim() == splitstr2[1]) {
                            tabrow1 += '<td valign="top" align="center">' + $(this).children("CP_Point").text().trim() + '</td>';
                            var tmpProc = ($(this).children("CP_Process").text().trim() == "") ? 0 : $(this).children("CP_Process").text().trim();
                            tabrow2 += '<td align="center"><span class="cpp" pv="pv' + typenum + l + '" item="' + $(this).children("CP_ParentId").text().trim().substr(0, 10) +'">' + tmpProc+ '</span>%</td>';
                            tabrow3 += $(this).children("CP_Point").text().trim() + '&nbsp;&nbsp;' + $(this).children("CP_Desc").text().trim() + '<br />';
                            rnum += 1;
                            //檢查是否有未填資料
                            if ($(this).children("CP_Process").text().trim() == "")
                                $("#nullstatus").val("【累計預定進度】未填寫完整，請再確認");
                            break;
                        }

                        //判斷是否中間有跳空的月份
                        if (rnum == l) {
                            tabrow1 += '<td valign="top"></td>';
                            tabrow2 += '<td></td>';
                            rnum += 1;
                        }
                    }

                    //如果為最後一筆資料
                    if ($(xmlstr).find(nodename).length == (k + 1)) {
                        //每筆row的後面未填滿時要補滿
                        if (rnum != monthAry.length) {
                            var mRange = monthAry.length - rnum
                            for (var a = 0; a < mRange; a++) {
                                tabrow1 += '<td valign="top"></td>';
                                tabrow2 += '<td></td>';
                            }
                        }
                        //結尾
                        tabrow1 += '</tr>';
                        tabrow2 += '</tr>';
                        tabrow3 += '</td></tr>';
                        rowstr += tabrow1 + tabrow2 + tabrow3;

                        switch (typenum) {
                            case "01":
                                wr.sum01 = wr.tmp;
                                break;
                            case "02":
                                wr.sum02 = wr.tmp;
                                break;
                            case "03":
                                wr.sum03 = wr.tmp;
                                break;
                            case "04":
                                wr.sum04 = wr.tmp;
                                break;
                        }
                    }
                }
                tmpid = $(this).children("CP_ParentId").text().trim();
            });
            //最下方合計
            rowstr += '<tr class="spe">';
            rowstr += '<td style="text-align:center;">合　計 </td>';
            //工作比重
            switch (typenum) {
                case "01":
                    rowstr += '<td align="center"><span id="wrTotal' + typenum + '">' + Number(wr.sum01.toFixed(2)) + '</span>%</td>';
                    break;
                case "02":
                    rowstr += '<td align="center"><span id="wrTotal' + typenum + '">' + Number(wr.sum02.toFixed(2)) + '</span>%</td>';
                    break;
                case "03":
                    rowstr += '<td align="center"><span id="wrTotal' + typenum + '">' + Number(wr.sum03.toFixed(2)) + '</span>%</td>';
                    break;
                case "04":
                    rowstr += '<td align="center"><span id="wrTotal' + typenum + '">' + Number(wr.sum04.toFixed(2)) + '</span>%</td>';
                    break;
            }

            //累計預定進度
            rowstr += '<td nowrap="nowrap" align="center">累計預定進度(%)</td >';
            for (var i = 0; i < monthAry.length; i++) {
                rowstr += '<td align="center"><span id="cppTotal' + typenum + i + '">0</span>%</td>';
            }
            rowstr += '</tr>';
            return rowstr;
        }

        function monthTotal(tp, mCount) {
            var itemAry = [];
            var valAry = [];
            for (var b = 0; b < mCount; b++) {
                var pvtmp = "pv" + tp + b;
                var sum = 0;
                $(".cpp").each(function () {
                    if ($(this).attr("pv") == pvtmp) {
                        var v = ($(this).html() == "") ? 0 : $(this).html();
                        //累計方式是取每筆推動項目該月份向前抓最後一個有資料的月份
                        if ($.inArray($(this).attr("item"), itemAry) == -1) {
                            itemAry.push($(this).attr("item"));
                            valAry.push(v);
                        }
                        else {
                            valAry.splice($.inArray($(this).attr("item"), itemAry), 1, v);
                        }
                    }
                });

                $.each(valAry, function (i) {
                    sum += parseFloat(valAry[i]);
                });
                $("#cppTotal" + tp + b).html(Number(sum.toFixed(2)));
            }
        }

        function transROCdate(day) {
            if (day == "") return;
            var tmpdate = new Date(day);
            var y = tmpdate.getFullYear() - 1911;
            var m = tmpdate.getMonth() + 1;
            var d = tmpdate.getDate();
            tmpdate = y + " 年 " + m + " 月 " + d + " 日";
            return tmpdate;
        }

        function returnFloat(v) {
            if (v == "")
                return 0;
            else
                return parseFloat(v);
        }

        function getCP_TypeCn(tp) {
            var str = "";
            switch (tp) {
                case "01":
                    str = "無風管空氣調節機";
                    break;
                case "02":
                    str = "老舊辦公室照明";
                    break;
                case "03":
                    str = "室內停車場智慧照明";
                    break;
                case "04":
                    str = "中型服務業、機關及學校導入能源管理系統";
                    break;
                case "05":
                    str = "大型服務業、機關及學校導入能源管理系統";
                    break;
                case "99":
                    str = "其他";
                    break;
            }
            return str;
        }

         function getExTypeCn(tp) {
            var str = "";
            switch (tp) {
                case "01":
                    str = "(接)風管空氣調節機(4KW/台)";
                    break;
                case "02":
                    str = "老舊辦公室照明燈具(T5螢光燈具)";
                    break;
                case "03":
                    str = "室內停車照明";
                    break;
                case "04":
                    str = "51KW以下能管系統";
                    break;
                case "05":
                    str = "家庭(冷氣670度/3.2KW)";
                    break;
                case "06":
                    str = "電冰箱";
                    break;
                case "07":
                    str = "電視";
                    break;
                case "08":
                    str = "電熱水瓶";
                    break;
                case "09":
                    str = "電熱水器(儲備型)";
                    break;
                case "10":
                    str = "電子鍋";
                    break;
                case "11":
                    str = "溫熱型開飲機";
                    break;
                case "12":
                    str = "電鍋";
                    break;
                case "13":
                    str = "吹風機";
                    break;
                case "14":
                    str = "公設LED照明(10顆)";
                    break;
                case "99":
                    str = "其他";
                    break;
            }
            return str;
        }

        //計算幾個月
        function monthDiff(d1, d2) {
            var months;
            months = (d2.getFullYear() - d1.getFullYear()) * 12;
            //months -= d1.getMonth() + 1;
            months -= d1.getMonth() - 1;
            months += d2.getMonth();
            //2017/01/01~2017/01/23 這樣要算1個月
            //2017/01/01~2017/02/12 這樣要算2個月
            return months <= 0 ? 1 : months;
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <input type="hidden" id="nullstatus" value="" />
    <div id="content" class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">計畫書基本資料</span></div>
        </div>
        <!--計畫基本資料-->
        <div class="OchiTrasTable width100 TitleLength08">
            <!-- 雙欄 -->
            <div class="OchiRow">
                <div class="OchiHalf " id="div_outer_city">
                    <div class="OchiCell OchiTitle TitleSetWidth">執行機關</div>
                    <div class="OchiCell width100">
                        <div class="OchiTableInner width100"><span id="I_City"></span></div>
                    </div>
                </div>
                <!-- OchiHalf -->
                <div class="OchiHalf " id="div_outer_office">
                    <div class="OchiCell OchiTitle TitleSetWidth">承辦局處</div>
                    <div class="OchiCell width100">
                        <div class="OchiTableInner width100"><span id="I_Office"></span></div>
                    </div>
                </div>
                <!-- OchiHalf -->
            </div>
            <!-- OchiRow -->

            <!--執行期程 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">執行期程</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100">
                        <div class="OchiCellInner nowrap textcenter">開始:</div>
                        <div class="OchiCellInner width20"><span id="startdate"></span></div>
                        <div class="OchiCellInner nowrap textcenter">&nbsp;~&nbsp;</div>
                        <div class="OchiCellInner nowrap textcenter">結束:</div>
                        <div class="OchiCellInner width20"><span id="enddate"></span></div>
                        <div class="OchiCellInner nowrap textcenter">合計:</div>
                        <div class="OchiCellInner width10"><span id="totalmonth"></span></div>
                        <div class="OchiCellInner nowrap width30">&nbsp;月</div>
                    </div>
                    <!-- OchiTableInner -->
                    <!-- cell內容end -->
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 第一期起訖 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">第1期</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100">
                        <div class="OchiCellInner nowrap textcenter">開始:</div>
                        <div class="OchiCellInner width20"><span id="I_1_Sdate"></span></div>
                        <div class="OchiCellInner nowrap textcenter">&nbsp;~&nbsp;</div>
                        <div class="OchiCellInner nowrap textcenter">結束:</div>
                        <div class="OchiCellInner width20"><span id="I_1_Edate"></span></div>
                        <div class="OchiCellInner nowrap textcenter" style="display:none;">合計:</div>
                        <div class="OchiCellInner width45"><span id="I_1_total"></span></div>
                        <div class="OchiCellInner nowrap width30" style="display:none;">&nbsp;月</div>
                    </div>
                    <!-- cell內容end -->
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 第二期起訖 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">第2期</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100">
                        <div class="OchiCellInner nowrap textcenter">開始:</div>
                        <div class="OchiCellInner width20"><span id="I_2_Sdate"></span></div>
                        <div class="OchiCellInner nowrap textcenter">&nbsp;~&nbsp;</div>
                        <div class="OchiCellInner nowrap textcenter">結束:</div>
                        <div class="OchiCellInner width20"><span id="I_2_Edate"></span></div>
                        <div class="OchiCellInner nowrap textcenter" style="display:none;">合計:</div>
                        <div class="OchiCellInner width45"><span id="I_2_total"></span></div>
                        <div class="OchiCellInner nowrap width30" style="display:none;">&nbsp;月</div>
                    </div>
                    <!-- OchiTableInner -->
                    <!-- cell內容end -->
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 第三期起訖 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">第3期</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100">
                        <div class="OchiCellInner nowrap textcenter">開始:</div>
                        <div class="OchiCellInner width20"><span id="I_3_Sdate"></span></div>
                        <div class="OchiCellInner nowrap textcenter">&nbsp;~&nbsp;</div>
                        <div class="OchiCellInner nowrap textcenter">結束:</div>
                        <div class="OchiCellInner width20"><span id="I_3_Edate"></span></div>
                        <div class="OchiCellInner nowrap textcenter" style="display:none;">合計:</div>
                        <div class="OchiCellInner width45"><span id="I_3_total"></span></div>
                        <div class="OchiCellInner nowrap width30" style="display:none;">&nbsp;月</div>
                    </div><!-- OchiTableInner -->
                    <!-- cell內容end -->
                </div><!-- OchiCell -->
            </div><!-- OchiRow -->

            <!-- 核定之經費額度 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">核定之經費額度</div>
                <div class="OchiCell width100">
                    <div class="stripeMe margin5T font-normal">
                        <table id="moneyTab" width="100%" border="0" cellspacing="0" cellpadding="0">
                            <thead>
                                <tr>
                                    <th nowrap="nowrap">期別</th>
                                    <th nowrap="nowrap">第1期</th>
                                    <th nowrap="nowrap">第2期</th>
                                    <th nowrap="nowrap">第3期</th>
                                    <th nowrap="nowrap">全程</th>
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                    <div class="textright">單位:仟元</div>
                    <div class="twocol margin5T">
                        <div class="left"><a href="javascript:void(0);" id="MoneyUpBtn" class="genbtnS" style="display:none;">核定函上傳</a></div> 
                    </div>
                    <div class="mfile" style="margin-top:10px;display:none;">附件檔</div>
                    <div class="stripeMe mfile" style="margin-bottom:10px;display:none;">
                        <table id="moneyFileList" width="100%" border="0" cellspacing="0" cellpadding="0">
                            <thead>
                                <tr>
                                    <th>檔案名稱</th>
                                    <th>刪除</th>
                                </tr>
                            </thead>
                        </table>
                    </div>
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 其他經費來源 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">其他經費來源</div>
                <div class="OchiCell width100">
                    <div id="self" style="display:none;">自籌款，金額：<span id="I_Other_Oneself_Money"></span>&nbsp;仟元</div>
                    <div id="other" class="margin10T" style="display:none;">其他機關補助，機關名稱：<span id="I_Other_Other_name"></span>，金額:<span id="I_Other_Other_Money"></span>&nbsp;仟元</div>
                </div><!-- OchiCell -->
            </div><!-- OchiRow -->

            <!-- 承諾節電目標 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">承諾節電目標</div>
                <div class="OchiCell width100">
                    <span id="I_Target"></span>
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 本期計畫推動摘要 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">本期計畫推動摘要</div>
                <div class="OchiCell width100">
                    <span id="I_Summary"></span>
                    <div class="twocol margin5T">
                        <div class="left"><a href="javascript:void(0);" id="PlanDescUpBtn" class="genbtnS" style="display:none;">核定函上傳</a></div>
                    </div>
                    <div class="pdfile" style="margin-top:5px;display:none;">附件檔</div>
                    <div class="stripeMe pdfile" style="margin-bottom:10px;display:none;">
                        <table id="plandescFileList" width="100%" border="0" cellspacing="0" cellpadding="0">
                            <thead>
                                <tr>
                                    <th>檔案名稱</th>
                                    <th>刪除</th>
                                </tr>
                            </thead>
                        </table>
                    </div>
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 設備汰換與智慧用電預計完成數 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">設備汰換與智慧用電預計完成數</div>
                <div class="OchiCell width100">
                    <div class="stripeMe margin5TB font-normal">
                        <table id="finishTab" width="100%" border="0" cellspacing="0" cellpadding="0">
                            <thead>
                                <tr>
                                    <th nowrap="nowrap">項目</th>
                                    <th nowrap="nowrap">第1期</th>
                                    <th nowrap="nowrap">第2期</th>
                                    <th nowrap="nowrap">第3期</th>
                                    <th nowrap="nowrap">全程</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td>無風管冷氣(kW)<br />
                                        註:每台冷氣約4kW</td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item1_1" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item1_2" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item1_3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item1_all" disabled="disabled" /></td>
                                </tr>
                                <tr>
                                    <td>老舊辦公室照明(具)</td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item2_1" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item2_2" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item2_3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_Finish_item2_all" disabled="disabled" /></td>
                                </tr>
                                <tr>
                                    <td nowrap="nowrap">室內停車場智慧照明(盞)</td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item3_1" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item3_2" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item3_3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_Finish_item3_all" disabled="disabled" /></td>
                                </tr>
                                <tr>
                                    <td>中型能管系統(套)</td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item4_1" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item4_2" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item4_3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_Finish_item4_all" disabled="disabled" /></td>
                                </tr>
                                <tr>
                                    <td>大型能管系統(套)</td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item5_1" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item5_2" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item5_3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_Finish_item5_all" disabled="disabled" /></td>
                                </tr>

                            </tbody>
                        </table>
                    </div>

                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 擴大補助預計完成數 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">擴大補助預計完成數</div>
                <div class="OchiCell width100">
                    <div class="stripeMe margin5TB font-normal">
                        <table id="exFinishTab" width="100%" border="0" cellspacing="0" cellpadding="0"></table>
                    </div>
                </div><!-- OchiCell -->
            </div><!-- OchiRow -->

            <!-- 簽核資料 單欄 -->
            <div class="OchiRow" name="div_chk">
                <div class="OchiCell OchiTitle TitleSetWidth">承辦人</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100">
                        <div class="OchiCellInner nowrap textcenter">姓名：</div>
                        <div class="OchiCellInner width28"><span id="mb_Name"></span></div>
                        <div class="OchiCellInner nowrap textcenter">職稱：</div>
                        <div class="OchiCellInner width28"><span id="mb_JobTitle"></span></div>
                        <div class="OchiCellInner nowrap textcenter">電話：</div>
                        <div class="OchiCellInner width28"><span id="mb_Tel"></span></div>
                    </div>
                    <!-- OchiTableInner -->
                    <div class="OchiTableInner width100 margin10T">
                        <div class="OchiCellInner nowrap textcenter">手機：</div>
                        <div class="OchiCellInner width28"><span id="mb_Phone"></span></div>
                        <div class="OchiCellInner nowrap textcenter">傳真：</div>
                        <div class="OchiCellInner width28"><span id="mb_Fax"></span></div>
                        <div class="OchiCellInner width33"></div>
                    </div>
                    <!-- OchiTableInner -->
                    <div class="OchiTableInner width100 margin10T">
                        <div class="OchiCellInner nowrap textcenter">Email：</div>
                        <div class="OchiCellInner width100"><span id="mb_Email"></span></div>
                    </div>
                    <!-- OchiTableInner -->
                    <div class="OchiTableInner width100 margin10T">
                        <div class="OchiCellInner nowrap textcenter">地址：</div>
                        <div class="OchiCellInner width100"><span id="mb_Addr"></span></div>
                    </div>
                    <!-- OchiTableInner -->
                    <!-- cell內容end -->
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 簽核資料 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">承辦主管</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100">
                        <div class="OchiCellInner nowrap textcenter">姓名：</div>
                        <div class="OchiCellInner width28"><span id="mm_Name"></span></div>
                        <div class="OchiCellInner nowrap textcenter">職稱：</div>
                        <div class="OchiCellInner width28"><span id="mm_JobTitle"></span></div>
                        <div class="OchiCellInner nowrap textcenter">電話：</div>
                        <div class="OchiCellInner width28"><span id="mm_Tel"></span></div>
                    </div>
                    <!-- OchiTableInner -->
                    <div class="OchiTableInner width100 margin10T">
                        <div class="OchiCellInner nowrap textcenter">手機：</div>
                        <div class="OchiCellInner width28"><span id="mm_Phone"></span></div>
                        <div class="OchiCellInner nowrap textcenter">傳真：</div>
                        <div class="OchiCellInner width28"><span id="mm_Fax"></span></div>
                        <div class="OchiCellInner width33"></div>
                    </div>
                    <!-- OchiTableInner -->
                    <div class="OchiTableInner width100 margin10T">
                        <div class="OchiCellInner nowrap textcenter">Email：</div>
                        <div class="OchiCellInner width100"><span id="mm_Email"></span></div>
                    </div>
                    <!-- OchiTableInner -->
                    <div class="OchiTableInner width100 margin10T">
                        <div class="OchiCellInner nowrap textcenter">地址：</div>
                        <div class="OchiCellInner width100"><span id="mm_Addr"></span></div>
                    </div>
                    <!-- OchiTableInner -->
                    <!-- cell內容end -->
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->
        </div>
        <br />
        <!--計畫查核點-->
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">計畫查核點</span></div>
        </div>
        <div class="font-size3 margin10T">期別：
            <select id="cp_period" name="period" class="inputex">
                <option value="1">第一期</option>
                <option value="2">第二期</option>
                <option value="3">第三期</option>
            </select>
            <%--<span id="loadimg" style="display:none;"><img src="../App_Themes/images/loading.gif" height="25" />讀取中...</span>--%>
        </div>
        <div class="font-size3 margin10T">(一)節電基礎工作</div>
        <div class="stripecomplex margin5T font-normal">
            <table id="basicworkTab" width="100%" border="0" cellspacing="0" cellpadding="0">
                <thead>
                    <tr>
                        <th nowrap="nowrap">推動項目</th>
                        <th nowrap="nowrap" class="width10">查核點</th>
                        <th nowrap="nowrap">預定時間</th>
                        <th nowrap="nowrap" class="width50">查核點概述</th>
                    </tr>
                </thead>
            </table>
        </div>
        <div class="font-size3 margin20T">(二)因地制宜</div>
        <div class="stripecomplex margin5T font-normal">
            <table id="placeTab" width="100%" border="0" cellspacing="0" cellpadding="0">
                <thead>
                    <tr>
                        <th nowrap="nowrap">推動項目</th>
                        <th nowrap="nowrap" class="width10">查核點</th>
                        <th nowrap="nowrap">預定時間</th>
                        <th nowrap="nowrap" class="width50">查核點概述</th>
                    </tr>
                </thead>
            </table>
        </div>
        <div class="font-size3 margin20T">(三)設備汰換及智慧用電</div>
        <div class="stripecomplex margin5T font-normal">
            <table id="smartTab" width="100%" border="0" cellspacing="0" cellpadding="0">
                <thead>
                    <tr>
                        <th nowrap="nowrap">推動項目</th>
                        <th nowrap="nowrap" class="width10">查核點</th>
                        <th nowrap="nowrap">預定時間</th>
                        <th nowrap="nowrap" class="width50">查核點概述</th>
                    </tr>
                </thead>
            </table>
        </div>
        <div class="font-size3 margin20T">(四)擴大補助</div>
        <div class="stripecomplex margin5T font-normal">
            <table id="allowanceTab" width="100%" border="0" cellspacing="0" cellpadding="0">
                <thead>
                    <tr>
                        <th nowrap="nowrap">推動項目</th>
                        <th nowrap="nowrap" class="width10">查核點</th>
                        <th nowrap="nowrap">預定時間</th>
                        <th nowrap="nowrap" class="width50">查核點概述</th>
                    </tr>
                </thead>
            </table>
        </div>
        <br />
        <!--預定工作進度-->
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">預定工作進度</span></div>
        </div>
        <div class="font-size3 margin10T">期別：
            <select id="p_period" name="period" class="inputex">
                <option value="1">第一期</option>
                <option value="2">第二期</option>
                <option value="3">第三期</option>
            </select>
        </div>

        <div class="font-size3 margin10T" id="div1">(一)節電基礎工作</div>
        <div class="stripeMe margin5T font-normal">
            <div id="bwTabLoad">資料讀取中...</div>
            <table width="100%" border="0" cellspacing="0" cellpadding="0" id="bwTab"></table>
        </div>

        <div class="font-size3 margin20T" id="div2">(二)因地制宜</div>
        <div class="stripeMe margin5T font-normal">
            <div id="pTabLoad">資料讀取中...</div>
            <table width="100%" border="0" cellspacing="0" cellpadding="0" id="pTab"></table>
        </div>

        <div class="font-size3 margin20T" id="div3">(三)設備汰換及智慧用電</div>
        <div class="stripeMe margin5T font-normal">
            <div id="sTabLoad">資料讀取中...</div>
            <table width="100%" border="0" cellspacing="0" cellpadding="0" id="sTab"></table>
        </div>
        
        <div class="font-size3 margin20T" id="div4">(四)擴大補助</div>
        <div class="stripeMe margin5T font-normal">
            <div id="exTabLoad">資料讀取中...</div>
            <table width="100%" border="0" cellspacing="0" cellpadding="0" id="exTab"></table>
        </div>

        <div style="text-align:right; margin-top:10px; margin-bottom:10px;">
            <a href="javascript:void(0);" class="genbtn" id="backbtn" style="display:none;">回上一頁</a>
            <a href="javascript:void(0);" class="genbtn" id="subbtn" style="display:none;">定稿</a>
            <a href="javascript:void(0);" class="genbtn" id="exbtn" style="display:none;">匯出</a>
        </div>
    </div>

    <div id="exblock" style="display:none; text-align:center;">
        <div style="margin-bottom:10px;">請選擇匯出檔案類型</div>
        <div style="margin-bottom:10px;">
            <input type="radio" name="extype" value="W" checked="checked" />&nbsp;Word&nbsp;&nbsp;
            <input type="radio" name="extype" value="P" />&nbsp;PDF
        </div>
        <input type="button" id="expbtn" value="確定" class="genbtn" />
    </div>
</asp:Content>

