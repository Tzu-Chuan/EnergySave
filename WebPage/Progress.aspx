<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="Progress.aspx.cs" Inherits="WebPage_Progress" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script type="text/javascript">
        //AutoSave Function - Auto run every 20 minutes
        setInterval(function () {
            if ($("#workratioStatus").val() != "") {
                alert("工作比重不可超過100%");
                return;
            }

            if ($("#autoStatus").val() == "true")
                return;

            var iframe = $('<iframe name="postiframe" id="postiframe" style="display: none" />');
            var mid = $('<input type="hidden" name="mid" id="mid" value="' + $.getParamValue('v') + '" />');

            var form = $("form")[0];

            $("#postiframe").remove();
            $("input[name='mid']").remove();

            form.appendChild(iframe[0]);
            form.appendChild(mid[0]);

            form.setAttribute("action", "../handler/AutoSave_Progress.ashx");
            form.setAttribute("method", "post");
            form.setAttribute("enctype", "multipart/form-data");
            form.setAttribute("encoding", "multipart/form-data");
            form.setAttribute("target", "postiframe");
            form.submit();
        }, 1200000);//20 minutes

        $(document).ready(function () {
            getData();

            //限制只能輸入數字
            $(document).on("keyup", ".num", function () {
                if (/[^0-9\.]/g.test(this.value)) {
                    this.value = this.value.replace(/[^0-9\.]/g, '');
                }
            });

            //上一步/上一頁 button
            $(document).on("click", "#preivousstep,#preivouspage", function () {
                if (this.id =="preivousstep") {
                    if (confirm("請確認資料已儲存，是否回到上一步？")) {
                        location.href = "CheckPoint.aspx?v=" + $.getParamValue('v');
                    }
                }
                else
                    location.href = "CheckPoint.aspx?v=" + $.getParamValue('v');
            });
            
            $(document).on("click", "#savebtn", function () {
                if ($("#workratioStatus").val() != "") {
                    alert("工作比重不可超過100%");
                    return;
                }

                //判斷是否為整數
                //var intStatus = false;
                //$("input[name='cp_process']").each(function () {
                //    if (Math.floor(this.value) != this.value && $.isNumeric(this.value))
                //        intStatus = true;
                //});
                //if (intStatus == true) {
                //    alert("【累計預定進度】請輸入整數");
                //    return;
                //}


                var iframe = $('<iframe name="postiframe" id="postiframe" style="display: none" />');
                var mid = $('<input type="hidden" name="mid" id="mid" value="' + $.getParamValue('v') + '" />');

                var form = $("form")[0];

                $("#postiframe").remove();
                $("input[name='mid']").remove();

                form.appendChild(iframe[0]);
                form.appendChild(mid[0]);

                form.setAttribute("action", "../handler/saveProgress.ashx");
                form.setAttribute("method", "post");
                form.setAttribute("enctype", "multipart/form-data");
                form.setAttribute("encoding", "multipart/form-data");
                form.setAttribute("target", "postiframe");
                form.submit();
            });

            //期別 change
            $(document).on("change", "#period", function () {
                if (confirm('請確認資料已儲存，是否切換期別？')) {
                    getData();
                }
            });

            //工作比重 change
            $(document).on("change", ".wr01,.wr02,.wr03,.wr04", function () {
                var thisid = $(this).attr("class").split(' ');
                var sum = 0;
                $("." + thisid[2]).each(function () {
                    var v = (this.value == "") ? 0 : this.value;
                    sum += parseFloat(v);
                });
                if (sum.toFixed(0) > 100) {
                    $($(this).attr("tid")).css("color", "red");
                    $("#workratioStatus").val("Y");
                }
                else {
                    $($(this).attr("tid")).css("color", "");
                    $("#workratioStatus").val("");
                }

                if (this.value != "")
                    this.value = Number(parseFloat(this.value).toFixed(2));
                $($(this).attr("tid")).html(Number(sum.toFixed(2)));
            });

            //累計預定進度 change
            $(document).on("change", ".cpp", function () {
                var itemAry = [];
                var valAry = [];
                for (var b = 0; b < parseInt($(this).attr("mcount")); b++) {
                    var pvtmp = "pv" + $(this).attr("tp") + b;
                    var sum = 0;
                    $(".cpp").each(function () {
                        if ($(this).attr("pv") == pvtmp) {
                            var v = (this.value == "") ? 0 : this.value;
                            //sum += parseFloat(v);
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

                    if (this.value != "")
                        this.value = Number(parseFloat(this.value).toFixed(2));

                    $("#cppTotal" + $(this).attr("tp") + b).html(Number(sum.toFixed(2)));
                }
            });
        });//js end

        function getData() {
            $.ajax({
                type: "POST",
                async: true, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getProgress.aspx",
                data: {
                    period: $("#period").val(),
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
                                            $("#bwTab").empty();
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
                                                    var year = new Date($(this).children("I_" + $("#period").val() + "_Sdate").text().trim()).getFullYear() - 1911;
                                                    var month = new Date($(this).children("I_" + $("#period").val() + "_Sdate").text().trim()).getMonth() + 1;
                                                    var day = new Date($(this).children("I_" + $("#period").val() + "_Sdate").text().trim()).getDate();
                                                    startday = year + "年" + month + "月" + day + "日";
                                                    //結束日
                                                    year = new Date($(this).children("I_" + $("#period").val() + "_Edate").text().trim()).getFullYear() - 1911;
                                                    month = new Date($(this).children("I_" + $("#period").val() + "_Edate").text().trim()).getMonth() + 1;
                                                    day = new Date($(this).children("I_" + $("#period").val() + "_Edate").text().trim()).getDate();
                                                    endday = year + "年" + month + "月" + day + "日";
                                                });

                                                var tabstr = '<tr>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">推動項目</th>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">工作比重(%)</th>';
                                                tabstr += '<th rowspan="3">年&nbsp;&nbsp;&nbsp;月<br>進度(%)</th>';
                                                tabstr += '<th colspan="' + monthAry.length + '">' + startday + '至' + endday + '</th></tr >';
                                                monthstr = '<tr>' + monthstr + '</tr>';
                                                tabstr += yearstr + monthstr + tabdetail;
                                                $("#bwTab").append(tabstr);

                                                //各月份合計(湊完table再抓#ID做處理)
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
                                            $("#pTab").empty();
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
                                                    var year = new Date($(this).children("I_" + $("#period").val() + "_Sdate").text().trim()).getFullYear() - 1911;
                                                    var month = new Date($(this).children("I_" + $("#period").val() + "_Sdate").text().trim()).getMonth() + 1;
                                                    var day = new Date($(this).children("I_" + $("#period").val() + "_Sdate").text().trim()).getDate();
                                                    startday = year + "年" + month + "月" + day + "日";
                                                    //結束日
                                                    year = new Date($(this).children("I_" + $("#period").val() + "_Edate").text().trim()).getFullYear() - 1911;
                                                    month = new Date($(this).children("I_" + $("#period").val() + "_Edate").text().trim()).getMonth() + 1;
                                                    day = new Date($(this).children("I_" + $("#period").val() + "_Edate").text().trim()).getDate();
                                                    endday = year + "年" + month + "月" + day + "日";
                                                });

                                                var tabstr = '<tr>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">推動項目</th>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">工作比重(%)</th>';
                                                tabstr += '<th rowspan="3">年&nbsp;&nbsp;&nbsp;月<br>進度(%)</th>';
                                                tabstr += '<th colspan="' + monthAry.length + '">' + startday + '至' + endday + '</th></tr >';
                                                tabstr += yearstr + monthstr + tabdetail;
                                                $("#pTab").append(tabstr);

                                                //各月份合計(湊完table再抓#ID做處理)
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
                                            $("#sTab").empty();
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
                                                    var year = new Date($(this).children("I_" + $("#period").val() + "_Sdate").text().trim()).getFullYear() - 1911;
                                                    var month = new Date($(this).children("I_" + $("#period").val() + "_Sdate").text().trim()).getMonth() + 1;
                                                    var day = new Date($(this).children("I_" + $("#period").val() + "_Sdate").text().trim()).getDate();
                                                    startday = year + "年" + month + "月" + day + "日";
                                                    //結束日
                                                    year = new Date($(this).children("I_" + $("#period").val() + "_Edate").text().trim()).getFullYear() - 1911;
                                                    month = new Date($(this).children("I_" + $("#period").val() + "_Edate").text().trim()).getMonth() + 1;
                                                    day = new Date($(this).children("I_" + $("#period").val() + "_Edate").text().trim()).getDate();
                                                    endday = year + "年" + month + "月" + day + "日";
                                                });

                                                var tabstr = '<tr>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">推動項目</th>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">工作比重(%)</th>';
                                                tabstr += '<th rowspan="3">年&nbsp;&nbsp;&nbsp;月<br>進度(%)</th>';
                                                tabstr += '<th colspan="' + monthAry.length + '">' + startday + '至' + endday + '</th></tr >';
                                                tabstr += yearstr + monthstr + tabdetail;
                                                $("#sTab").append(tabstr);

                                                //各月份合計(湊完table再抓#ID做處理)
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
                                            $("#exTab").empty();
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
                                                    var year = new Date($(this).children("I_" + $("#period").val() + "_Sdate").text().trim()).getFullYear() - 1911;
                                                    var month = new Date($(this).children("I_" + $("#period").val() + "_Sdate").text().trim()).getMonth() + 1;
                                                    var day = new Date($(this).children("I_" + $("#period").val() + "_Sdate").text().trim()).getDate();
                                                    startday = year + "年" + month + "月" + day + "日";
                                                    //結束日
                                                    year = new Date($(this).children("I_" + $("#period").val() + "_Edate").text().trim()).getFullYear() - 1911;
                                                    month = new Date($(this).children("I_" + $("#period").val() + "_Edate").text().trim()).getMonth() + 1;
                                                    day = new Date($(this).children("I_" + $("#period").val() + "_Edate").text().trim()).getDate();
                                                    endday = year + "年" + month + "月" + day + "日";
                                                });

                                                var tabstr = '<tr>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">推動項目</th>';
                                                tabstr += '<th rowspan="3" nowrap="nowrap">工作比重(%)</th>';
                                                tabstr += '<th rowspan="3">年&nbsp;&nbsp;&nbsp;月<br>進度(%)</th>';
                                                tabstr += '<th colspan="' + monthAry.length + '">' + startday + '至' + endday + '</th></tr >';
                                                tabstr += yearstr + monthstr + tabdetail;
                                                $("#exTab").append(tabstr);

                                                //各月份合計(湊完table再抓#ID做處理)
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

                            //權限
                            if ($("unVisiable", data).text() == "Y") {
                                $(".num").attr("disabled", "disabled");
                                $("#preivouspage").show();
                                $("#autoStatus").val("true");
                            }
                            else {
                                $("#preivousstep").show();
                                $("#savebtn").show();
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
                    //第一筆開頭
                    var tmpiname = "";
                    switch (typenum) {
                        default:
                            tmpiname = $(this).children("P_ItemName").text().trim();
                            break;
                        case "03":
                            tmpiname = get_TypeCn("07", $(this).children("P_ItemName").text().trim());
                            if ($(this).children("P_ItemName").text().trim() == "99")
                                tmpiname += "-" + $(this).children("CP_Desc").text().trim();
                            break;
                        case "04":
                            tmpiname = get_TypeCn("09", $(this).children("P_ItemName").text().trim());
                            if ($(this).children("P_ItemName").text().trim() == "99")
                                tmpiname += "-" + $(this).children("CP_Desc").text().trim();
                            break;
                    }
                    //var tmpiname = (typenum == "03") ? getCP_TypeCn($(this).children("P_ItemName").text().trim()) : $(this).children("P_ItemName").text().trim();
                    tabrow1 += '<tr><td rowspan="3" align="center"><strong>' + tmpiname + '</strong></td><td rowspan="3" align="center">';
                    tabrow1 += '<input type="text" class="inputex num wr' + typenum + '" tid="#wrTotal' + typenum + '" size="3" name="wr_value" value="' + $(this).children("P_WorkRatio").text().trim() + '" maxlength="5" />';
                    tabrow1 += '<input type="hidden" name="wr_pid" value="' + $(this).children("CP_ParentId").text().trim() + '" /></td><td align="center">查核點</td>';
                    tabrow2 += '<tr><td align="center">累計預定進度(%)</td>';
                    tabrow3 += '<tr><td nowrap="nowrap" align="center">查核點<br />進度說明</td><td colspan="' + monthAry.length + '">';
                    //因為資料庫與Table的X、Y軸相反
                    //為了好寫用上面整理好的月份陣列跑迴圈轉成同方向
                    for (var l = 0; l < monthAry.length; l++) {
                        var splitstr2 = monthAry[l].split('-');
                        //當對應到年跟月時
                        if ($(this).children("CP_ReserveYear").text().trim() == splitstr2[0] && $(this).children("CP_ReserveMonth").text().trim() == splitstr2[1]) {
                            tabrow1 += '<td valign="top" align="center">' + $(this).children("CP_Point").text().trim() + '</td>';
                            tabrow2 += '<td align="center"><input type="text" class="inputex num cpp" pv="pv' + typenum + l + '" item="' + $(this).children("CP_ParentId").text().trim().substr(0, 10) +
                                '" mcount="' + monthAry.length + '" tp="' + $(this).children("P_Type").text().trim()+'" size="3" name="cp_process" value="' + $(this).children("CP_Process").text().trim() + '" maxlength="5" />';
                            tabrow2 += '<input type="hidden" name="pv_pid" value="' + $(this).children("CP_Guid").text().trim() + '" /></td>';
                            tabrow3 += $(this).children("CP_Point").text().trim() + '&nbsp;&nbsp;' + $(this).children("CP_Desc").text().trim() + '<br />';
                            rnum += 1;
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
                            tabrow1 += '<td valign="top"  align="center">' + $(this).children("CP_Point").text().trim() + '</td>';
                            tabrow2 += '<td  align="center"><input type="text" class="inputex num cpp" pv="pv' + typenum + l + '" item="' + $(this).children("CP_ParentId").text().trim().substr(0, 10) +
                                '" mcount="' + monthAry.length + '" tp="' + $(this).children("P_Type").text().trim() +'" size="3" name="cp_process" value="' + $(this).children("CP_Process").text().trim() + '" maxlength="5" />';
                            tabrow2 += '<input type="hidden" name="pv_pid" value="' + $(this).children("CP_Guid").text().trim() + '" /></td>';
                            tabrow3 += $(this).children("CP_Point").text().trim() + '&nbsp;&nbsp;' + $(this).children("CP_Desc").text().trim() + '<br />';
                            rnum += 1;
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
            rowstr += '<td align="center">合　計 </td>';
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
                rowstr += '<td align="center"><span id="cppTotal' + typenum + i + '"></span>%</td>';
            }
            rowstr += '</tr>';
            return rowstr;
        }

        //累計預定進度(%)合計
        function monthTotal(tp, mCount) {
            var itemAry = [];
            var valAry = [];
            for (var b = 0; b < mCount; b++) {
                var pvtmp = "pv" + tp+ b;
                var sum = 0;
                $(".cpp").each(function () {
                    if ($(this).attr("pv") == pvtmp) {
                        var v = (this.value == "") ? 0 : this.value;
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

                $.each(valAry,function (i) {
                    sum += parseFloat(valAry[i]);
                });
                $("#cppTotal" + tp + b).html(Number(sum.toFixed(2)));
            }
        }

        function get_TypeCn(group,tp) {
            var str = "";
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/GetDDL.aspx",
                data: {
                    group: group
                },
                error: function (xhr) {
                    alert(xhr.responseText);
                },
                success: function (data) {
                    if ($(data).find("Error").length > 0) {
                        alert($(data).find("Error").attr("Message"));
                    }
                    else {
                        if ($(data).find("data_item").length > 0) {
                            $(data).find("data_item").each(function () {
                                if ($(this).children("C_Item").text().trim() == tp)
                                    str = $(this).children("C_Item_cn").text().trim();
                            });
                        }
                    }
                }
            });
            return str;
        }

        function feedback(str) {
            var form = document.body.getElementsByTagName('form')[0];
            form.target = '';
            form.method = "post";
            form.enctype = "application/x-www-form-urlencoded";
            form.encoding = "application/x-www-form-urlencoded";
            form.action = location;

            if (str.indexOf("Error") > -1)
                alert(str);

            if (str == "succeed") {
                alert("儲存完成");
            }
        }

        //自動存檔 feedback
        function autofeedback(str) {
            var form = document.body.getElementsByTagName('form')[0];
            form.target = '';
            form.method = "post";
            form.enctype = "application/x-www-form-urlencoded";
            form.encoding = "application/x-www-form-urlencoded";
            form.action = location;
            
            if (str == "succeed") {

            }
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <input type="hidden" id="autoStatus" value="false" />
    <input type="hidden" id="workratioStatus" value="" />
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">預定工作進度</span></div>
            <!-- left -->
            <div class="right">基本資料 / 預定工作進度</div>
            <!-- right -->
        </div>
        <!-- twocol -->
        <div class="font-size3 margin10T">期別：
            <select id="period" name="period" class="inputex">
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

        <div class="twocol margin15T margin5B">
            <div class="right">
                <a id="preivousstep" href="javascript:void(0);" class="genbtn" style="display:none;">上一步</a>
                <a id="savebtn" href="javascript:void(0);" class="genbtn" style="display:none;">儲存</a>
                <a id="preivouspage" href="javascript:void(0);" class="genbtn" style="display:none;">上一頁</a>
            </div>
        </div>
        <!-- twocol -->
    </div>
</asp:Content>

