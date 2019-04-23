<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="Chart_ExpectSave.aspx.cs" Inherits="WebPage_Chart_ExpectSave" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
<script src="../js_highchart/highcharts.js"></script>
<script src="../js_highchart/modules/exporting.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            getddl("02", "#ddlCity");
            getData();
            Highcharts.setOptions({
                lang: {
                    thousandsSep: ','
                }
            });

            $(document).on("change", "#ddlCity,#ddlStage", function () {
                getData();
            });

            //var aircond_json = [{
            //    name: '當期規劃數節電量',
            //    data: [7936875]
            //}, {
            //    name: '第一季',
            //    data: ['',38906, 31125]
            //}, {
            //    name: '第二季',
            //    data: ['',1867500, 1182750]
            //}, {
            //    name: '第三季',
            //    data: ['',4668750, 3735000]
            //}, {
            //    name: '第四季',
            //    data: ['',7781250, 7470000]
            //}];
        });

        function getData() {
            $.ajax({
                type: "POST",
                async: true, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/ChartExSave.aspx",
                data: {
                    City: $("#ddlCity").val(),
                    Stage: $("#ddlStage").val()
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if (data != null) {
                        data = $.parseXML(data);
                        if ($("Comp", data).text() == "SA")
                            $("#mblock").show();

                        $("table").empty();
                        $(".chart").empty();
                        if ($(data).find("air").length > 0) {
                            $(data).find("air").each(function () {
                                var series = $.parseJSON($(this).children("series").text().trim());
                                columnchart("AirCond", "無風管冷氣預期節電量", series);
                            });
                        }

                        if ($(data).find("dt_air").length > 0) {
                            var tabstr = getTable(data, "無風管冷氣", "dt_air");
                            $("#airlist").append(tabstr);
                        }

                        if ($(data).find("old").length > 0) {
                            $(data).find("old").each(function () {
                                var series = $.parseJSON($(this).children("series").text().trim());
                                columnchart("OldOffice", "老舊辦公室照明", series);
                            });
                        }

                        if ($(data).find("dt_old").length > 0) {
                            var tabstr = getTable(data, "老舊辦公室照明", "dt_old");
                            $("#oldlist").append(tabstr);
                        }

                        if ($(data).find("parking").length > 0) {
                            $(data).find("parking").each(function () {
                                var series = $.parseJSON($(this).children("series").text().trim());
                                columnchart("CarParking", "室內停車場智慧照明", series);
                            });
                        }

                        if ($(data).find("dt_parking").length > 0) {
                            var tabstr = getTable(data, "室內停車場智慧照明", "dt_parking");
                            $("#parkinglist").append(tabstr);
                        }

                        if ($(data).find("middle").length > 0) {
                            $(data).find("middle").each(function () {
                                var series = $.parseJSON($(this).children("series").text().trim());
                                columnchart("MiddleManage", "中型能源管理系統", series);
                            });
                        }

                        if ($(data).find("dt_middle").length > 0) {
                            var tabstr = getTable(data, "中型能源管理系統", "dt_middle");
                            $("#middlelist").append(tabstr);
                        }

                        if ($(data).find("large").length > 0) {
                            $(data).find("large").each(function () {
                                var series = $.parseJSON($(this).children("series").text().trim());
                                columnchart("LargeManage", "大型能源管理系統", series);
                            });
                        }

                        if ($(data).find("dt_large").length > 0) {
                            var tabstr = getTable(data, "大型能源管理系統", "dt_large");
                            $("#largelist").append(tabstr);
                        }
                    }
                }
            });
        }

        function columnchart(tag, titletext, series) {
            Highcharts.chart(tag, {
                credits: {
                    enabled: false
                },
                chart: {
                    type: 'column'
                },
                title: {
                    text: titletext + " (度)"
                },
                xAxis: {
                    categories: [
                        '當期規劃數節電量',
                        '申請數節電量',
                        '完成數節電量'
                    ],
                    crosshair: true
                },
                yAxis: {
                    title: {
                        text: ''
                    }
                },
                tooltip: {
                    headerFormat: '<span style="font-size:12px">{point.key}</span><table>',
                    pointFormat: '<tr><td style="color:{series.color};padding:0">{series.name}: </td>' +
                    '<td style="padding:0"><b>{point.y}</b></td></tr>',
                    footerFormat: '</table>',
                    shared: true,
                    useHTML: true
                },
                plotOptions: {
                    column: {
                        dataLabels: {
                            enabled: true
                        },
                        pointPadding: 0.2,
                        borderWidth: 0
                    }
                },
                series: series
            });

            //Highcharts 千分位
            //Highcharts.setOptions({
            //    lang: {
            //        thousandsSep: ','
            //    }
            //});
        }

        function getTable(data, title, xmltag) {
            var tmpval = '';
            var tabstr1 = '<tr><th>' + title + ' (度)</th>';
            var tabstr2 = '';
            var tabstr3 = '<tr><th>申請數節電量</th>';
            var tabstr4 = '<tr><th>完成數節電量</th>';
            $(data).find(xmltag).each(function () {
                tabstr1 += '<th>第' + ReROCStr($(this).children("RS_Season").text().trim()) + '季 (度)</th>';
                tabstr2 = $(this).children("SUM_S_money").text().trim();
                tabstr3 += '<td align="center">' + $(this).children("RM_SUMPre_money").text().trim() + '</td>';
                tabstr4 += '<td align="center">' + $(this).children("RM_SUMFinish_money").text().trim() + '</td>';
            });
            tabstr1 += '</tr>';
            tabstr2 = '<tr><th>當期規劃數節電量</th><td align="center" colspan="' + $(data).find("dt_air").length + '">' + tabstr2 + '</td></tr>';
            tabstr3 += '</tr>';
            tabstr4 += '</tr>';
            tmpval = tabstr1 + tabstr2 + tabstr3 + tabstr4;
            return tmpval;
        }

        function ReROCStr(str) {
            var tmp = "";
            switch (str) {
                case "1":
                    tmp = "一";
                    break;
                case "2":
                    tmp = "二";
                    break;
                case "3":
                    tmp = "三";
                    break;
                case "4":
                    tmp = "四";
                    break;
            }
            return tmp;
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
                        var ddlstr = '';
                        if ($(data).find("code").length > 0) {
                            if (gno=="02") {
                                ddlstr += "<option value='all'>全部</option>";
                            }
                            $(data).find("code").each(function (i) {
                                ddlstr += '<option value="' + $(this).attr("v") + '">' + $(this).attr("desc") + '</option>';
                            });
                        }
                        $(tagName).append(ddlstr);
                    }
                }
            });
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">預期節電量</span></div>
            <div class="right">附加圖表 / 地方圖表 / 預期節電量</div>
        </div>

        <div id="mblock" style="margin-top:10px;display:none;">執行機關：<select id="ddlCity" class="inputex"></select></div>
        期別：
        <select id="ddlStage" class="inputex" style="margin-top:10px;">
            <option value="1">第一期</option>
            <option value="2">第二期</option>
            <option value="3">第三期</option>
        </select>

        <div id="AirCond" class="margin20T chart"></div>
        <div class="stripeMe">
            <table id="airlist" border="0" cellspacing="0" cellpadding="0" width="100%"></table>
        </div>

        <div id="OldOffice" class="margin20T chart"></div>
        <div class="stripeMe">
            <table id="oldlist" border="0" cellspacing="0" cellpadding="0" width="100%"></table>
        </div>

        <div id="CarParking" class="margin20T chart"></div>
        <div class="stripeMe">
            <table id="parkinglist" border="0" cellspacing="0" cellpadding="0" width="100%"></table>
        </div>

        <div id="MiddleManage" class="margin20T chart"></div>
        <div class="stripeMe">
            <table id="middlelist" border="0" cellspacing="0" cellpadding="0" width="100%"></table>
        </div>

        <div id="LargeManage" class="margin20T chart"></div>
        <div class="stripeMe">
            <table id="largelist" border="0" cellspacing="0" cellpadding="0" width="100%"></table>
        </div>
    </div><br />
</asp:Content>

