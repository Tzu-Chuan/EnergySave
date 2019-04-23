<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="Chart_DeviceReplace.aspx.cs" Inherits="WebPage_Chart_DeviceReplace" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
<script src="../js_highchart/highcharts.js"></script>
<script src="../js_highchart/modules/series-label.js"></script>
<script src="../js_highchart/modules/exporting.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            getddl("02", "#ddlCity");
            getData();
            //Highcharts 千分位
            Highcharts.setOptions({
                lang: {
                    thousandsSep: ','
                }
            });

            $(document).on("change", "#ddlCity,#ddlStage", function () {
                getData();
            });

            //var aircond_json = [{
            //    name: '當期規劃數量',
            //    data: [25500, 25500, 25500, 25500]
            //}, {
            //    name: '申請數量',
            //    data: [125, 6000, 15000, 25000]
            //}, {
            //    name: '完成數量',
            //    data: [100, 3800, 12000, 24000]
            //}];
        });

        function getData() {
            $.ajax({
                type: "POST",
                async: true, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/ChartReDevice.aspx",
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
                                var category = $.parseJSON($(this).children("type").text().trim());
                                var series = $.parseJSON($(this).children("series").text().trim());
                                linechart("AirCond", "無風管冷氣 (kW)", category, series);
                            });
                        }

                        if ($(data).find("dt_air").length > 0) {
                            var tabstr = getTable(data, "無風管冷氣 (kW)", "dt_air");
                            $("#airlist").append(tabstr);
                        }

                        if ($(data).find("old").length > 0) {
                            $(data).find("old").each(function () {
                                var category = $.parseJSON($(this).children("type").text().trim());
                                var series = $.parseJSON($(this).children("series").text().trim());
                                linechart("OldOffice", "老舊辦公室照明 (具)", category, series);
                            });
                        }

                        if ($(data).find("dt_old").length > 0) {
                            var tabstr = getTable(data, "老舊辦公室照明 (具)", "dt_old");
                            $("#oldlist").append(tabstr);
                        }

                        if ($(data).find("parking").length > 0) {
                            $(data).find("parking").each(function () {
                                var category = $.parseJSON($(this).children("type").text().trim());
                                var series = $.parseJSON($(this).children("series").text().trim());
                                linechart("CarParking", "室內停車場智慧照明 (盞)", category, series);
                            });
                        }

                        if ($(data).find("dt_parking").length > 0) {
                            var tabstr = getTable(data, "室內停車場智慧照明 (盞)", "dt_parking");
                            $("#parkinglist").append(tabstr);
                        }

                        if ($(data).find("middle").length > 0) {
                            $(data).find("middle").each(function () {
                                var category = $.parseJSON($(this).children("type").text().trim());
                                var series = $.parseJSON($(this).children("series").text().trim());
                                linechart("MiddleManage", "中型能源管理系統 (套)", category, series);
                            });
                        }

                        if ($(data).find("dt_middle").length > 0) {
                            var tabstr = getTable(data, "中型能源管理系統 (套)", "dt_middle");
                            $("#middlelist").append(tabstr);
                        }

                        if ($(data).find("large").length > 0) {
                            $(data).find("large").each(function () {
                                var category = $.parseJSON($(this).children("type").text().trim());
                                var series = $.parseJSON($(this).children("series").text().trim());
                                linechart("LargeManage", "大型能源管理系統 (套)", category, series);
                            });
                        }

                        if ($(data).find("dt_large").length > 0) {
                            var tabstr = getTable(data, "大型能源管理系統 (套)", "dt_large");
                            $("#largelist").append(tabstr);
                        }
                    }
                }
            });
        }

        function linechart(tag, titletext, type, series) {
            Highcharts.chart(tag, {
                chart: {
                    type: 'column'
                },
                credits: {
                    enabled: false
                },
                title: {
                    text: titletext
                },
                xAxis: {
                    categories: type
                },
                yAxis: {
                    title: {
                        text: ''
                    }
                },
                legend: {
                    layout: 'horizontal',
                    align: 'center',
                    verticalAlign: 'bottom'
                },
                plotOptions: {
                    series: {
                        label: {
                            connectorAllowed: true
                        }
                    }
                },
                series: series,
                responsive: {
                    rules: [{
                        condition: {
                            maxWidth: 500
                        },
                        chartOptions: {
                            legend: {
                                layout: 'horizontal',
                                align: 'center',
                                verticalAlign: 'bottom'
                            }
                        }
                    }]
                }
            });
        }

        function getTable(data, title, xmltag) {
            var unitStr = "";
            switch (xmltag) {
                case "dt_air":
                    unitStr = " (kW)";
                    break;
                case "dt_old":
                    unitStr = " (具)";
                    break;
                case "dt_parking":
                    unitStr = " (盞)";
                    break;
                case "dt_middle":
                case "dt_large":
                    unitStr = " (套)";
                    break;
            }
            var tmpval = '';
            var tabstr1 = '<tr><th>' + title + '</th>';
            var tabstr2 = '';
            var tabstr3 = '<tr><th>申請數量</th>';
            var tabstr4 = '<tr><th>完成數量</th>';
            $(data).find(xmltag).each(function () {
                var yearROC = parseInt($(this).children("RS_Year").text().trim()) - 1911;
                tabstr1 += '<th>' + yearROC + '年第' + ReROCStr($(this).children("RS_Season").text().trim()) + '季' + unitStr + '</th>';
                tabstr2 = $(this).children("SUM_S_money").text().trim();
                tabstr3 += '<td align="center">' + $(this).children("RM_SUM_money").text().trim() + '</td>';
                tabstr4 += '<td align="center">' + $(this).children("SUM_C_money").text().trim() + '</td>';
            });
            tabstr1 += '</tr>';
            tabstr2 = '<tr><th>當期規劃數量</th><td align="center" colspan="' + $(data).find("dt_air").length + '">' + tabstr2 + '</td></tr>';
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
            <div class="left"><span class="filetitle font-size5">設備汰換進度</span></div>
            <div class="right">附加圖表 / 地方圖表 / 設備汰換進度</div>
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

