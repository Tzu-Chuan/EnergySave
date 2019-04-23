<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="Chart_ApplyUnitAnalyze.aspx.cs" Inherits="WebPage_Chart_ApplyUnitAnalyze" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
<script src="../js_highchart/highcharts.js"></script>
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

            var aircond_json = [{
                name: '機關',
                y: 250
            }, {
                name: '學校',
                y: 107,
            }, {
                name: '服務業',
                y: 3200
            }];

            var oldoffice_json = [{
                name: '機關',
                y: 1500
            }, {
                name: '學校',
                y: 3100,
            }, {
                name: '服務業',
                y: 8054
            }];
        });


        function getData() {
            $.ajax({
                type: "POST",
                async: true, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/ChartUnitAnalyze.aspx",
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
                                piechart("AirCond", "無風管冷氣汰換申請數量 (台)", series);
                            });
                        }

                        if ($(data).find("dt_air").length > 0) {
                            var tabstr = getTable(data, "01", "dt_air");
                            $("#airlist").append(tabstr);
                        }

                        if ($(data).find("old").length > 0) {
                            $(data).find("old").each(function () {
                                var series = $.parseJSON($(this).children("series").text().trim());
                                piechart("OldOffice", "老舊辦公室照明汰換申請數量 (具)", series);
                            });
                        }

                        if ($(data).find("dt_old").length > 0) {
                            var tabstr = getTable(data, "02", "dt_old");
                            $("#oldlist").append(tabstr);
                        }

                        if ($(data).find("parking").length > 0) {
                            $(data).find("parking").each(function () {
                                var series = $.parseJSON($(this).children("series").text().trim());
                                piechart("CarParking", "室內停車場智慧照明汰換申請數量 (盞)", series);
                            });
                        }

                        if ($(data).find("dt_parking").length > 0) {
                            var tabstr = getTable(data, "03", "dt_parking");
                            $("#parkinglist").append(tabstr);
                        }

                        if ($(data).find("middle").length > 0) {
                            $(data).find("middle").each(function () {
                                var series = $.parseJSON($(this).children("series").text().trim());
                                piechart("MiddleManage", "中型能源管理系統汰換申請數量 (套)", series);
                            });
                        }

                        if ($(data).find("dt_middle").length > 0) {
                            var tabstr = getTable(data, "04", "dt_middle");
                            $("#middlelist").append(tabstr);
                        }

                        if ($(data).find("large").length > 0) {
                            $(data).find("large").each(function () {
                                var series = $.parseJSON($(this).children("series").text().trim());
                                piechart("LargeManage", "大型能源管理系統汰換申請數量 (套)", series);
                            });
                        }

                        if ($(data).find("dt_large").length > 0) {
                            var tabstr = getTable(data, "05", "dt_large");
                            $("#largelist").append(tabstr);
                        }
                    }
                }
            });
        }

        function piechart(tag, titletext, series) {
            Highcharts.chart(tag, {
                credits: {
                    enabled: false
                },
                chart: {
                    plotBackgroundColor: null,
                    plotBorderWidth: null,
                    plotShadow: false,
                    type: 'pie'
                },
                title: {
                    text: titletext
                },
                tooltip: {
                    pointFormat: '{series.name}: <b>{point.y}</b>'
                },
                plotOptions: {
                    pie: {
                        allowPointSelect: true,
                        cursor: 'pointer',
                        dataLabels: {
                            enabled: true,
                            format: '<b>{point.name}</b>: {point.percentage:.0f} %',
                            style: {
                                color: (Highcharts.theme && Highcharts.theme.contrastTextColor) || 'black'
                            }
                        }
                    }
                },
                series: [{
                    name: '數量',
                    colorByPoint: true,
                    data: series
                }]
            });

            //Highcharts 千分位
            //Highcharts.setOptions({
            //    lang: {
            //        thousandsSep: ','
            //    }
            //});
        }


        function getTable(data, type, xmltag) {
            var unitStr = "";
            switch (xmltag) {
                case "dt_air":
                    unitStr = " (台)";
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
            var tabstr = '<tr>';
            if (type == "03") {
                tabstr += '<th>項目</th><th>集合住宅</th><th>辦公大樓</th><th>服務業</th>';
            } else {
                tabstr += '<th>項目</th><th>機關</th><th>學校</th><th>服務業</th>';
            }
            tabstr += '</tr>';
            $(data).find(xmltag).each(function () {
                tabstr += '<tr>';
                tabstr += '<th>數量' + unitStr + '</th>';
                tabstr += '<td align="center">' + $(this).children("RM_SUM1_money").text().trim() + '</td><td align="center">' + $(this).children("RM_SUM2_money").text().trim() + '</td><td align="center">' + $(this).children("RM_SUM3_money").text().trim() +'</td>';
                tabstr += '</tr>';
                tabstr += '<tr>';
                tabstr += '<th>占比</th>';
                tabstr += '<td align="center">' + parseInt($(this).children("sumAvg").text().trim())+ '%</td><td align="center">' + parseInt($(this).children("sumAvg2").text().trim()) + '%</td><td align="center">' + parseInt($(this).children("sumAvg3").text().trim()) + '%</td>';
                tabstr += '</tr>';
            });
            return tabstr;
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
            <div class="left"><span class="filetitle font-size5">申請數量分析</span></div>
            <div class="right">附加圖表 / 地方圖表 / 申請數量分析</div>
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

