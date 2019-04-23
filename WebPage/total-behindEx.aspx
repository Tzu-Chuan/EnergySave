<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="total-behindEx.aspx.cs" Inherits="WebPage_total_behindEx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script type="text/javascript">
        $(function () {
            loadData();
            //期數下拉選單change
            $("#ddlStage").change(function () {
                loadData();
            });
            //匯出按鈕
            $("#btn_export").click(function () {
                window.location = "../handler/ExportTotalBehind.aspx?s=" + encodeURIComponent($("#ddlStage").val()) + "";
            });
        });
        function loadData(){
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/accReport.ashx",
                data: {
                    func: "load_reportTotalBehindForEx",
                    strStage: $("#ddlStage").val()
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if (data.indexOf("Error") > -1) {
                        alert(data);
                    } else {
                        if (data == "reLogin") {
                            alert("請重新登入");
                            window.location = "Login.aspx";
                            return;
                        }
                        if (data == "noacc") {
                            alert("您沒有權限進入本頁面");
                            window.location = "ProjectList.aspx";
                            return;
                        }
                        var strHtml = "";
                        var sum01 = 0.0, sumC01 = 0.0, sumF01 = 0.0;
                        var sum02 = 0.0, sumC02 = 0.0, sumF02 = 0.0;
                        var sum03 = 0.0, sumC03 = 0.0, sumF03 = 0.0;
                        var sum04 = 0, sumC04 = 0, sumF04 = 0;
                        var sum05 = 0, sumC05 = 0, sumF05 = 0;
                        var strguid = "", stroldguid = "";
                        if (data != null && data != "") {
                            data = $.parseXML(data);
                            //************************表頭************************//
                            if ($(data).find("dataHead").length > 0) {
                                strHtml += '<tr>';
                                strHtml += '<th rowspan="2">縣市</th>';
                                strHtml += '<th rowspan="2">季</th>';
                                $(data).find("dataHead").each(function (i) {
                                    strHtml += '<th colspan="3">' + $(this).children("C_Item_cn").text().trim() + '</th>';
                                });
                                strHtml += '</tr>';
                                strHtml += '<tr>';
                                for (var i = 0; i < $(data).find("dataHead").length; i++) {
                                    strHtml += '<th>規劃數</th>';
                                    strHtml += '<th>申請數</th>';
                                    strHtml += '<th>完成數</th>';
                                }
                                strHtml += '</tr>';
                            }
                            //************************END************************//

                            //************************縣市資料************************//
                            var tabstr = "";
                            var tmpYear = "";
                            var tmpSeason = "";
                            var tmpStage = "";
                            var tmpType = "";
                            //$(data).find("dataCity").each(function (i) {
                            //    tmpYear = $(this).children("city_Year").text().trim();
                            //    tmpSeason = $(this).children("city_Season").text().trim();
                            //    tmpStage = $(this).children("city_Stage").text().trim();
                            //    tabstr += "<tr>";
                            //    tabstr += '<td>' + $(this).children("city_Item_cn").text().trim() + '</td>';
                            //    tabstr += '<td>' + $(this).children("city_Year").text().trim() + '年第' + $(this).children("city_Season").text().trim() + '季</td>';
                            //    $(data).find("dataHead").each(function (j) {
                            //        tmpType = $(this).children("C_Item").text().trim();
                            //        $(data).find("data_item").each(function (j) {
                            //            if (tmpYear == $(this).children("RM_Year").text().trim() &&
                            //                tmpSeason == $(this).children("RM_Season").text().trim() &&
                            //                tmpStage== $(this).children("RM_Stage").text().trim() &&
                            //                tmpType == $(this).children("RM_CPType").text().trim()) {
                            //                //01 & 03  申請、完成 抓 sum3、sum4  其他的申請、完成 抓 sum1、sum2
                            //                switch (tmpType) {
                            //                    case "01":
                            //                    case "03":
                            //                        tabstr += '<td>' + $(this).children("RM_Planning").text().trim() + '</td>';
                            //                        tabstr += '<td>' + $(this).children("sum3").text().trim() + '</td>';
                            //                        tabstr += '<td>' + $(this).children("sum4").text().trim() + '</td>';
                            //                        break;
                            //                    default:
                            //                        tabstr += '<td>' + $(this).children("RM_Planning").text().trim() + '</td>';
                            //                        tabstr += '<td>' + $(this).children("sum1").text().trim() + '</td>';
                            //                        tabstr += '<td>' + $(this).children("sum2").text().trim() + '</td>';
                            //                        break;
                            //                }
                            //            }
                            //            else {
                            //                tabstr += '<td></td>';
                            //                tabstr += '<td></td>';
                            //                tabstr += '<td></td>';

                            //            }

                            //        });
                            //    });
                            //    tabstr += "</tr>";
                            //});
                            //************************END************************//
                        }
                        
                        else {
                            strHtml += "<tr><td colspan='17'>尚無資料</td></tr>";
                        }
                        $("#thead_list").empty().append(strHtml + tabstr);
                    }
                }//success end
            });//ajax end
        }

    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">各縣市申請數(擴大補助)</span></div>
            <div class="right">附加圖表 / 管理員總覽表 / 各縣市申請數(擴大補助)</div>
        </div>

        期別：
        <select id="ddlStage" class="inputex" style="margin-top:10px;">
            <option value="1"  selected="selected">第一期</option>
            <option value="2">第二期</option>
            <option value="3">第三期</option>
        </select>
        類別：
        <select id="ddlType" class="inputex" style="margin-top:10px;">
            <option value="1"  selected="selected">服務業(機關、學校)</option>
            <option value="2">住宅</option>
        </select>
        &nbsp;&nbsp;<a href="javascript:void(0);" class="genbtn" id="btn_export">匯出Excel</a><br /><br />
        <div class="stripeMe" style="overflow-x:scroll;width:970px;">
            <table id="airlist" border="0" cellspacing="0" cellpadding="0" width="100%">
                <thead id="thead_list"></thead>
                <tbody id="tbody_list"></tbody>
            </table>
        </div>
    </div><br />
</asp:Content>


