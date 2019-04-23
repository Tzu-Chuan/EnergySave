<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="total-behind.aspx.cs" Inherits="WebPage_total_behind" %>

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
                    func: "load_reportTotalBehind",
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
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function (i) {
                                    if ($(this).children("city_I_Guid").text().trim() != "") {
                                        strguid = $(this).children("city_I_Guid").text().trim();
                                    }
                                    strHtml += "<tr>";
                                    strHtml += "<td style='text-align:center;'>" + $(this).children("city_Item_cn").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:center;'>" + $(this).children("city_Year").text().trim() + "年<br />第" + $(this).children("city_Season").text().trim() + "季</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("RM_Sum01").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("RM_SumC01").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("RM_SumF01").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("RM_Sum02").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("RM_SumC02").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("RM_SumF02").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("RM_Sum03").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("RM_SumC03").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("RM_SumF03").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + parseInt($(this).children("RM_Sum04").text().trim()) + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + parseInt($(this).children("RM_SumC04").text().trim()) + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + parseInt($(this).children("RM_SumF04").text().trim()) + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + parseInt($(this).children("RM_Sum05").text().trim()) + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + parseInt($(this).children("RM_SumC05").text().trim()) + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + parseInt($(this).children("RM_SumF05").text().trim()) + "</td>";
                                    strHtml += "</tr>";
                                    //第一筆資料進來就直接加總"規劃數"，如果換其他縣市了再加總，不然規畫數會重複加到
                                    if (stroldguid == "" || (strguid != stroldguid)) {
                                        sum01 += parseFloat($(this).children("RM_Sum01").text().trim());
                                        sum02 += parseFloat($(this).children("RM_Sum02").text().trim());
                                        sum03 += parseFloat($(this).children("RM_Sum03").text().trim());
                                        sum04 += parseFloat($(this).children("RM_Sum04").text().trim());
                                        sum05 += parseFloat($(this).children("RM_Sum05").text().trim());
                                        stroldguid = strguid;
                                    }
                                    sumC01 += parseFloat($(this).children("RM_SumC01").text().trim());
                                    sumF01 += parseFloat($(this).children("RM_SumF01").text().trim());
                                    sumC02 += parseFloat($(this).children("RM_SumC02").text().trim());
                                    sumF02 += parseFloat($(this).children("RM_SumF02").text().trim());
                                    sumC03 += parseFloat($(this).children("RM_SumC03").text().trim());
                                    sumF03 += parseFloat($(this).children("RM_SumF03").text().trim());
                                    sumC04 += parseFloat($(this).children("RM_SumC04").text().trim());
                                    sumF04 += parseFloat($(this).children("RM_SumF04").text().trim());
                                    sumC05 += parseFloat($(this).children("RM_SumC05").text().trim());
                                    sumF05 += parseFloat($(this).children("RM_SumF05").text().trim());
                                    //sum01 += parseFloat($(this).children("RM_Sum01").text().trim());
                                    //sumC01 += parseFloat($(this).children("RM_SumC01").text().trim());
                                    //sumF01 += parseFloat($(this).children("RM_SumF01").text().trim());
                                    //sum02 += parseFloat($(this).children("RM_Sum02").text().trim());
                                    //sumC02 += parseFloat($(this).children("RM_SumC02").text().trim());
                                    //sumF02 += parseFloat($(this).children("RM_SumF02").text().trim());
                                    //sum03 += parseFloat($(this).children("RM_Sum03").text().trim());
                                    //sumC03 += parseFloat($(this).children("RM_SumC03").text().trim());
                                    //sumF03 += parseFloat($(this).children("RM_SumF03").text().trim());
                                    //sum04 += parseFloat($(this).children("RM_Sum04").text().trim());
                                    //sumC04 += parseFloat($(this).children("RM_SumC04").text().trim());
                                    //sumF04 += parseFloat($(this).children("RM_SumF04").text().trim());
                                    //sum05 += parseFloat($(this).children("RM_Sum05").text().trim());
                                    //sumC05 += parseFloat($(this).children("RM_SumC05").text().trim());
                                    //sumF05 += parseFloat($(this).children("RM_SumF05").text().trim());
                                    if (i == $(data).find("data_item").length - 1) {
                                        strHtml += "<tr>";
                                        strHtml += "<td style='text-align:center;'>合計</td>";
                                        strHtml += "<td></td>";
                                        strHtml += "<td style='text-align:right;'>" + parseFloat(sum01).toFixed(1) + "</td>";
                                        strHtml += "<td style='text-align:right;'>" + parseFloat(sumC01).toFixed(1) + "</td>";
                                        strHtml += "<td style='text-align:right;'>" + parseFloat(sumF01).toFixed(1) + "</td>";
                                        strHtml += "<td style='text-align:right;'>" + parseFloat(sum02).toFixed(1) + "</td>";
                                        strHtml += "<td style='text-align:right;'>" + parseFloat(sumC02).toFixed(1) + "</td>";
                                        strHtml += "<td style='text-align:right;'>" + parseFloat(sumF02).toFixed(1) + "</td>";
                                        strHtml += "<td style='text-align:right;'>" + parseFloat(sum03).toFixed(1) + "</td>";
                                        strHtml += "<td style='text-align:right;'>" + parseFloat(sumC03).toFixed(1) + "</td>";
                                        strHtml += "<td style='text-align:right;'>" + parseFloat(sumF03).toFixed(1) + "</td>";
                                        strHtml += "<td style='text-align:right;'>" + sum04 + "</td>";
                                        strHtml += "<td style='text-align:right;'>" + sumC04 + "</td>";
                                        strHtml += "<td style='text-align:right;'>" + sumF04 + "</td>";
                                        strHtml += "<td style='text-align:right;'>" + sum05 + "</td>";
                                        strHtml += "<td style='text-align:right;'>" + sumC05 + "</td>";
                                        strHtml += "<td style='text-align:right;'>" + sumF05 + "</td>";
                                        strHtml += "</tr>";
                                    }
                                });
                            }
                        } else {
                            strHtml += "<tr><td colspan='17'>尚無資料</td></tr>";
                        }
                        $("#tbody_list").empty().append(strHtml);
                    }
                }//success end
            });//ajax end
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">各縣市申請數</span></div>
            <div class="right">附加圖表 / 管理員總覽表 / 各縣市申請數</div>
        </div>

        期別：
        <select id="ddlStage" class="inputex" style="margin-top:10px;">
            <option value="1"  selected="selected">第一期</option>
            <option value="2">第二期</option>
            <option value="3">第三期</option>
        </select>
        &nbsp;&nbsp;<a href="javascript:void(0);" class="genbtn" id="btn_export">匯出Excel</a><br /><br />
        <div class="stripeMe">
            <table id="airlist" border="0" cellspacing="0" cellpadding="0" width="100%">
                <thead>
                    <tr>
                        <th rowspan="2">縣市</th>
                        <th rowspan="2">季</th>
                        <th colspan="3">無風管冷氣(KW)</th>
                        <th colspan="3">老舊辦公室照明(具)</th>
                        <th colspan="3">室內停車場<br />智慧照明(盞)</th>
                        <th colspan="3">中型能管系統(套)</th>
                        <th colspan="3">大型能管系統(套)</th>
                    </tr>
                    <tr>
                        <th>規劃數</th>
                        <th>申請數</th>
                        <th>完成數</th>
                        <th>規劃數</th>
                        <th>申請數</th>
                        <th>完成數</th>
                        <th>規劃數</th>
                        <th>申請數</th>
                        <th>完成數</th>
                        <th>規劃數</th>
                        <th>申請數</th>
                        <th>完成數</th>
                        <th>規劃數</th>
                        <th>申請數</th>
                        <th>完成數</th>
                    </tr>
                </thead>
                <tbody id="tbody_list"></tbody>
            </table>
        </div>
    </div><br />
</asp:Content>

