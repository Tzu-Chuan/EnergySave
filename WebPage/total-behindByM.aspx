<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="total-behindByM.aspx.cs" Inherits="WebPage_total_behindByM" %>

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
                window.location = "../handler/ExportTotalBehindByM.aspx?s=" + encodeURIComponent($("#ddlStage").val()) + "";
            });
        });
        function loadData(){
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/accReport.ashx",
                data: {
                    func: "load_reportTotalBehindByM",
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
                        var I1 = "", I2 = "", I3 = "", I4 = "", I5 = "";
                        var I1C = "", I2C = "", I3C = "", I4C = "", I5C = "";
                        var I1F = "", I2F = "", I3F = "", I4F = "", I5F = "";
                        if (data != null && data !="") {
                            data = $.parseXML(data);
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function (i) {
                                    strHtml += "<tr>";
                                    strHtml += "<td style='text-align:center;'>" + $(this).children("city_Item_cn").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:center;'>" + $(this).children("city_Year").text().trim() + "/" + $(this).children("city_Month").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("I_Finish_item1").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("sumC_item1").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("sumF_item1").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("I_Finish_item2").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("sumC_item2").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("sumF_item2").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("I_Finish_item3").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("sumC_item3").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("sumF_item3").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("I_Finish_item4").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("sumC_item4").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("sumF_item4").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("I_Finish_item5").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("sumC_item5").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:right;'>" + $(this).children("sumF_item5").text().trim() + "</td>";
                                        
                                    //第一筆資料進來就直接加總"規劃數"，如果換其他縣市了再加總，不然規畫數會重複加到
                                    sum01 += chkNUM($(this).children("I_Finish_item1").text().trim());
                                    sum02 += chkNUM($(this).children("I_Finish_item2").text().trim());
                                    sum03 += chkNUM($(this).children("I_Finish_item3").text().trim());
                                    sum04 += chkNUM($(this).children("I_Finish_item4").text().trim());
                                    sum05 += chkNUM($(this).children("I_Finish_item5").text().trim());
                                    sumC01 += chkNUM($(this).children("sumC_item1").text().trim());
                                    sumF01 += chkNUM($(this).children("sumF_item1").text().trim());
                                    sumC02 += chkNUM($(this).children("sumC_item2").text().trim());
                                    sumF02 += chkNUM($(this).children("sumF_item2").text().trim());
                                    sumC03 += chkNUM($(this).children("sumC_item3").text().trim());
                                    sumF03 += chkNUM($(this).children("sumF_item3").text().trim());
                                    sumC04 += chkNUM($(this).children("sumC_item4").text().trim());
                                    sumF04 += chkNUM($(this).children("sumF_item4").text().trim());
                                    sumC05 += chkNUM($(this).children("sumC_item5").text().trim());
                                    sumF05 += chkNUM($(this).children("sumF_item5").text().trim());
                                    strHtml += "</tr>";

                                    if (i == $(data).find("data_item").length-1) {
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

        function chkNUM(val) {
            if (val == "") {
                return 0;
            }
            else {
                return parseFloat(val);
            }
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">設備汰換 - 各縣市申請數</span></div>
            <div class="right">附加圖表 / 管理員總覽表 / 設備汰換 - 各縣市申請數</div>
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
                        <th rowspan="2" style="width:7%">縣市</th>
                        <th rowspan="2" style="width:8%">年/月</th>
                        <th colspan="3" style="width:17%;">無風管冷氣(KW)</th>
                        <th colspan="3" style="width:17%;">老舊辦公室照明(具)</th>
                        <th colspan="3" style="width:17%;">室內停車場<br />智慧照明(盞)</th>
                        <th colspan="3" style="width:17%;">中型能管系統(套)</th>
                        <th colspan="3" style="width:17%;">大型能管系統(套)</th>
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


