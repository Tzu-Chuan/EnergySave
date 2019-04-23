<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="total-process.aspx.cs" Inherits="WebPage_total_process" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script type="text/javascript">
        $(function () {
            loadData();
            //期數下拉選單change
            $("#ddlStage").change(function () {
                loadData();
            });
            //匯出按鈕
            $("#btn_export").click(function () {
                window.location = "../handler/ExportTotalProcess.aspx?s=" + encodeURIComponent($("#ddlStage").val()) + "";
            });
        });
        function loadData() {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/accReport.ashx",
                data: {
                    func: "load_reportProcess",
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
                        var sum01S = "", sum01F = "", sum01SF = "";
                        var sum02S = "", sum02F = "", sum02SF = "";
                        var sum03S = "", sum03F = "", sum03SF = "";
                        var sum04S = "", sum04F = "", sum04SF = "";
                        if (data != null && data != "") {
                            data = $.parseXML(data);
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function () {
                                    if ($(this).children("RS_Year").text().trim() != "" && $(this).children("RS_Season").text().trim() != "") {
                                        strHtml += "<tr>";
                                        strHtml += "<td style='text-align:center;'>" + $(this).children("citycn").text().trim() + "</td>";
                                        strHtml += "<td style='text-align:center;'>" + $(this).children("RS_Year").text().trim() + "年<br />第" + $(this).children("RS_Season").text().trim() + "季</td>";
                                        //如果沒有預定的值 就顯示空白
                                        sum01S = $(this).children("RS_Sum01S").text().trim();
                                        sum02S = $(this).children("RS_Sum02S").text().trim();
                                        sum03S = $(this).children("RS_Sum03S").text().trim();
                                        sum04S = $(this).children("RS_Sum04S").text().trim();
                                        sum01F = $(this).children("RS_Sum01F").text().trim();
                                        sum02F = $(this).children("RS_Sum02F").text().trim();
                                        sum03F = $(this).children("RS_Sum03F").text().trim();
                                        sum04F = $(this).children("RS_Sum04F").text().trim();
                                        sum01SF = $(this).children("RS_Sum01S_F").text().trim();
                                        sum02SF = $(this).children("RS_Sum02S_F").text().trim();
                                        sum03SF = $(this).children("RS_Sum03S_F").text().trim();
                                        sum04SF = $(this).children("RS_Sum04S_F").text().trim();

                                        if (sum01S == "0" || sum01S == "0.0" || sum01S == "0.00") {//
                                            strHtml += "<td></td>";
                                            strHtml += "<td></td>";
                                            strHtml += "<td></td>";
                                        } else {
                                            strHtml += "<td style='text-align:right;'>" + sum01S + "</td>";
                                            strHtml += "<td style='text-align:right;'>" + sum01F + "</td>";
                                            strHtml += "<td style='text-align:right;'>" + sum01SF + "</td>";
                                        }
                                        if (sum02S == "0" || sum02S == "0.0" || sum02S == "0.00") {//
                                            strHtml += "<td></td>";
                                            strHtml += "<td></td>";
                                            strHtml += "<td></td>";
                                        } else {
                                            strHtml += "<td style='text-align:right;'>" + sum02S + "</td>";
                                            strHtml += "<td style='text-align:right;'>" + sum02F + "</td>";
                                            strHtml += "<td style='text-align:right;'>" + sum02SF + "</td>";
                                        }
                                        if (sum03S == "0" || sum03S == "0.0" || sum03S == "0.00") {// 
                                            strHtml += "<td></td>";
                                            strHtml += "<td></td>";
                                            strHtml += "<td></td>";
                                        } else {
                                            strHtml += "<td style='text-align:right;'>" + sum03S + "</td>";
                                            strHtml += "<td style='text-align:right;'>" + sum03F + "</td>";
                                            strHtml += "<td style='text-align:right;'>" + sum03SF + "</td>";
                                        }
                                        if (sum04S == "0" || sum04S == "0.0" || sum04S == "0.00") {// 
                                            strHtml += "<td></td>";
                                            strHtml += "<td></td>";
                                            strHtml += "<td></td>";
                                        } else {
                                            strHtml += "<td style='text-align:right;'>" + sum04S + "</td>";
                                            strHtml += "<td style='text-align:right;'>" + sum04F + "</td>";
                                            strHtml += "<td style='text-align:right;'>" + sum04SF + "</td>";
                                        }
                                        //算整體的時候要看前面三項，有幾項就除以幾 (沒執行的項目要空白或變灰色)
                                        if ((sum01S == "0" || sum01S == "0.0" || sum01S == "0.00") && (sum02S == "0" || sum02S == "0.0" || sum02S == "0.00") && (sum03S == "0" || sum03S == "0.0" || sum03S == "0.00") && (sum04S == "0" || sum04S == "0.0" || sum04S == "0.00")) {
                                            strHtml += "<td></td>";
                                            strHtml += "<td></td>";
                                            strHtml += "<td></td>";
                                        } else {
                                            var intall = 0;
                                            if (sum01S != "0") {
                                                intall = intall + 1;
                                            }
                                            if (sum02S != "0") {
                                                intall = intall + 1;
                                            }
                                            if (sum03S != "0") {
                                                intall = intall + 1;
                                            }
                                            if (sum04S != "0") {
                                                intall = intall + 1;
                                            }
                                            strHtml += "<td style='text-align:right;'>" + parseFloat((parseInt(sum01S) + parseInt(sum02S) + parseInt(sum03S) + parseInt(sum04S)) / intall).toFixed(2) + "</td>";
                                            strHtml += "<td style='text-align:right;'>" + parseFloat((parseInt(sum01F) + parseInt(sum02F) + parseInt(sum03F) + parseInt(sum04S)) / intall).toFixed(2) + "</td>";
                                            //parseFloat((parseFloat(parseFloat((parseInt(sum01S) + parseInt(sum02S) + parseInt(sum03S)) / intall).toFixed(2))-parseFloat(parseFloat((parseInt(sum01F) + parseInt(sum02F) + parseInt(sum03F)) / intall).toFixed(2)))
                                            strHtml += "<td style='text-align:right;'>" + parseFloat(parseFloat(parseFloat((parseInt(sum01S) + parseInt(sum02S) + parseInt(sum03S) + parseInt(sum04S)) / intall).toFixed(2))-parseFloat(parseFloat((parseInt(sum01F) + parseInt(sum02F) + parseInt(sum03F) + parseInt(sum04F)) / intall).toFixed(2))).toFixed(2) + "</td>";
                                        }
                                        strHtml += "</tr>";
                                    }
                                });
                            }
                            
                        } else {
                            strHtml += "<tr><td colspan='14'>尚無資料</td></tr>";
                        }
                        $("#tbody_list").empty().append(strHtml);
                    }
                }//success end
            });//ajax end
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">當期累計執行進度</span></div>
            <div class="right">附加圖表 / 管理員總覽表 / 當期累計執行進度</div>
        </div>

        期別：
       
        <select id="ddlStage" class="inputex" style="margin-top: 10px;">
            <option value="1" selected="selected">第一期</option>
            <option value="2">第二期</option>
            <option value="3">第三期</option>
        </select>
        &nbsp;&nbsp;<a href="javascript:void(0);" class="genbtn" id="btn_export">匯出Excel</a><br />
        <br />
        <div class="stripeMe">
            <table id="airlist" border="0" cellspacing="0" cellpadding="0" width="100%">
                <thead>
                    <tr>
                        <th rowspan="2">縣市</th>
                        <th rowspan="2">季</th>
                        <th colspan='3'>節電基礎工作%</th>
                        <th colspan='3'>因地制宜%</th>
                        <th colspan='3'>設備汰換與智慧用電%</th>
                        <th colspan='3'>擴大補助%</th>
                        <th colspan='3'>整體%</th>
                    </tr>
                    <tr>
                        <th>預定</th>
                        <th>實際</th>
                        <th>差異</th>
                        <th>預定</th>
                        <th>實際</th>
                        <th>差異</th>
                        <th>預定</th>
                        <th>實際</th>
                        <th>差異</th>
                        <th>預定</th>
                        <th>實際</th>
                        <th>差異</th>
                        <th>預定</th>
                        <th>實際</th>
                        <th>差異</th>
                    </tr>
                </thead>
                <tbody id="tbody_list"></tbody>
            </table>
        </div>
    </div>
    <br />
</asp:Content>


