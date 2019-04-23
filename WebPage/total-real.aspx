<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="total-real.aspx.cs" Inherits="WebPage_total_real" %>

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
                window.location = "../handler/ExportTotalReal.aspx?s=" + encodeURIComponent($("#ddlStage").val()) + "";
            });
        });
        function loadData(){
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/accReport.ashx",
                data: {
                    func: "load_reportReal",
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
                        var sumM = 0, sumR = 0;
                        var sum03RR, sum01RR = "",sum02RR = "";
                        var sM = "", sR = "";
                        if (data != null && data!="") {
                            data = $.parseXML(data);
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function () {
                                    sum01RR = $(this).children("RS_Type01RealRate").text().trim();
                                    sum02RR = $(this).children("RS_Type02RealRate").text().trim();
                                    sum03RR = $(this).children("RS_Type03RealRate").text().trim();
                                    sum04RR = $(this).children("RS_Type04RealRate").text().trim();
                                    if ($(this).children("RS_Type01Money").text().trim() != "") {
                                        sumM += parseFloat($(this).children("RS_Type01Money").text().trim());
                                        sM = sumM.toString();
                                    }
                                    if ($(this).children("RS_Type02Money").text().trim() != "") {
                                        sumM += parseFloat($(this).children("RS_Type02Money").text().trim());
                                        sM = sumM.toString();
                                    }
                                    if ($(this).children("RS_Type03Money").text().trim() != "") {
                                        sumM += parseFloat($(this).children("RS_Type03Money").text().trim());
                                        sM = sumM.toString();
                                    }
                                    if ($(this).children("RS_Type04Money").text().trim() != "") {
                                        sumM += parseFloat($(this).children("RS_Type04Money").text().trim());
                                        sM = sumM.toString();
                                    }
                                    if ($(this).children("RS_Type01Real").text().trim() != "") {
                                        sumR += parseFloat($(this).children("RS_Type01Real").text().trim());
                                        sR = sumR.toString();
                                    }
                                    if ($(this).children("RS_Type02Real").text().trim() != "") {
                                        sumR += parseFloat($(this).children("RS_Type02Real").text().trim());
                                        sR = sumR.toString();
                                    }
                                    if ($(this).children("RS_Type03Real").text().trim() != "") {
                                        sumR += parseFloat($(this).children("RS_Type03Real").text().trim());
                                        sR = sumR.toString();
                                    }
                                    if ($(this).children("RS_Type04Real").text().trim() != "") {
                                        sumR += parseFloat($(this).children("RS_Type04Real").text().trim());
                                        sR = sumR.toString();
                                    }
                                    strHtml += "<tr>";
                                    strHtml += "<td style='text-align:center;'>" + $(this).children("citycn").text().trim() + "</td>";
                                    strHtml += "<td style='text-align:center;'>" + $(this).children("RS_Year").text().trim() + "年<br />第" + $(this).children("RS_Season").text().trim() + "季</td>";
                                    strHtml += "<td style='text-align:center;'>" + splitStr($(this).children("RS_CostDesc").text().trim()) + "</td>";
                                    strHtml += "<td style='text-align:center;'>" + sum01RR + "</td>";
                                    strHtml += "<td style='text-align:center;'>" + sum02RR + "</td>";
                                    strHtml += "<td style='text-align:center;'>" + sum03RR + "</td>";
                                    strHtml += "<td style='text-align:center;'>" + sum04RR + "</td>";
                                    //分母不為0
                                    if ( parseFloat(sM).toFixed(0) > 0) {
                                        strHtml += "<td style='text-align:right;'>" + ((parseFloat(sR) / parseFloat(sM))*100).toFixed(0)+ "</td>";
                                    } else {
                                        strHtml += "<td style='text-align:right;'>0</td>";
                                    }
                                    strHtml += "</tr>";
                                    sum01RR = "";
                                    sum02RR = "";
                                    sum03RR = "";
                                    sM = "";
                                    sR = "";
                                    sumM = 0;
                                    sumR = 0;
                                });
                            }
                        } else {
                            strHtml += "<tr><td colspan='8'>尚無資料</td></tr>";
                        }
                        $("#tbody_list").empty().append(strHtml);
                    }
                }//success end
            });//ajax end
    }

    //將XML回傳回來的遭遇困難字串拆開 然後塞換行符號
        function splitStr(str) {
            var strVal = "";//
            var splitVal = "";//split的字串 \n
            if (str != null && str.trim() != "") {
                splitVal = str.split("\n");
                for (var i = 0; i < splitVal.length; i++) {
                    if (splitVal[i].toString().trim()!="") {
                        if (strVal == "") {
                            strVal += splitVal[i].toString().trim();
                        } else {
                            strVal += "<br />" + splitVal[i].toString().trim();
                        }
                    }
                }
            }
            return strVal;
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">工作經費狀態及實支率</span></div>
            <div class="right">附加圖表 / 管理員總覽表 / 工作經費狀態及實支率</div>
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
                        <th rowspan='2'>預算狀態</th>
                        <th colspan='5'>經費實支率%</th>
                    </tr>
                    <tr>
                        <th>節電基礎</th>
                        <th>因地制宜</th>
                        <th>設備汰換</th>
                        <th>擴大補助</th>
                        <th>整體</th>
                    </tr>
                </thead>
                <tbody id="tbody_list"></tbody>
            </table>
        </div>
    </div><br />
</asp:Content>

