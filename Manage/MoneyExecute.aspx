<%@ Page Title="" Language="C#" MasterPageFile="~/Manage/Admin.master" AutoEventWireup="true" CodeFile="MoneyExecute.aspx.cs" Inherits="Manage_MoneyExecute" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            getddl("02", "#ddlCity");
            getData();

            //期數&機關 下拉選單change
            $("#ddlCity,#ddlStage").change(function () {
                getData();
            });

            //匯出按鈕
            $("#exportbtn").click(function () {
                window.location = "../handler/ExportMoneyExecuteManage.aspx?s=" + encodeURIComponent($("#ddlStage").val()) + "&c=" + encodeURIComponent($("#ddlCity").val()) + "";
            });

        });//end js
        
        //取得經費執行列表&經費
        function getData() {
            $.ajax({
                type: "POST",
                async: false,
                url: "../handler/getMoneyExecute_Manager.aspx",
                data: {
                    stage: $("#ddlStage").val(),
                    city: $("#ddlCity").val()
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if ($(data).find("Error").length > 0) {
                        alert($(data).find("Error").attr("Message"));
                        return;
                    }
                    if (data == "reLogin") {
                        alert("請重新登入");
                        window.location = "../WebPage/Login.aspx";
                        return;
                    }
                    $("#tablist tbody").empty();
                    var tabstr = '', sumA = 0, sumB = 0, sumC = 0;
                    if ($(data).find("data_item").length > 0) {
                        $(data).find("data_item").each(function (i) {
                            tabstr += '<tr>';
                            tabstr += '<td align="center" nowrap="nowrap">' + (i + 1) + '</td>';
                            tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("CityName").text().trim() + '</td>';
                            tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("PR_Office").text().trim() + '</td>';
                            tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("planNamer").text().trim() + '</td>';
                            tabstr += '<td align="right" nowrap="nowrap">' + $(this).children("PR_Money_m").text().trim() + '</td>';
                            tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("PR_CaseName").text().trim() + '</td>';
                            tabstr += '<td align="right" nowrap="nowrap">' + $(this).children("PR_CaseMoney_m").text().trim() + '</td>';
                            tabstr += '<td align="right" nowrap="nowrap">' + $(this).children("PR_SelfMoney_m").text().trim() + '</td>';
                            tabstr += '<td align="center" nowrap="nowrap">' + splitStr($(this).children("PR_Steps").text().trim()) + '</td>';
                            tabstr += '</tr>';
                            if ($(this).children("PR_Money").text().trim() != "") {
                                sumA += parseInt($(this).children("PR_Money").text().trim());
                            }
                            if ($(this).children("PR_CaseMoney").text().trim() != "") {
                                sumB += parseInt($(this).children("PR_CaseMoney").text().trim());
                            }
                            if ($(this).children("PR_SelfMoney").text().trim() != "") {
                                sumC += parseInt($(this).children("PR_SelfMoney").text().trim());
                            }
                            if (i == $(data).find("data_item").length - 1) {
                                tabstr += '<tr>';
                                tabstr += '<td align="right" colspan="4">合計</td>';
                                tabstr += '<td align="right">' + FormatNumber(sumA) + '</td>';
                                tabstr += '<td></td>';
                                tabstr += '<td align="right">' + FormatNumber(sumB) + '</td>';
                                tabstr += '<td align="right">' + FormatNumber(sumC) + '</td>';
                                tabstr += '<td></td>';
                                tabstr += '</tr>';
                            }
                        });

                        //經費：
                        if ($(data).find("data_money").length > 0 && $("#ddlCity").val() != "") {
                            $(data).find("data_money").each(function (i) {
                                $("#div_money").empty().append("經費：" + FormatNumber(parseFloat($(this).children("MoneyAll").text().trim()) * 1000).toString() + "(元)");
                            });
                        }
                    }
                    else {
                        tabstr += "<tr><td colspan='9'>查詢無資料</td></tr>";
                    }
                    $("#tablist tbody").append(tabstr);
                    $(".stripeMe tr").mouseover(function () { $(this).addClass("spe"); }).mouseout(function () { $(this).removeClass("spe"); });
                    $(".stripeMe table:not(td > table) > tbody > tr:not('.spe'):even").addClass("alt");

                }
            });
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
                        var ddlstr = '<option value="">全部</option>';
                        if ($(data).find("code").length > 0) {
                            $(data).find("code").each(function (i) {
                                ddlstr += '<option value="' + $(this).attr("v") + '">' + $(this).attr("desc") + '</option>';
                            });
                        }
                        $(tagName).append(ddlstr);
                    }
                }
            });
        }

        //替換\n => <BR>
        function splitStr(str) {
            var strVal = "";///回傳回去的字串
            var splitVal = "";//split的字串 \n
            if (str != null && str.trim() != "") {
                splitVal = str.split("\n");
                for (var i = 0; i < splitVal.length; i++) {
                    if (splitVal[i].toString().trim() != "") {
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

        //加入千分號
        function FormatNumber(n) {
            n += "";
            var arr = n.split(".");
            var re = /(\d{1,3})(?=(\d{3})+$)/g;
            return arr[0].replace(re, "$1,") + (arr.length == 2 ? "." + arr[1] : "");
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">經費執行</span></div>
            <div class="right"></div>
        </div>

        <div class="margin10T">
            縣市：<select id="ddlCity" class="inputex" style="margin-right: 10px;"></select>
            期：<select id="ddlStage" class="inputex" style="margin-right: 10px;">
                <option value="1">第一期</option>
                <option value="2">第二期</option>
                <option value="3">第三期</option>
            </select>
        </div>
        <div id="div_money" class="margin10T" style="display: mone;"></div>
        <div style="text-align: right;">
            <input class="genbtn" type="button" id="exportbtn" value="匯出" />
        </div>
        <div class="stripeMe margin10T font-normal">
            <table id="tablist" width="100%" border="0" cellspacing="0" cellpadding="0">
                <thead>
                    <tr>
                        <th nowrap="nowrap" rowspan="3" style="width: 50px;">編號</th>
                        <th nowrap="nowrap" rowspan="3">執行機關</th>
                        <th nowrap="nowrap" rowspan="3">主責局處</th>
                        <th nowrap="nowrap" rowspan="3">計畫項目</th>
                        <th nowrap="nowrap" rowspan="3" style="width: 150px;">金額(元)<br />
                            (A)=(B)+(C)</th>
                        <th nowrap="nowrap" colspan="3">處理方式</th>
                        <th nowrap="nowrap" rowspan="3" style="width: 200px;">涉及措施</th>
                    </tr>
                    <tr>
                        <th nowrap="nowrap" colspan="2">標案</th>
                        <th nowrap="nowrap">自辦</th>
                    </tr>
                    <tr>
                        <th nowrap="nowrap">標案名稱</th>
                        <th nowrap="nowrap" style="width: 100px;">發包金額(元)(B)</th>
                        <th nowrap="nowrap" style="width: 100px;">金額(元)(C)</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
        </div>
    </div>
</asp:Content>

