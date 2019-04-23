<%@ Page Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="MoneyExecution.aspx.cs" Inherits="WebPage_MoneyExecution" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script type="text/javascript">
        var btnStatus = "";// new 新增 mod 修改
        $(document).ready(function () {
            getddl("08", "#PR_PlanTitle");
            getData();

            //限制只能輸入數字
            $(document).on("keyup", ".num", function () {
                if (/[^0-9\.]/g.test(this.value)) {
                    this.value = this.value.replace(/[^0-9\.]/g, '');
                }
            });

            //新增 button
            $(document).on("click", "#addbtn", function () {
                $("#hidden_id").val($(this).attr("aid"));
                $("#PR_Stage").val("1");
                $("#PR_PlanTitle").val($("#PR_PlanTitle option:first").val());
                $("#PR_Office").val("");
                $("#PR_Money").val("");
                $("#PR_Steps").val("");
                $("#PR_CaseName").val("");
                $("#PR_CaseMoney").val("");
                $("#PR_SelfMoney").val("");
                btnStatus = "new";
                doShowDialog();
            });

            $(document).on("change", "#PR_Method", function () {
                if (this.value == "99") {
                    $("#PR_MethodDesc").show();
                } else {
                    $("#PR_MethodDesc").hide();
                }
            });

            //確定(儲存)按鈕
            $(document).on("click", "#subbtn", function () {
                if (chkData()) {
                    moddata();
                }
            });

            //修改按鈕
            $(document).on("click", "input[name='modbtn']", function () {
                $("#hidden_id").val($(this).attr("aid"));
                btnStatus = "mod";
                getDataById($(this).attr("aid"));
                doShowDialog();
            });

            //刪除 button
            $(document).on("click", "input[name='delbtn']", function () {
                var id = $(this).attr("aid");
                if (confirm("確定刪除？")) {
                    $.ajax({
                        type: "POST",
                        async: false, //在沒有返回值之前,不會執行下一步動作
                        url: "../handler/delMoney_Execution.aspx",
                        data: {
                            PR_ID: id
                        },
                        error: function (xhr) {
                            alert("Error " + xhr.status);
                            console.log(xhr.responseText);
                        },
                        success: function (data) {
                            var strMsg = "";
                            if (data.indexOf("Error") > -1) {
                                alert(data);
                                return;
                            }
                            if (data == "reLogin") {
                                alert("請重新登入");
                                window.location = ".Login.aspx";
                                return;
                            }  
                            if (data == "success") {
                                alert("刪除成功");
                                getData();
                            }
                        }
                    });
                }
            });

            //取消 fancybox button
            $(document).on("click", "#cancelbtn", function () {
                if (confirm('資料尚未儲存，確定取消？'))
                    $("#NewBlock").dialog("close");
            });

            //期數 下拉選單change事件
            $("#n_Stage").change(function () {
                getData();
            });

        });

        //取得經費執行列表&經費
        function getData() {
            $.ajax({
                type: "POST",
                async: false,
                url: "../handler/getMoney_Execution.aspx",
                data: {
                    stage: $("#n_Stage").val()
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
                            tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("PR_Office").text().trim() + '</td>';
                            tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("C_Item_cn").text().trim() + '</td>';
                            tabstr += '<td align="right" nowrap="nowrap">' + $(this).children("PR_Money_m").text().trim() + '</td>';
                            tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("PR_CaseName").text().trim() + '</td>';
                            tabstr += '<td align="right" nowrap="nowrap">' + $(this).children("PR_CaseMoney_m").text().trim() + '</td>';
                            tabstr += '<td align="right" nowrap="nowrap">' + $(this).children("PR_SelfMoney_m").text().trim() + '</td>';
                            tabstr += '<td align="center" nowrap="nowrap">' + splitStr($(this).children("PR_Steps").text().trim()) + '</td>';
                            tabstr += '<td align="center" style="width:15%"><input type="button" class="genbtn" name="modbtn" aid="' + $(this).children("PR_ID").text().trim() + '" value="修改" />';
                            tabstr += '<input type="button" class="genbtn" name= "delbtn" aid= "' + $(this).children("PR_ID").text().trim() + '" value="刪除" /></td >';
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
                                tabstr += '<td align="right" colspan="3">合計</td>';
                                tabstr += '<td align="right">' + FormatNumber(sumA) + '</td>';
                                tabstr += '<td></td>';
                                tabstr += '<td align="right">' + FormatNumber(sumB) + '</td>';
                                tabstr += '<td align="right">' + FormatNumber(sumC) + '</td>';
                                tabstr += '<td></td>';
                                tabstr += '<td></td>';
                                tabstr += '</tr>';
                            }
                        });
                    }
                    else {
                        tabstr += "<tr><td colspan='9'>查詢無資料</td></tr>";
                    }

                    if ($(data).find("total").length > 0 && $(data).find("total").text().trim() != "") {
                        $("#n_money").empty().append(FormatNumber(parseFloat($(data).find("total").text().trim()) * 1000) + "(元)");
                    } else {
                        $("#n_money").empty().append("");
                    }
                    if ($(data).find("cityName").length > 0 && $(data).find("cityName").text().trim() != "") {
                        $("#n_cityName").empty().append($(data).find("cityName").text().trim());
                        $("#PR_CityName").empty().append($(data).find("cityName").text().trim());
                    }
                    if ($(data).find("cityName").length > 0 && $(data).find("cityName").text().trim() != "") {
                        $("#PR_City").val($(data).find("cityno").text().trim());
                    }

                    $("#tablist tbody").append(tabstr);
                    $(".stripeMe tr").mouseover(function () { $(this).addClass("spe"); }).mouseout(function () { $(this).removeClass("spe"); });
                    $(".stripeMe table:not(td > table) > tbody > tr:not('.spe'):even").addClass("alt");
                    
                }
            });
        }

        //新增/修改
        function moddata() {
            //PR_CaseName 標案名稱 PR_CaseMoney 標案：發包金額(B) PR_SelfMoney 自辦：金額(C)
            //PR_Stage 期數 PR_City 執行機關(代碼) PR_CityName 執行機關(名稱)  PR_PlanTitle 計畫名稱 PR_Office 主責局處 PR_Money 金額 PR_Steps 涉及措施
            $.ajax({
                type: "POST",
                async: false,
                url: "../handler/modMoney_Execution.aspx",
                data: {
                    mode: status,
                    PR_ID: $("#hidden_id").val(),
                    PR_Stage: $("#PR_Stage").val(),
                    PR_PlanTitle: $("#PR_PlanTitle").val(),
                    PR_Office: $("#PR_Office").val(),
                    PR_Money: $("#PR_Money").val(),
                    PR_Steps: $("#PR_Steps").val(),
                    PR_CaseName: $("#PR_CaseName").val(),
                    PR_CaseMoney: $("#PR_CaseMoney").val(),
                    PR_SelfMoney: $("#PR_SelfMoney").val(),
                    save_flag: btnStatus
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    var strMsg = "";
                    if (data.indexOf("Error") > -1)
                        alert(data);
                    else {
                        if (data == "reLogin") {
                            alert("請重新登入");
                            window.location = "../WebPage/Login.aspx";
                            return;
                        }
                        if (data == "success") {
                            if (btnStatus == "new") {
                                strMsg = "新增成功";
                            }
                            if (btnStatus == "mod") {
                                strMsg = "修改成功";
                            }
                            alert(strMsg);
                            $("#NewBlock").dialog("close");
                            getData();
                        }
                    }
                }
            });
        }

        //修改 撈該筆資料
        function getDataById(aid) {
            $.ajax({
                type: "POST",
                async: false,
                url: "../handler/getMoney_ExecutionById.aspx",
                data: {
                    aid: aid
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
                        window.location = "Login.aspx";
                        return;
                    }
                    var tabstr = '';
                    if ($(data).find("data_item").length > 0) {
                        $(data).find("data_item").each(function (i) {
                            //PR_CaseName 標案名稱 PR_CaseMoney 標案：發包金額(B) PR_SelfMoney 自辦：金額(C)
                            //PR_Stage 期數 PR_City 執行機關(代碼) PR_CityName 執行機關(名稱)  PR_PlanTitle 計畫名稱 PR_Office 主責局處 PR_Money 金額 PR_Steps 涉及措施
                            $("#PR_Stage").val($(this).children("PR_Stage").text().trim());
                            $("#PR_CityName").empty().append($(this).children("CityName").text().trim());
                            $("#PR_City").val($(this).children("PR_City").text().trim());
                            $("#PR_PlanTitle").val($(this).children("PR_PlanTitle").text().trim());
                            $("#PR_Office").val($(this).children("PR_Office").text().trim());
                            $("#PR_Money").val($(this).children("PR_Money").text().trim());
                            $("#PR_CaseName").val($(this).children("PR_CaseName").text().trim());
                            $("#PR_CaseMoney").val($(this).children("PR_CaseMoney").text().trim());
                            $("#PR_SelfMoney").val($(this).children("PR_SelfMoney").text().trim());
                            $("#PR_Steps").val($(this).children("PR_Steps").text().trim());
                        });
                    }
                }
            });
        }

        function getddl(gno, tagName) {
            $.ajax({
                type: "POST",
                async: false,
                url: "../handler/GetDDL.ashx",
                data: { Group: gno },
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
                        var ddlstr = "";
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
        //新增視窗
        function doShowDialog() {
            var strtitle = "";
            if (btnStatus=="new") {
                strtitle = "新增";
            }
            if (btnStatus=="mod") {
                strtitle = "修改";
            }
            /// dialog setting
            $("#NewBlock").dialog({
                title: strtitle,
                autoOpen: false,
                width: 700,
                height: 650,
                closeOnEscape: true,
                position: { my: "center", at: "center", of: window },
                modal: true,
                resizable: false,
                close: function (event, ui) {
                    $(".dialogInput").val("");
                }
            });
            $("#NewBlock").dialog("open");
        }

        //檢查輸入欄位
        function chkData() {
            //PR_CaseName 標案名稱 PR_CaseMoney 標案：發包金額(B) PR_SelfMoney 自辦：金額(C)
            //PR_Stage 期數 PR_City 執行機關(代碼) PR_CityName 執行機關(名稱)  PR_PlanTitle 計畫名稱 PR_Office 主責局處 PR_Money 金額 PR_Steps 涉及措施
            var status = true;
            var errMsg = "";
            if ($("#PR_Stage").val() == "") {
                status = false;
                errMsg += "請選擇[期數]\n";
            }
            if ($("#PR_Money").val() != "") {
                if (isNaN($("#PR_Money").val())) {
                    status = false;
                    errMsg += "[金額(A)]只能輸入數字\n";
                }
            }
            if ($("#PR_CaseMoney").val() != "") {
                if (isNaN($("#PR_CaseMoney").val())) {
                    status = false;
                    errMsg += "[標案:發包金額(B)]只能輸入數字\n";
                }
            }
            if ($("#PR_SelfMoney").val() != "") {
                if (isNaN($("#PR_SelfMoney").val())) {
                    status = false;
                    errMsg += "[自辦:金額(C)]只能輸入數字\n";
                }
            }
            if (errMsg != "") {
                alert(errMsg);
            }
            return status;
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
            期：<select id="n_Stage" class="inputex" style="margin-right: 10px;">
                <option value="1" selected="selected">第一期</option>
                <option value="2">第二期</option>
                <option value="3">第三期</option>
            </select>
        </div>

        <div id="div_cityName" class="margin10TB font-size3" style="float:left;">
            執行機關：<span id="n_cityName"></span><br />
            經費：<span id="n_money"></span>
        </div>
        <%--<div id="div_money" class="margin10T" style="display: none;">
            經費：<span id="n_money"></span>
        </div>--%>
        <div style="text-align: right; margin-top:25px;">
            <input class="genbtn" type="button" id="addbtn" value="新增經費" onclick="doShowDialog()" />
        </div>
        <div class="stripeMe margin10T font-normal">
            <table id="tablist" width="100%" border="0" cellspacing="0" cellpadding="0">
                <thead>
                    <tr>
                        <th nowrap="nowrap" rowspan="3" style="width: 50px;">編號</th>
                        <th nowrap="nowrap" rowspan="3">主責局處</th>
                        <th nowrap="nowrap" rowspan="3">計畫項目</th>
                        <th nowrap="nowrap" rowspan="3" style="width: 150px;">金額(元)<br />
                            (A)=(B)+(C)</th>
                        <th nowrap="nowrap" colspan="3">處理方式</th>
                        <th nowrap="nowrap" rowspan="3" style="width: 150px;">涉及措施</th>
                        <th nowrap="nowrap" rowspan="3" style="width: 120px;">功能</th>
                    </tr>
                    <tr>
                        <th nowrap="nowrap" colspan="2">標案</th>
                        <th nowrap="nowrap">自辦</th>
                    </tr>
                    <tr>
                        <th nowrap="nowrap">標案名稱</th>
                        <th nowrap="nowrap">發包金額(元)(B)</th>
                        <th nowrap="nowrap">金額(元)(C)</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
            <div id="changpage" class="margin20T textcenter"></div>
        </div>
    </div>

    <div id="NewBlock" style="display: none; text-align: left;">
        <div class="stripeMe margin5T font-normal">
            <table id="addlist" width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td align="right" class="width20">期：
                    </td>
                    <td>
                        <select id="PR_Stage" class="inputex margin50B dialogInput">
                            <option value="">--請選擇--</option>
                            <option value="1">第一期</option>
                            <option value="2">第二期</option>
                            <option value="3">第三期</option>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td align="right">執行機關：</td>
                    <td>
                        <div id="PR_CityName"></div>
                        <input id="PR_City" style="display: none;" />
                    </td>
                </tr>
                <tr>
                    <td align="right">主責局處：</td>
                    <td>
                        <input type="text" class="inputex width100" id="PR_Office" />
                    </td>
                </tr>
                <tr>
                    <td align="right">計畫項目：</td>
                    <td>
                        <select id="PR_PlanTitle" class="inputex margin50B"></select>
                    </td>
                </tr>
                <tr>
                    <td align="right">金額：<br />
                        (A)=(B)+(C)</td>
                    <td>
                        <input type="text" class="inputex width100 num" id="PR_Money" />
                    </td>
                </tr>
                <tr>
                    <td align="right">標案：<br />
                        標案名稱</td>
                    <td>
                        <input type="text" class="inputex width100" id="PR_CaseName" />
                    </td>
                </tr>
                <tr>
                    <td align="right">標案：<br />
                        發包金額(B)</td>
                    <td>
                        <input type="text" class="inputex width100 num" id="PR_CaseMoney" />
                    </td>
                </tr>
                <tr>
                    <td align="right">自辦：<br />
                        金額(C)</td>
                    <td>
                        <input type="text" class="inputex width100 num" id="PR_SelfMoney" />
                    </td>
                </tr>
                <tr>
                    <td align="right">涉及措施：</td>
                    <td>
                        <textarea class="inputex width100" rows="5" id="PR_Steps"></textarea>
                    </td>
                </tr>
                <!--<tr>
                <td align="right">備註：</td>
                <td>
                    <textarea class="inputex width100" rows="5" id="PR_Remark"></textarea>
                </td>
            </tr>-->


            </table>
            <br />
            <div style="text-align: right;">
                <input type="button" class="genbtn" id="subbtn" value="確定" />
                <input type="button" class="genbtn" id="cancelbtn" style="background-color:gray;" value="取消" />
            </div>
        </div>
    </div>
    <input type="hidden" id="hidden_id" />
</asp:Content>
