<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="total-lag.aspx.cs" Inherits="WebPage_total_lag" %>

<asp:Content ID="Content3" ContentPlaceHolderID="head" Runat="Server">
    <script type="text/javascript">
        $(function () {
            loadData();
            //期數下拉選單change
            $("#ddlStage").change(function () {
                loadData();
            });
            $("#btn_export").click(function () {
                window.location = "../handler/ExportTotalLag.aspx?s=" + encodeURIComponent($("#ddlStage").val()) + "";
            });
        });
        function loadData(){
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/accReport.ashx",
                data: {
                    func: "load_reportTotalLog",
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
                        if (data != null && data != "") {
                            data = $.parseXML(data);
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function () {
                                    //if ($(this).children("RS_year").text().trim() == "" && $(this).children("RS_Season").text().trim() == "") {
                                    //    strHtml += "<tr>";
                                    //    strHtml += "<th style='text-align:center;'>" + $(this).children("C_Item_cn").text().trim() + "</th>";
                                    //    strHtml += "<td></td><td></td><td></td><td></td>";
                                    //    strHtml += "</tr>";
                                    //} else {
                                        strHtml += "<tr>";
                                        strHtml += "<th>" + $(this).children("C_Item_cn").text().trim() + "</th>";
                                        strHtml += "<td style='text-align:center;'>" + $(this).children("RS_year").text().trim() + "年<br />第" + $(this).children("RS_Season").text().trim() + "季</td>";
                                        strHtml += "<td>" + splitStr($(this).children("RS_01Why").text().trim()) + "</td>";
                                        strHtml += "<td>" + splitStr($(this).children("RS_02Why").text().trim()) + "</td>";
                                        strHtml += "<td>" + splitStr($(this).children("RS_03Why").text().trim()) + "</td>";
                                        strHtml += "<td>" + splitStr($(this).children("RSExWhy").text().trim()) + "</td>";
                                        strHtml += "</tr>";
                                    //}
                                });
                            }
                        } else {
                            strHtml += "<tr><td colspan='6'>尚無資料</td></tr>";
                        }
                        $("#tbody_list").empty().append(strHtml);
                    }
                }//success end
            });//ajax end
        }

        //將XML回傳回來的遭遇困難字串拆開 然後塞換行符號
        function splitStr(str) {
            var strVal = "";//
            var strValAll = "";//回傳回去的字串
            var splitVal = "";//split的字串 &#x0D;
            var splitVal2 = "";//split的字串 \n
            if (str != null && str.trim() != "") {
                splitVal = str.trim().split("&#x0D;");
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
            if (strVal != null && strVal.trim() != "") {
                splitVal2 = strVal.split("\\n");
                for (var i = 0; i < splitVal2.length; i++) {
                    if (splitVal2[i].toString().trim()!="") {
                        if (strValAll == "") {
                            strValAll += splitVal2[i].toString().trim();
                        } else {
                            strValAll += "<br />" + splitVal2[i].toString().trim();
                        }
                    }
                }
            }

            return strValAll;
        }
    </script>
</asp:Content>
<asp:Content ID="Content4" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">計畫執行進度遭遇困難</span></div>
            <div class="right">附加圖表 / 管理員總覽表 / 計畫執行進度遭遇困難</div>
        </div>

        期別：
        <select id="ddlStage" class="inputex" style="margin-top:10px;">
            <option value="1"  selected="selected">第一期</option>
            <option value="2">第二期</option>
            <option value="3">第三期</option>
        </select>
        &nbsp;&nbsp;<a href="javascript:void(0);" class="genbtn" id="btn_export">匯出Excel</a><br /><br />
        <div class="stripeMe">
            <table id="airlist" border="0" cellspacing="0" cellpadding="0" width="100%" id="tbody_list">
                <thead>
                    <tr>
                        <th style="width:8%;">縣市</th>
                        <th style="width:8%;">季</th>
                        <th style="width:16%;">節電基礎工作</th>
                        <th style="width:16%;">因地制宜</th>
                        <th style="width:16%;">設備汰換與智慧用電</th>
                        <th style="width:16%;">擴大補助</th>
                    </tr>
                </thead>
                <tbody id="tbody_list"></tbody>
            </table>
        </div>
    </div><br />
</asp:Content>


