<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="total-behindEx.aspx.cs" Inherits="WebPage_total_behindEx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script type="text/javascript">
        $(function () {
            loadData("pageload");
            //期數下拉選單change
            $("#ddlStage").change(function () {
                loadData();
            });
            //匯出按鈕
            $("#btn_export").click(function () {
                window.location = "../handler/ExportTotalBehind.aspx?s=" + encodeURIComponent($("#ddlStage").val()) + "";
            });
        });

        function loadData(status) {
            $.ajax({
                type: "POST",
                async: true, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/accReport.ashx",
                data: {
                    func: "load_reportTotalBehindForEx",
                    strStage: $("#ddlStage").val()
                },
                beforeSend: function () {
                   $("#tbody_list").empty().append('<tr><td colspan="62">資料統計中，請稍後...</td></tr>');
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
                            //************************表頭************************//
                            if (status == "pageload") {
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
                                $("#thead_list").empty().append(strHtml);
                            }
                            //************************END************************//

                            //************************縣市資料************************//
                            var tabstr = "";
                            var objCity = "";
                            if ($(data).find("City").length > 0) {
                                $(data).find("City").each(function (i) {
                                    tabstr += '<tr>';
                                    tabstr += '<td>' + $(this).attr("city_Item_cn") + '</td>';
                                    tabstr += '<td>' + $(this).attr("city_Year") + '年第' + $(this).attr("city_Season") + '季</td>';
                                    objCity = $(this);
                                    $(data).find("dataHead").each(function (i) {
                                        var tmpChildData = "";
                                        var tmpType = $(this).children("C_Item").text().trim();
                                        objCity.children().each(function () {
                                            if (tmpType == $(this).children("RM_CPType").text().trim()) {
                                                tmpChildData += '<td>' + $(this).children("RM_Planning").text().trim() + '</td>';
                                                //01 & 03  申請、完成 抓 sum3、sum4  其他的申請、完成 抓 sum1、sum2
                                                if (tmpType == "01" || tmpType == "03") {
                                                    tmpChildData += '<td>' + $(this).children("sum3").text().trim() + '</td>';
                                                    tmpChildData += '<td>' + $(this).children("sum4").text().trim() + '</td>';
                                                }
                                                else {
                                                    tmpChildData += '<td>' + $(this).children("sum1").text().trim() + '</td>';
                                                    tmpChildData += '<td>' + $(this).children("sum2").text().trim() + '</td>';
                                                }
                                            }
                                        });
                                        tabstr += (tmpChildData == "") ? '<td>0</td><td>0</td><td>0</td>' : tmpChildData;
                                    });
                                    tabstr += '</tr>';
                                });
                            }
                            else {
                                tabstr += "<tr><td colspan='62'>查詢無資料</td></tr>";
                            }
                            //************************END************************//
                        }
                        $("#tbody_list").empty().append(tabstr);
                    }
                }//success end
            });//ajax end
        }

    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div style="padding-left:10px; padding-right:10px;">
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
        <div class="stripecomplex margin5T font-normal" style="overflow-x:scroll;">
            <table id="airlist" border="0" cellspacing="0" cellpadding="0" width="100%">
                <thead id="thead_list"></thead>
                <tbody id="tbody_list"></tbody>
            </table>
        </div>
    </div><br />
</asp:Content>


