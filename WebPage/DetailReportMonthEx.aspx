<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="DetailReportMonthEx.aspx.cs" Inherits="WebPage_DetailReportMonthEx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script type="text/javascript">
        $(function () {
            $(".pbtn").hide();
            $("#hidden_val").val($.getParamValue('v'));
            load_data();
            //通過 不通過按鈕
            $("#btn_ok,#btn_notok").click(function () {
                var str = "";
                var this_id = $(this).attr("id");
                if (this_id == "btn_ok") {
                    str = "通過";
                }
                if (this_id == "btn_notok") {
                    str = "不通過";
                }
                if (confirm("該月報是否[" + str + "]?")) {
                    goCheck(this_id);
                }

            });

            //匯出 button
            $(document).on("click", "#exbtn", function () {
                openExport();
            });

            //確認匯出
            $(document).on("click", "#expbtn", function () {
                var type = $("input[name='extype']:checked").val();
                window.open("../handler/ExportReportMonthEx.aspx?v=" + encodeURIComponent($.getParamValue('v')) + "&tp=" + type);
                $.fancybox.close();
            });

            //審核不通過確定按鈕(發送MAIL)
            $("#submitbtn").click(function () {
                var type = $("input[name='extype']:checked").val();
                $.ajax({
                    type: "POST",
                    async: false, //在沒有返回值之前,不會執行下一步動作
                    url: "../handler/SendMailCheck.aspx",
                    data: {
                        func: "load_projectbyperson",
                        reportGuid: $("#hidden_val").val(),
                        mailtype: "MnotOK",
                        mailbody: $("#textOp").val(),
                        yyyy: $("#hidden_year").val(),
                        mm: $("#hidden_month").val()
                    },
                    error: function (xhr) {
                        alert("Error " + xhr.status);
                        console.log(xhr.responseText);
                    },
                    success: function (data) {
                        $.fancybox.close();
                        if (data == "timeout") {
                            alert("請重新登入");
                            location.href = "Login.aspx";
                        } else if (data.indexOf("Error") > -1) {
                            alert(data);
                        } else {
                            checkUpdate($("#hidden_typeval").val());
                        }
                    }
                });//ajax end

                //window.open("../Handler/SendMailCheck.aspx?v=" + $("#hidden_reportGuid").val() + "&tp=MnotOK");
                $.fancybox.close();
            });
        });

        //撈帶過來的v<mid>資料 填表人 電話 主管 機關 局處
        function load_citydatabymid() {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/ReportMonth.ashx",
                data: {
                    func: "load_projectbyperson",
                    str_mid: decodeURIComponent($.getParamValue('v'))
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    var str_html = "";
                    if (data == "timeout") {
                        alert("請重新登入");
                        location.href = "Login.aspx";
                    } else if (data.indexOf("Error") > -1) {
                        alert(data);
                    } else {
                        data = $.parseXML(data);
                        if ($(data).find("data_item").length > 0) {
                            $(data).find("data_item").each(function (i) {
                                $("#div_cityname").empty().append($(this).children("C_Item_cn").text().trim());
                                $("#div_office").empty().append($(this).children("M_Office").text().trim());
                            });
                        }
                    }
                }
            });//ajax end
        }
        //撈月報資料
        function load_data() {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/DetailReport.ashx",
                data: {
                    func: "load_MonthReportDataEx",//txt_stage  
                    str_mid: decodeURIComponent($.getParamValue('v'))
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    var str_html = "";
                    if (data == "timeout") {
                        alert("請重新登入");
                        location.href = "Login.aspx";
                    } else if (data.indexOf("Error") > -1) {
                        alert(data);
                    } else {
                        var strHtml = "";
                        var chk_status, chk_type, chk_date, strpeople, strphone, strbossname, strrange, strReportGuid;
                        data = $.parseXML(data);
                        if ($(data).find("data_item").length > 0) {
                            var strCreateDate, strChkDate;
                            var strday = new Date();
                            var str = "", sdate = "", edate = "";
                            var splits, splite;
                            $(data).find("data_item").each(function (i) {
                                if (i == 0) {
                                    chk_status = $(this).children("RC_Status").text().trim();
                                    chk_type = $(this).children("RC_CheckType").text().trim();
                                    strReportGuid = $(this).children("RM_ReportGuid").text().trim();
                                    if ($(this).children("RC_CheckDate").text().trim() != "" && $(this).children("RC_CheckDate").text().trim() != "1900-01-01T00:00:00+08:00") {
                                        chk_date = formatDate($(this).children("RC_CheckDate").text().trim());
                                        strCreateDate = formatDate($(this).children("RM_ModDate").text().trim());
                                    } else {
                                        chk_date = formatDate("");
                                        strCreateDate = strday.getFullYear() + "/" + (strday.getMonth() + 1) + "/" + strday.getDate();
                                    }
                                    strpeople = $(this).children("M_Name").text().trim();
                                    strphone = $(this).children("M_Tel").text().trim();
                                    strbossname = $(this).children("bossname").text().trim();
                                    $("#div_cityname").empty().append($(this).children("cityname").text().trim());
                                    $("#div_office").empty().append($(this).children("M_Office").text().trim());
                                    $("#div_stage").empty().append("第" + $(this).children("RM_Stage").text().trim() + "期");

                                    if ($(this).children("RM_Stage").text().trim() == "1") {
                                        sdate = $(this).children("I_1_Sdate").text().trim();
                                        edate = $(this).children("I_1_Edate").text().trim();
                                        strrange = monthDiff(new Date($(this).children("I_1_Sdate").text().trim()), new Date($(this).children("I_1_Edate").text().trim()));
                                    }
                                    if ($(this).children("RM_Stage").text().trim() == "2") {
                                        sdate = $(this).children("I_2_Sdate").text().trim();
                                        edate = $(this).children("I_2_Edate").text().trim();
                                        strrange = monthDiff(new Date($(this).children("I_2_Sdate").text().trim()), new Date($(this).children("I_2_Edate").text().trim()));
                                    }
                                    if ($(this).children("RM_Stage").text().trim() == "3") {
                                        sdate = $(this).children("I_3_Sdate").text().trim();
                                        edate = $(this).children("I_3_Edate").text().trim();
                                        strrange = monthDiff(new Date($(this).children("I_3_Sdate").text().trim()), new Date($(this).children("I_3_Edate").text().trim()));
                                    }
                                    splits = sdate.split("/");
                                    splite = edate.split("/");
                                    //自107年1月1日起至 年 月 日止，計    個月
                                    str += "自" + (parseInt(splits[0]) - 1911) + "年" + splits[1] + "月" + splits[2] + "日起至";
                                    str += "自" + (parseInt(splite[0]) - 1911) + "年" + splite[1] + "月" + splite[2] + "日止，共計";
                                    str += monthDiff(new Date(sdate), new Date(edate)) + "個月";
                                    $("#div_stage").empty().append(str);
                                    $("#div_yyyymm").empty().append(($(this).children("RM_Year").text().trim() - 1911) + "年" + $(this).children("RM_Month").text().trim() + "月");
                                    $("#hidden_year").val(($(this).children("RM_Year").text().trim() - 1911));
                                    $("#hidden_month").val($(this).children("RM_Month").text().trim());
                                }

                                strHtml += getHtml($(this));

                            });
                            strHtml += '<div class="OchiRow">';
                            strHtml += '<div class="OchiThird">';
                            strHtml += '<div class="OchiCell OchiTitle TitleSetWidth">填表人</div>';
                            strHtml += '<div class="OchiCell width100" id="div_person">' + strpeople + '</div>';
                            strHtml += '</div>';
                            strHtml += '<div class="OchiThird">';
                            strHtml += '<div class="OchiCell OchiTitle TitleSetWidth">電話</div>';
                            strHtml += '<div class="OchiCell width100" id="div_phone">' + strphone + '</div>';
                            strHtml += '</div>';
                            strHtml += '<div class="OchiThird">';
                            strHtml += '<div class="OchiCell OchiTitle TitleSetWidth">填表日期</div>';
                            strHtml += '<div class="OchiCell width100" id="div_createdate">' + strCreateDate + '</div>';
                            strHtml += '</div>';
                            strHtml += '</div>';
                            strHtml += '<div class="OchiRow">';
                            strHtml += '<div class="OchiHalf">';
                            strHtml += '<div class="OchiCell OchiTitle TitleSetWidth">主管</div>';
                            strHtml += '<div class="OchiCell width100">';
                            strHtml += '<div class="OchiTableInner width100">';
                            strHtml += '<div class="OchiCellInner width100" id="div_dossname">' + strbossname + '</div>';
                            strHtml += '</div>';
                            strHtml += '</div>';
                            strHtml += '</div>';
                            strHtml += '<div class="OchiHalf">';
                            strHtml += '<div class="OchiCell OchiTitle TitleSetWidth">簽核日期</div>';
                            strHtml += '<div class="OchiCell width100">';
                            strHtml += '<div class="OchiTableInner width100">';
                            strHtml += '<div class="OchiCellInner width100" id="div_chkdate">' + chk_date + '</div>';
                            strHtml += '</div>';
                            strHtml += '</div>';
                            strHtml += '</div>';
                            strHtml += '</div>';
                            //strHtml += '<div class="twocol margin15T margin5B">';
                            //strHtml += '<div style="text-align:center;">';
                            //strHtml += '<a href="javascript:void(0);" class="genbtn" id="btn_ok">通過</a>&nbsp;&nbsp;&nbsp;';
                            //strHtml += '<a href="javascript:void(0);" class="genbtn" id="btn_notok">不通過</a>';
                            //strHtml += '</div>';
                            //strHtml += '</div>';
                            $("#div_data").empty().append(strHtml);
                            $("#hidden_reportGuid").val(strReportGuid);
                            //if (chk_type == "Y") {
                            //    $(".genbtn").hide();
                            //}
                        }
                        if ($(data).find("data_acc").length > 0) {
                            var strCreateDate, strChkDate;
                            $(data).find("data_acc").each(function (i) {
                                if ($(this).children("competence").text().trim() != "02") {
                                    $("#btn_ok").hide();
                                    $("#btn_notok").hide();
                                    //if (chk_type == "Y") {
                                        $("#exbtn").show();
                                        $("#exbtnPDF").show();
                                    //}
                                } else {
                                    if (chk_type != "Y") {
                                        $("#btn_ok").show();
                                        $("#btn_notok").show();
                                        $("#exbtn").show();
                                        $("#exbtnPDF").show();
                                    } else {
                                        $("#btn_ok").hide();
                                        $("#btn_notok").hide();
                                        $("#exbtn").show();
                                        $("#exbtnPDF").show();
                                    }
                                }

                            });

                        }
                    }
                }
            });//ajax end
        }


        //湊出各類別表格
        function getHtml(xmlData) {
            var strItem = xmlData.children("P_ItemName").text().trim();
            var type1 = xmlData.children("P_ExType").text().trim();
            var type2 = xmlData.children("P_ExDeviceType").text().trim();
            var ItemCname = xmlData.children("P_ItemName_c").text().trim();
            var str = "",strUnit="";
            var splitRM_Finish, splitRM_Finish01, P_ExFinish;
            var ptno = xmlData.children("P_ItemName").text().trim();

            P_ExFinish = xmlData.children("P_ExFinish").text().trim().split('.');
            splitRM_Finish = xmlData.children("countFinishKW").text().trim().split('.');
            splitRM_Finish01 = xmlData.children("countFinish02").text().trim().split('.');
            strP_ExFinish = P_ExFinish[0];
            strRM_Finish = splitRM_Finish[0];
            strRM_Finish01 = splitRM_Finish01[0];
            


            //if (strItem == "99") {
            //    ptno = ptno + xmlData.children("P_ExType").text().trim() + xmlData.children("P_ExDeviceType").text().trim();
            //    ItemCname += "<br />" + xmlData.children("P_ExType_c").text().trim() + "-" + xmlData.children("P_ExDeviceType_c").text().trim();
            //}

            if (strItem == "01" || strItem == "23" || strItem == "33") {
                strUnit = "KW";
            }
            if (strItem == "02" || strItem == "21") {
                strUnit = "具";
            }
            if (strItem == "03" || strItem == "22" || strItem == "29") {
                strUnit = "盞";
            }
            if (strItem == "04" || strItem == "26" || strItem == "30" || strItem == "31") {
                strUnit = "套";
            }
            if (strItem == "05" || strItem == "06" || strItem == "07" || strItem == "08" || strItem == "09" || strItem == "10" || strItem == "11" || strItem == "12" || strItem == "13" || strItem == "14" || strItem == "15" || strItem == "19" || strItem == "20") {
                strUnit = "台";
            }
            if (strItem == "17" || strItem == "18" || strItem == "24" || strItem == "28" || strItem == "32") {
                strUnit = "顆";
            }
            if (strItem == "27") {
                strUnit = "噸";
            }
            if (strItem == "25") {
                strUnit = "個";
            }
            if (strItem=="16") {
                strUnit = "10顆一單位";
            }

            if (strItem == "01" || strItem == "33") {//type1 == "02" && type2 == "01" =>空調

                str += '<div class="OchiRow">';
                str += '<div class="OchiCell OchiTitle TitleSetWidth">' + ItemCname + '</div>';
                str += '<div class="OchiCell width100">';
                str += '<div class="OchiTableInner width100 margin5T">';
                //str += '<div class="OchiCellInner nowrap textcenter">本期累計申請數：</div>';
                //str += '<div class="OchiCellInner width50">';
                //str += xmlData.children("RM_Planning").text().trim() + '&nbsp;(kW)';
                //str += '</div>';
                str += '<div class="OchiCellInner nowrap textcenter" style="text-align:left;">本期累計完成數：' + xmlData.children("RM_Finish").text().trim()  + '&nbsp;(kW)</div>';
                //str += '<div class="OchiCellInner width50"></div>';

                str += '</div>';
                str += '<div class="OchiTableInner width100 margin5T">';
                str += '<div class="OchiTableInner width100">';
                str += '<div class="OchiCellInner nowrap textcenter">本期規劃數：</div>';
                str += '<div class="OchiCellInner width50">';
                str += '<span>' + xmlData.children("RM_Planning").text().trim() + '&nbsp;(kW)</span>';
                str += '</div>';
                str += '<div class="OchiCellInner nowrap textcenter">本期累計核定數：</div>';
                str += '<div class="OchiCellInner width50">';
                str += '<span>' + xmlData.children("RM_Finish01").text().trim() + '&nbsp;(台)</span>';
                str += '</div>';
                str += '</div>';
                str += '</div>';
                str += '<div class="OchiTableInner width100">';
                str += '<div class="stripeMepure font-normal margin5T">';
                str += '<table border="0" cellspacing="0" cellpadding="0" width="100%">';
                str += '<tr>';
                str += '<th colspan="3">本月申請數量(台)</th>';
                str += '<th colspan="3">本月核定數量(台)</th>';
                str += '<th colspan="3">本月申請總冷氣能力(kW) </th>';
                str += '<th colspan="3">本月完成總冷氣能力(kW)</th>';
                str += '</tr>';
                str += '<tr>';
                str += '<td align="center">服務業</td>';
                str += '<td align="center">機關學校</td>';
                str += '<td align="center">住宅</td>';
                str += '<td align="center">服務業</td>';
                str += '<td align="center">機關學校</td>';
                str += '<td align="center">住宅</td>';
                str += '<td align="center">服務業</td>';
                str += '<td align="center">機關學校</td>';
                str += '<td align="center">住宅</td>';
                str += '<td align="center">服務業</td>';
                str += '<td align="center">機關學校</td>';
                str += '<td align="center">住宅</td>';
                str += '</tr>';
                str += '<tr>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type1Value1").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type1Value2").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type1Value3").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type2Value1").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type2Value2").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type2Value3").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type3Value1").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type3Value2").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type3Value3").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type4Value1").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type4Value2").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type4Value3").text().trim() + '</td>';
                str += '</tr>';
                str += '<tr>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">' + (xmlData.children("RM_Type1ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type1ValueSum").text().trim()) + '</td>';
                str += '<td align="right">合計</td>';                                                                                                                                                                                                                                                                                                                                                                  
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">' + (xmlData.children("RM_Type2ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type2ValueSum").text().trim()) + '</td>';
                str += '<td align="right">合計</td>';                                                                                                                                                                                                                                                                                                                                                                  
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">' + (xmlData.children("RM_Type3ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type3ValueSum").text().trim()) + '</td>';
                str += '<td align="right">合計</td>';                                                                                                                                                                                                                                                                                                                                                                  
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">' + (xmlData.children("RM_Type4ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type4ValueSum").text().trim()) + '</td>';
                str += '</td>';
                str += '</tr>';
                str += '<tr>';
                str += '<th colspan="4">本月申請數預期年節電量(度)</th>';
                str += '<th colspan="4">本月核定數預期年節電量(度)</th>';
                str += '<th colspan="4">本月未核定數之年節電量(度) </th>';
                str += '</tr>';
                str += '<tr>';
                str += '<td colspan="4" style="text-align:right;">' + (xmlData.children("RM_PreVal").text().trim() == "" ? "0" : xmlData.children("RM_PreVal").text().trim()) + '</td>';
                str += '<td colspan="4" style="text-align:right;">' + (xmlData.children("RM_ChkVal").text().trim() == "" ? "0" : xmlData.children("RM_ChkVal").text().trim()) + '</td>';
                str += '<td colspan="4" style="text-align:right;">' + (xmlData.children("RM_NotChkVal").text().trim() == "" ? "0" : xmlData.children("RM_NotChkVal").text().trim()) + '</td>';
                str += '</tr>';
                str += '</table>';
                str += '</div>';
                str += '</div>';
                str += '<div class="OchiTableInner width100 margin5TB">';
                str += '補充說明：<br />' + xmlData.children("RM_Remark").text().trim();
                str += '</div>';
                str += '</div>';
                str += '</div>';
            }
            else if (strItem == "02" || strItem == "03") {//type1 == "02" && type2 == "01" =>照明
                str += '<div class="OchiRow">';
                str += '<div class="OchiCell OchiTitle TitleSetWidth">' + ItemCname + '</div>';
                str += '<div class="OchiCell width100">';

                str += '<div class="OchiTableInner width100">';
                str += '<div class="OchiCellInner nowrap textcenter">本期累計核定數：</div>';
                str += '<div class="OchiCellInner width33">';
                str += '<span>' + xmlData.children("countFinish02").text().trim() + '&nbsp;(' + strUnit + ')</span>';
                str += '</div>';
                str += '<div class="OchiCellInner nowrap textcenter">本期規劃數：</div>';
                str += '<div class="OchiCellInner width50">';
                str += '<span>' + xmlData.children("RM_Planning").text().trim() + '&nbsp;(' + strUnit + ')</span>';
                str += '</div>';
                str += '</div>';


                str += '<div class="OchiTableInner width100">';
                str += '<div class="stripeMepure font-normal margin5T">';
                str += '<table border="0" cellspacing="0" cellpadding="0" width="100%">';
                str += '<tr>';
                str += '<th colspan="3">本月申請數量(' + strUnit + ')</th>';
                str += '<th colspan="3">本月核定數量(' + strUnit + ')</th>';
                str += '<th colspan="3">本月申請更換照明瓦數(W) </th>';
                str += '<th colspan="3">本月完成更換照明瓦數(W)</th>';
                str += '</tr>';
                str += '<tr>';
                str += '<td align="center">服務業</td>';
                str += '<td align="center">機關學校</td>';
                str += '<td align="center">住宅</td>';
                str += '<td align="center">服務業</td>';
                str += '<td align="center">機關學校</td>';
                str += '<td align="center">住宅</td>';
                str += '<td align="center">服務業</td>';
                str += '<td align="center">機關學校</td>';
                str += '<td align="center">住宅</td>';
                str += '<td align="center">服務業</td>';
                str += '<td align="center">機關學校</td>';
                str += '<td align="center">住宅</td>';
                str += '</tr>';
                str += '<tr>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type1Value1").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type1Value2").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type1Value3").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type2Value1").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type2Value2").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type2Value3").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type3Value1").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type3Value2").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type3Value3").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type4Value1").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type4Value2").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type4Value3").text().trim() + '</td>';
                str += '</tr>';
                str += '<tr>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">' + (xmlData.children("RM_Type1ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type1ValueSum").text().trim()) + '</td>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">' + (xmlData.children("RM_Type1ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type2ValueSum").text().trim()) + '</td>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">' + (xmlData.children("RM_Type1ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type3ValueSum").text().trim()) + '</td>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">' + (xmlData.children("RM_Type1ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type4ValueSum").text().trim()) + '</td>';
                str += '</tr>';
                str += '<tr>';
                str += '<th colspan="4">本月申請數預期年節電量(度)</th>';
                str += '<th colspan="4">本月核定數預期年節電量(度)</th>';
                str += '<th colspan="4">本月未核定數之年節電量(度) </th>';
                str += '</tr>';
                str += '<tr>';
                str += '<td colspan="4" style="text-align:right;">' + (xmlData.children("RM_PreVal").text().trim() == "" ? "0" : xmlData.children("RM_PreVal").text().trim()) + '</td>';
                str += '<td colspan="4" style="text-align:right;">' + (xmlData.children("RM_ChkVal").text().trim() == "" ? "0" : xmlData.children("RM_ChkVal").text().trim()) + '</td>';
                str += '<td colspan="4" style="text-align:right;">' + (xmlData.children("RM_NotChkVal").text().trim() == "" ? "0" : xmlData.children("RM_NotChkVal").text().trim()) + '</td>';
                str += '</tr>';
                str += '</table>';
                str += '</div>';
                str += '</div>';
                str += '<div class="OchiTableInner width100 margin5TB">';
                str += '補充說明：<br />' + xmlData.children("RM_Remark").text().trim();
                str += '</div>';
                str += '</div>';
                str += '</div>';
            }
            else {
                //if (strItem == "05") {
                //    strUnit = "KW";
                //}
                //else if (strItem == "14") {
                //    strUnit = "組";
                //} else {
                //    strUnit = "台";
                //}
                
                var splitcountFinish03 = "", countFinish03 = "";
                splitcountFinish03 = xmlData.children("countFinish03").text().trim().split('.');
                countFinish03 = splitcountFinish03[0];

                if (strItem == "27") {
                    countFinish03 = xmlData.children("countFinish03").text().trim();
                }

                str += '<div class="OchiRow">';
                str += '<div class="OchiCell OchiTitle TitleSetWidth">' + ItemCname + '</div>';
                str += '<div class="OchiCell width100">';
                str += '<div class="OchiTableInner width100">';
                str += '<div class="OchiCellInner nowrap textcenter">本期累計完成數:</div>';
                str += '<div class="OchiCellInner width33">' + countFinish03 + '&nbsp;(' + strUnit + ')</div>';
                str += '<div class="OchiCellInner nowrap textcenter">本期累計申請數:</div>';
                str += '<div class="OchiCellInner width33"><span>' + xmlData.children("countApply01").text().trim() + '&nbsp;(' + strUnit + ')</span></div>';
                str += '<div class="OchiCellInner nowrap textcenter">本期規劃數:</div>';
                str += '<div class="OchiCellInner width33">';
                str += '<span>' + (strP_ExFinish == "" ? "0" : strP_ExFinish) + '&nbsp;(' + strUnit + ')</span>';
                str += '</div>';
                str += '</div>';
                str += '<div class="OchiTableInner width100">';
                str += '<div class="stripeMepure font-normal margin5T">';
                str += '<table border="0" cellspacing="0" cellpadding="0" width="100%">';
                str += '<tr>';
                str += '<th colspan="3" valign="top">本月申請數量(' + strUnit + ')</th>';
                str += '<th colspan="3" valign="top">本月核定數量(' + strUnit + ')</th>';
                str += '<th colspan="3" valign="top">本月完成數量(' + strUnit + ')</th>';
                str += '</tr>';
                str += '<tr>';
                str += '<td align="center">服務業</td>';
                str += '<td align="center">機關學校</td>';
                str += '<td align="center">住宅</td>';
                str += '<td align="center">服務業</td>';
                str += '<td align="center">機關學校</td>';
                str += '<td align="center">住宅</td>';
                str += '<td align="center">服務業</td>';
                str += '<td align="center">機關學校</td>';
                str += '<td align="center">住宅</td>';
                str += '</tr>';
                str += '<tr>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type1Value1").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type1Value2").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type1Value3").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type2Value1").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type2Value2").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type2Value3").text().trim() + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type3Value1").text().trim().replace(".0","") + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type3Value2").text().trim().replace(".0","") + '</td>';
                str += '<td style="text-align:right;">' + xmlData.children("RM_Type3Value3").text().trim().replace(".0","") + '</td>';
                str += '</tr>';
                str += '<tr>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">' + (xmlData.children("RM_Type1ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type1ValueSum").text().trim()) + '</td>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">' + (xmlData.children("RM_Type2ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type2ValueSum").text().trim()) + '</td>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">' + (xmlData.children("RM_Type3ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type3ValueSum").text().trim()) + '</td>';
                str += '</tr>';
                str += '<tr>';
                str += '<th colspan="3">本月申請數預期年節電量(度)</th>';
                str += '<th colspan="3">本月核定數預期年節電量(度)</th>';
                str += '<th colspan="3">本月未核定數之年節電量(度)</th>';
                str += '</tr>';
                str += '<tr>';
                str += '<td colspan="3" style="text-align:right;">' + (xmlData.children("RM_PreVal").text().trim() == "" ? "0" : xmlData.children("RM_PreVal").text().trim()) + '</td>';
                str += '<td colspan="3" style="text-align:right;">' + (xmlData.children("RM_ChkVal").text().trim() == "" ? "0" : xmlData.children("RM_ChkVal").text().trim()) + '</td>';
                str += '<td colspan="3" style="text-align:right;">' + (xmlData.children("RM_NotChkVal").text().trim() == "" ? "0" : xmlData.children("RM_NotChkVal").text().trim()) + '</td>';
                str += '</tr>';
                str += '</table>';
                str += '</div>';
                str += '</div>';
                str += '<div class="OchiTableInner width100 margin5T">補充說明：<br />' + xmlData.children("RM_Remark").text().trim() + '</div>';
                str += '</div>';
                str += '</div>';
            }
            return str;
        }


        //計算幾個月
        function monthDiff(d1, d2) {
            var months;
            months = (d2.getFullYear() - d1.getFullYear()) * 12;
            //months -= d1.getMonth() + 1;
            months -= d1.getMonth() - 1;
            months += d2.getMonth();
            //2017/01/01~2017/01/23 這樣要算1個月
            //2017/01/01~2017/02/12 這樣要算2個月
            return months <= 0 ? 1 : months;
        }
        //格式化日期
        function formatDate(date) {
            if (date == "") {
                var d = new Date();
            } else {
                var d = new Date(date);
            }
            var month = '' + (d.getMonth() + 1);
            var day = '' + d.getDate();
            var year = d.getFullYear();

            if (month.length < 2) month = '0' + month;
            if (day.length < 2) day = '0' + day;

            return [year, month, day].join('/');
        }
        //送審
        function goCheck(id) {
            var chkType = "";
            if (id == "btn_ok") {
                chkType = "Y";
                checkUpdate(chkType);
            }
            if (id == "btn_notok") {
                chkType = "N";
                $("#submitbtn").show();
                $("#hidden_typeval").val(chkType);
                openMailOP();

            }

        }

        //更新月報審核通過/不通過
        function checkUpdate(strtype) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/ReportCheck.ashx",
                data: {
                    func: "reportCheck",
                    str_reportGuid: $("#hidden_reportGuid").val(),
                    str_checkType: strtype,
                    str_ReportType: "01",
                    yyyy: $("#hidden_year").val(),
                    mm: $("#hidden_month").val()
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    var str_html = "";
                    if (data == "timeout") {
                        alert("請重新登入");
                        location.href = "Login.aspx";
                    } else if (data.indexOf("Error") > -1) {
                        alert(data);
                    } else {
                        if (data == "success") {
                            alert("審核操作成功");
                            location.href = "ReviewMonth.aspx";
                        }
                    }
                }
            });//ajax end
        }

        //打開輸入"審核不通過"dialog
        function openMailOP() {
            $.fancybox({
                href: "#opblock",
                title: "",
                //closeBtn: false,
                minWidth: "400",
                minHeight: "200",
                wrapCSS: 'fancybox-custom',
                openEffect: 'elastic',
                closeEffect: 'elastic',
                helpers: {
                    title: {
                        type: 'inside'
                    },
                    overlay: {
                        css: {
                            'background': 'gary'
                        },
                        locked: false   //開始fancybox時,背景是否回top
                        //closeClick: false //點背景關閉 fancybox
                    }
                }
            });
        }

        function openExport() {
            $.fancybox({
                href: "#exblock",
                title: "",
                //closeBtn: false,
                minWidth: "200",
                minHeight: "100",
                wrapCSS: 'fancybox-custom',
                openEffect: 'elastic',
                closeEffect: 'elastic',
                helpers: {
                    title: {
                        type: 'inside'
                    },
                    overlay: {
                        css: {
                            'background': 'gary'
                        },
                        locked: false   //開始fancybox時,背景是否回top
                        //closeClick: false //點背景關閉 fancybox
                    }
                }
            });
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">月報</span></div>
            <!-- left -->
            <div class="right">進度報表 / 月報</div>
            <!-- right -->
        </div>
        <div style="text-align: right; margin-top: 10px;">
            <a href="javascript:void(0);" class="genbtn" id="exbtn" style="display: none;">匯出</a>
        </div>
        <div class="OchiTrasTable width100 TitleLength04 margin20T">
            <!-- 雙欄 執行機關 & 承辦局處-->
            <div class="OchiRow">
                <div class="OchiHalf">
                    <div class="OchiCell OchiTitle TitleSetWidth">執行機關</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <div class="OchiCellInner width100" id="div_cityname"></div>
                        </div>
                        <!-- OchiTableInner -->
                        <!-- cell內容end -->
                    </div>
                </div>
                <!-- OchiHalf -->
                <div class="OchiHalf">
                    <div class="OchiCell OchiTitle TitleSetWidth">承辦局處</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <div class="OchiCellInner width100" id="div_office"></div>
                        </div>
                        <!-- OchiTableInner -->
                        <!-- cell內容end -->
                    </div>
                </div>
                <!-- OchiHalf -->
            </div>
            <!-- OchiRow -->


            <!-- 單欄 第幾期 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">期數</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100" id="div_stage"></div>
                    <!-- OchiTableInner -->
                    <!-- cell內容end -->
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 單欄 填報月份 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">提報月份</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100">
                        <div class="OchiCellInner nowrap textcenter" style="text-align: left;" id="div_yyyymm"></div>
                    </div>
                    <!-- OchiTableInner -->
                    <!-- cell內容end -->
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->


            <div class="font-bold font-title margin10T">本月執行進度</div>
            <div id="div_data">
            </div>
            <div class="margin15T margin5B">
                <div style="text-align: center;">
                    <a href="javascript:void(0);" class="genbtn pbtn" id="btn_ok">通過</a>&nbsp;&nbsp;&nbsp;
                    <a href="javascript:void(0);" class="genbtn pbtn" id="btn_notok">不通過</a>
                </div>
            </div>
        </div>

    </div>
    <!--container end-->

    <!--審核不通過 跳出fancybox輸入意見-->
    <div id="opblock" style="display: none; text-align: center;">
        <div style="margin-bottom: 10px;">請輸入審核不通過意見</div>
        <div style="margin-bottom: 10px;">
            <textarea id="textOp" rows="7" cols="50"></textarea>
        </div>
        <input type="button" id="submitbtn" value="確定" class="genbtn" />
    </div>

    <!--匯出-->
    <div id="exblock" style="display: none; text-align: center;">
        <div style="margin-bottom: 10px;">請選擇匯出檔案類型</div>
        <div style="margin-bottom: 10px;">
            <input type="radio" name="extype" value="WORD" checked="checked" />&nbsp;Word&nbsp;&nbsp;
            <input type="radio" name="extype" value="PDF" />&nbsp;PDF
        </div>
        <input type="button" id="expbtn" value="確定" class="genbtn" />
    </div>
    <input type="hidden" id="hidden_reportGuid" />
    <input type="hidden" id="hidden_val" />
    <input type="hidden" id="hidden_year" />
    <input type="hidden" id="hidden_month" />
    <input type="hidden" id="hidden_typeval" />

</asp:Content>

