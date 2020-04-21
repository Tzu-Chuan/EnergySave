<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="ReportMonthEx.aspx.cs" Inherits="WebPage_ReportMonthEx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script>
        //setInterval(function () {
        //    var breakStatus = false;
        //    if ($.getParamValue('stage') == "" || $.getParamValue('year') == null || $.getParamValue('year') == "" || $.getParamValue('month') == "") {
        //        breakStatus = true;
        //    }
        //    if ($("#hidden_chktype").val() == "Y" || ($("#hidden_chktype").val() == "" && $("#hidden_chkstatus").val() == "A")) {
        //        breakStatus = true;
        //    }
        //    if (breakStatus == true)
        //        return;
        //    //判斷如果現在是定稿或送審
        //    if ($("#btn_save").prop("disabled") || $("#txt_Type01Real").prop("disabled"))
        //        return;
        //    var iframe = $('<iframe name="postiframe" id="postiframe" style="display: none" />');
        //    var mid = $('<input type="hidden" name="mid" id="mid" value="' + $.getParamValue('v') + '" />');
        //    var stage = $('<input type="hidden" name="stage" id="stage" value="' + $.getParamValue('stage') + '" />');
        //    var year = $('<input type="hidden" name="year" id="year" value="' + $.getParamValue('year') + '" />');
        //    var month = $('<input type="hidden" name="month" id="month" value="' + $.getParamValue('month') + '" />');
        //    var form = $("form")[0];
        //    $("#postiframe").remove();
        //    $("input[name='mid']").remove();
        //    $("input[name='stage']").remove();
        //    $("input[name='year']").remove();
        //    $("input[name='month']").remove();
        //    form.appendChild(iframe[0]);
        //    form.appendChild(mid[0]);
        //    form.appendChild(stage[0]);
        //    form.appendChild(year[0]);
        //    form.appendChild(month[0]);
        //    form.setAttribute("action", "../handler/AutoSave_ReportMonth.ashx");
        //    form.setAttribute("method", "post");
        //    form.setAttribute("enctype", "multipart/form-data");
        //    form.setAttribute("encoding", "multipart/form-data");
        //    form.setAttribute("target", "postiframe");
        //    form.submit();
        //}, 1200000);//20 minutes 1200000

        $(function () {
            //$("#txt_stage").html("第" + $.getParamValue('stage') + "期");
            $("#txt_yearmonth").html($.getParamValue('year') + " 年" + $.getParamValue('month') + "月");
            monthFirstLastDate($.getParamValue('year'), $.getParamValue('month'));

            $(".pbtn").hide();
            $("#hidden_mid").val($.getParamValue('v'));
            $("#div_data").empty();
            //期數change
            load_stagedata($.getParamValue('stage'));
            load_data();
            load_citydatabymid();

            //儲存
            $(document).on("click", "#btn_save", function () {
                 saveFunc("save");
            });

            //數字的輸入欄位 keyup
            $(document).on("keyup", "input", function () {
                $("input").css("border-color", "");
                var thist = $(this).attr("t");
                var strVal = $(this).val();
                var re = /^(\+|-)?\d+$/;
                if ($(this).val() != "" && thist == "strint")//正整數
                {
                    if (re.test(strVal) && strVal >= 0) {
                        return true;
                    } else {
                        alert("請輸入數字且為正整數或零");
                        $(this).css("border-color", "red");
                        $(this).val("");
                        return false;
                    }
                }
                if ($(this).val() != "" && thist == "strflot")//小數
                {
                    if (strVal.indexOf(".") > 0) {//有小數點
                        var splitVal = strVal.split('.');
                        if (splitVal[1].length > 1) {
                            if (re.test(strVal) && strVal >= 0) {
                                return true;
                            } else {
                                alert("請輸入數字且只能到小數第一位");
                                $(this).css("border-color", "red");
                                $(this).val("");
                                return false;
                            }

                        } else {
                            if (splitVal[1] == "" || (splitVal[1] != "" && re.test(splitVal[1]) && splitVal[1] >= 0)) {
                                return true;
                            } else {
                                alert("請輸入數字且為正整數或零");
                                $(this).css("border-color", "red");
                                $(this).val("");
                                return false;
                            }
                        }
                    }
                    //else {
                    //    if (re.test(strVal) && strVal > 0) {
                    //        return true;
                    //    } else {
                    //        alert("請輸入數字且為正整數");
                    //        $(this).css("border-color", "red");
                    //        $(this).val("");
                    //        return false;
                    //    }
                    //}
                }
                return true;
            });
            ////數字欄位 blur 計算合計用
            //$(document).on("blur", "input", function () {
            //    if ($(this).attr("tp") == "f") {//年節電量input
            //        if ($(this).val()=="") {
            //            $(this).val("0");
            //        }
            //        sumValF($(this).attr("id"));
            //    } else {//其他input
            //        sumVal($(this).attr("id"));
            //    }
                
            //});
            $(document).on("change", "input", function () {
                if ($(this).attr("tp") == "f") {//年節電量input
                    if ($(this).val()=="") {
                        $(this).val("0");
                    }
                }

                sumVal($(this).attr("id"));
                saveFunc("autosave");
            });
             $(document).on("change", "textarea", function () {
                saveFunc("autosave");
            });

            //取消
            $("#btn_cancel").click(function () {
                if(confirm("即將放棄剛剛填寫過的資料，是否繼續?")){
                    location.href = "ReportMonthList.aspx";
                }
            });
            //送審
            $("#btn_goCheck").click(function () {
                if (confirm("是否確定送審?")) {
                    goCheck();
                }
            });
        });

        //撈帶過來的v<mid>資料 填表人 電話 主管 機關 局處
        function load_citydatabymid() {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/ReportMonth.ashx",
                data: {
                    func: "load_projectbyperson"//,
                    //str_mid: $.getParamValue('v')
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
        //撈期程資料
        function load_stagedata(strStage) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/ReportMonth.ashx",
                data: {
                    func: "load_stagedate",
                    str_stage: strStage
                    //str_mid: $.getParamValue('v')

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
                                var str = "";
                                var sdate = $(this).children("sdate").text().trim();
                                var edate = $(this).children("edate").text().trim();
                                if (sdate != "" && edate != "") {
                                    var splits = sdate.split("/");
                                    var splite = edate.split("/");
                                    //自107年1月1日起至 年 月 日止，計    個月
                                    str += "自" + (parseInt(splits[0]) - 1911) + "年" + splits[1] + "月" + splits[2] + "日起至";
                                    str += "自" + (parseInt(splite[0]) - 1911) + "年" + splite[1] + "月" + splite[2] + "日止，共計";
                                    str += monthDiff(new Date(sdate), new Date(edate)) + "個月";
                                }
                                
                                $("#txt_datesanddatee").empty().append(str);
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
                url: "../handler/ReportMonth.ashx",
                data: {
                    func: "load_MonthReportData",//txt_stage
                    str_stage: $.getParamValue('stage'),
                    str_year: $.getParamValue('year'),
                    str_month: $.getParamValue('month'),
                    str_rmptype: $.getParamValue('rmtype')
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
                    }
                    else {
                        var strHtml = "";
                        var chk_status;
                        var chk_type;
                        var chk_date;
                        var strReportGuiud = "";
                        data = $.parseXML(data);
                        if (data == null) {
                            //沒資料連button都不顯示
                            disabledInputandBtn();
                            $("#div_data").empty().append("<font size='5' color='red'>尚無查核點及預訂工作進度資料，請至計畫基本資料確認</font>");
                            return;
                        }
                        if ($(data).find("data_item").length > 0) {
                            var strCreateDate;
                            var strChkDate;
                            //防止ReportGuid不一致，有可能在儲存完月報之後再新增推動項目，
                            //這樣新的項目的reportGuid就會是空，後端新增後就會變成新的reportGuid
                            $(data).find("data_item").each(function (i) {
                                if (strReportGuiud == "") {
                                    strReportGuiud = $(this).children("RM_ReportGuid").text().trim()
                                }
                            });

                            $(data).find("data_item").each(function (i) {
                                var strday = new Date();
                                //如果不過濾掉，chk_status會變成空直，沒辦法判斷現在到底是不是送審中
                                if (i == 0) {
                                    chk_status = $(this).children("RC_Status").text().trim();
                                    chk_type = $(this).children("RC_CheckType").text().trim();
                                    //填表日 如果還沒審核過 填表日永遠是填表當下的日期
                                    if ($(this).children("RC_CheckDate").text().trim() != "" && $(this).children("RC_CheckDate").text().trim() != "1900-01-01T00:00:00+08:00") {
                                        chk_date = formatDate($(this).children("RC_CheckDate").text().trim());
                                        strCreateDate = formatDate($(this).children("RM_ModDate").text().trim());
                                    } else {
                                        chk_date = "";
                                        strCreateDate = strday.getFullYear() + "/" + (strday.getMonth() + 1) + "/" + strday.getDate();
                                    }
                                }
                                strHtml += getHtml($(this),strReportGuiud);
                                
                            });
                            //審核資料
                            if ($(data).find("people_item").length > 0) {
                                $(data).find("people_item").each(function (i) {
                                    strHtml += '<div class="OchiRow">';
                                    strHtml += '<div class="OchiThird">';
                                    strHtml += '<div class="OchiCell OchiTitle TitleSetWidth">填表人</div>';
                                    strHtml += '<div class="OchiCell width100" id="div_person">' + $(this).children("M_Name").text().trim() + '</div>';
                                    strHtml += '</div>';
                                    strHtml += '<div class="OchiThird">';
                                    strHtml += '<div class="OchiCell OchiTitle TitleSetWidth">電話</div>';
                                    strHtml += '<div class="OchiCell width100" id="div_phone">' + $(this).children("M_Tel").text().trim() + '</div>';
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
                                    strHtml += '<div class="OchiCellInner width100" id="div_dossname">' + $(this).children("bossname").text().trim() + '</div>';
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
                                });
                            }
                            $("#div_data").empty().append(strHtml);
                            $("#hidden_chktype").val(chk_type);
                            $("#hidden_chkstatus").val(chk_status);
                            //審核中 & 已審核通過都不能修改
                            if (chk_status == "A") {
                                disabledInputandBtn();
                            } else {
                                showbutton();
                            }
                            changeZero();
                            
                        }
                    }
                }
            });//ajax end
        }

        //湊出各類別表格
        function getHtml(xmlData,strRGuiud) {
            var strItem = xmlData.children("P_ItemName").text().trim();
            var type1 = xmlData.children("P_ExType").text().trim();
            var type2 = xmlData.children("P_ExDeviceType").text().trim();
            var ItemCname = xmlData.children("P_ItemName_c").text().trim();
            var str = "", strUnit = "", strFo = "每單位年節電量(度)";
            var splitRM_Finish, splitRM_Finish01, P_ExFinish;
            var ptno = xmlData.children("P_ItemName").text().trim();
            var intflort = "";

            P_ExFinish = xmlData.children("P_ExFinish").text().trim().split('.');
            splitRM_Finish = xmlData.children("countFinishKW").text().trim().split('.');
            splitRM_Finish01 = xmlData.children("countFinish02").text().trim().split('.');
            strP_ExFinish = P_ExFinish[0];
            strRM_Finish = splitRM_Finish[0];
            strRM_Finish01 = splitRM_Finish01[0];
            
            var strcountApplyKW = xmlData.children("countApplyKW").text().trim();//累計申請數KW
            var strcountApply01 = xmlData.children("countApply01").text().trim();//累計申請數

            

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


            if (strItem == "01" || strItem == "33" ) {
                strFo = "每單位年節電量(度/kW)";
                str += '<div class="OchiRow">';
                str += '<div class="OchiCell OchiTitle TitleSetWidth">' + ItemCname + '</div>';
                str += '<div class="OchiCell width100">';
                str += '<div class="OchiTableInner width100">';
                str += '<div class="OchiCellInner nowrap textcenter">本期累計申請數：</div>';
                str += '<div class="OchiCellInner width50">';
                str += '<span>' + (strcountApplyKW == "" ? "0" : strcountApplyKW) + '&nbsp;(kW)</span>';
                str += '<input type="text" class="inputex" size="18" name="' + ptno + 'RM_ApplyKW" value="' + (strcountApplyKW == "" ? "0" : strcountApplyKW) + '" t = "strflot" readonly = "readonly" style = "background-color:#DDDDDD;display:none;" />';
                str += '</div>';
                str += '<div class="OchiCellInner nowrap textcenter">本期規劃數：</div>';
                str += '<div class="OchiCellInner width50">';
                str += '<span>' + (strP_ExFinish == "" ? "0" : strP_ExFinish) + '</span>&nbsp;(kW)';
                str += '<input type="text" class="inputex" size="18" name="' + ptno + 'RM_Planning" value="' + (strP_ExFinish == "" ? "0" : strP_ExFinish) + '" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</div>';
                str += '</div>';
                str += '<div class="OchiTableInner width100 margin5T">';
                str += '<div class="OchiTableInner width100">';
                str += '<div class="OchiCellInner nowrap textcenter">本期累計完成數：</div>';
                str += '<div class="OchiCellInner width50">';
                str += '<span>' + (strRM_Finish == "" ? "0" : strRM_Finish) + '&nbsp;(kW)</span>';
                str += '<input type="text" class="inputex" size="18" name="' + ptno + 'RM_Finish" value="' + (strRM_Finish == "" ? "0" : strRM_Finish) + '" t="strflot" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</div>';
                str += '<div class="OchiCellInner nowrap textcenter">本期累計核定數：</div>';
                str += '<div class="OchiCellInner width50">';
                str += '<span>' + (strRM_Finish01 == "" ? "0" : strRM_Finish01) + '&nbsp;(台)</span>';
                str += '<input type="text" class="inputex" size="12" name="' + ptno + 'RM_Finish01" value="' + (strRM_Finish01 == "" ? "0" : strRM_Finish01) + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
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
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1Value1" id="' + ptno + 'RM_Type1Value1" value="' + xmlData.children("RM_Type1Value1").text().trim() + '" t="strint" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1Value2" id="' + ptno + 'RM_Type1Value2" value="' + xmlData.children("RM_Type1Value2").text().trim() + '" t="strint" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1Value3" id="' + ptno + 'RM_Type1Value3" value="' + xmlData.children("RM_Type1Value3").text().trim() + '" t="strint" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2Value1" id="' + ptno + 'RM_Type2Value1" value="' + xmlData.children("RM_Type2Value1").text().trim() + '" t="strint" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2Value2" id="' + ptno + 'RM_Type2Value2" value="' + xmlData.children("RM_Type2Value2").text().trim() + '" t="strint" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2Value3" id="' + ptno + 'RM_Type2Value3" value="' + xmlData.children("RM_Type2Value3").text().trim() + '" t="strint" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3Value1" id="' + ptno + 'RM_Type3Value1" value="' + xmlData.children("RM_Type3Value1").text().trim() + '" t="strflot" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3Value2" id="' + ptno + 'RM_Type3Value2" value="' + xmlData.children("RM_Type3Value2").text().trim() + '" t="strflot" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3Value3" id="' + ptno + 'RM_Type3Value3" value="' + xmlData.children("RM_Type3Value3").text().trim() + '" t="strflot" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type4Value1" id="' + ptno + 'RM_Type4Value1" value="' + xmlData.children("RM_Type4Value1").text().trim() + '" t="strflot" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type4Value2" id="' + ptno + 'RM_Type4Value2" value="' + xmlData.children("RM_Type4Value2").text().trim() + '" t="strflot" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type4Value3" id="' + ptno + 'RM_Type4Value3" value="' + xmlData.children("RM_Type4Value3").text().trim() + '" t="strflot" /></td>';
                str += '</tr>';
                str += '<tr>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">';
                str += '<span id="' + ptno + 'RM_Type1ValueSum_span">' + (xmlData.children("RM_Type1ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type1ValueSum").text().trim()) + '</span>';
                str += '<input type="text" class="inputex width100" name="' + ptno + 'RM_Type1ValueSum" id="' + ptno + 'RM_Type1ValueSum" value="' + (xmlData.children("RM_Type1ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type1ValueSum").text().trim()) + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</td>';
                str += '<td align="right">合計</td>';                                                                                                                                                                                                                                                                                                                                                                  
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">';
                str += '<span id="' + ptno + 'RM_Type2ValueSum_span">' + (xmlData.children("RM_Type2ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type2ValueSum").text().trim()) + '</span>';
                str += '<input type="text" class="inputex width100" name="' + ptno + 'RM_Type2ValueSum" id="' + ptno + 'RM_Type2ValueSum" value="' + (xmlData.children("RM_Type2ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type2ValueSum").text().trim()) + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</td>';
                str += '<td align="right">合計</td>';                                                                                                                                                                                                                                                                                                                                                                  
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">';
                str += '<span id="' + ptno + 'RM_Type3ValueSum_span">' + (xmlData.children("RM_Type3ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type3ValueSum").text().trim()) + '</span>';
                str += '<input type="text" class="inputex width100" name="' + ptno + 'RM_Type3ValueSum" id="' + ptno + 'RM_Type3ValueSum" value="' + (xmlData.children("RM_Type3ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type3ValueSum").text().trim()) + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</td>';
                str += '<td align="right">合計</td>';                                                                                                                                                                                                                                                                                                                                                                  
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">';
                str += '<span id="' + ptno + 'RM_Type4ValueSum_span">' + (xmlData.children("RM_Type4ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type4ValueSum").text().trim()) + '</span>';
                str += '<input type="text" class="inputex width100" name="' + ptno + 'RM_Type4ValueSum" id="' + ptno + 'RM_Type4ValueSum" value="' + (xmlData.children("RM_Type4ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type4ValueSum").text().trim()) + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</td>';
                str += '</tr>';
                str += '<tr>';
                str += '<th colspan="4">本月申請數預期年節電量(度)</th>';
                str += '<th colspan="4">本月核定數預期年節電量(度)</th>';
                str += '<th colspan="4">本月未核定數之年節電量(度) </th>';
                str += '</tr>';
                str += '<tr>';
                //如果是"其他"，因為沒有公式，所以讓使用者自己輸入計算後的數值
                //if (ptno.length > 2) {
                //    str += '<td colspan="4" style="text-align:right;"><inputclass="inputex width100" type="text" name="' + ptno + 'RM_PreVal" id="' + ptno + 'RM_PreVal" value="' + (xmlData.children("RM_PreVal").text().trim() == "" ? "0" : xmlData.children("RM_PreVal").text().trim()) + '" t="strflot" /></td>';
                //    str += '<td colspan="4" style="text-align:right;"><inputclass="inputex width100" type="text" name="' + ptno + 'RM_ChkVal" id="' + ptno + 'RM_ChkVal" value="' + (xmlData.children("RM_ChkVal").text().trim() == "" ? "0" : xmlData.children("RM_ChkVal").text().trim()) + '" t="strflot" /></td>';
                //    str += '<td colspan="4" style="text-align:right;"><inputclass="inputex width100" type="text"  name="' + ptno + 'RM_NotChkVal" id="' + ptno + 'RM_NotChkVal" value="' + (xmlData.children("RM_NotChkVal").text().trim() == "" ? "0" : xmlData.children("RM_NotChkVal").text().trim()) + '" t="strflot" /></td>';
                //} else {
                    str += '<td colspan="4" style="text-align:right;"><span id="' + ptno + 'RM_PreVal_span">' + (xmlData.children("RM_PreVal").text().trim() == "" ? "0" : xmlData.children("RM_PreVal").text().trim()) + '</span><input type="text" name="' + ptno + 'RM_PreVal" id="' + ptno + 'RM_PreVal" value="' + (xmlData.children("RM_PreVal").text().trim() == "" ? "0" : xmlData.children("RM_PreVal").text().trim()) + '" t="strflot" style="background-color:#DDDDDD;display:none;" /></td>';
                    str += '<td colspan="4" style="text-align:right;"><span id="' + ptno + 'RM_ChkVal_span">' + (xmlData.children("RM_ChkVal").text().trim() == "" ? "0" : xmlData.children("RM_ChkVal").text().trim()) + '</span><input type="text" name="' + ptno + 'RM_ChkVal" id="' + ptno + 'RM_ChkVal" value="' + (xmlData.children("RM_ChkVal").text().trim() == "" ? "0" : xmlData.children("RM_ChkVal").text().trim()) + '" t="strflot" style="background-color:#DDDDDD;display:none;" /></td>';
                    str += '<td colspan="4" style="text-align:right;"><span id="' + ptno + 'RM_NotChkVal_span">' + (xmlData.children("RM_NotChkVal").text().trim() == "" ? "0" : xmlData.children("RM_NotChkVal").text().trim()) + '</span><input type="text"  name="' + ptno + 'RM_NotChkVal" id="' + ptno + 'RM_NotChkVal" value="' + (xmlData.children("RM_NotChkVal").text().trim() == "" ? "0" : xmlData.children("RM_NotChkVal").text().trim()) + '" t="strflot" style="background-color:#DDDDDD;display:none;" /></td>';
                //}
                str += '</tr>';
                str += '<tr>';
                //if ($.getParamValue('month') == "01" || xmlData.children("RM_Formula").text().trim() == "") {
                str += '<td colspan="2" style="text-align:right;">' + strFo + '：</td><td colspan="10"><input type="text" value="' + xmlData.children("RM_Formula").text().trim() + '" class="inputex width15" id="' + ptno + 'RM_Formula" name="' + ptno + 'RM_Formula" tp="f" /></td>';
                //} else {
                    //str += '<td colspan="2" style="text-align:right;">年節電量(度)：</td><td colspan="10"><input type="text" value="' + xmlData.children("RM_Formula").text().trim() + '" class="inputex width15" id="' + ptno + 'RM_Formula" name="' + ptno + 'RM_Formula" tp="f" readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                //}

                str += '</tr>';
                str += '</table>';
                str += '</div>';
                str += '</div>';
                str += '<div class="OchiTableInner width100 margin5TB">';
                str += '補充說明：<textarea rows="3" class="inputex width100" name="' + ptno + 'RM_Remark">' + xmlData.children("RM_Remark").text().trim() + '</textarea>';
                str += '</div>';
                str += '</div>';
                str += '</div>';
                str += '<div><input type="hidden" name="report_type" value="' + ptno + '" /></div>';
                str += '<div><input type="hidden" name="report_P_Guid" value="' + xmlData.children("P_Guid").text().trim() + '" /></div>';
                str += '<div><input type="hidden" name="report_Guid" value="' + strRGuiud + '" /></div>';
            }
            else if (strItem == "02" || strItem == "03") {//type1 == "02" && type2 == "02" =>照明
                str += '<div class="OchiRow">';
                str += '<div class="OchiCell OchiTitle TitleSetWidth">' + ItemCname + '</div>';
                str += '<div class="OchiCell width100">';
                str += '<div class="OchiTableInner width100">';
                str += '<div class="OchiCellInner nowrap textcenter">本期累計核定數：</div>';
                str += '<div class="OchiCellInner width33">';
                str += '<span>' + (splitRM_Finish01 == "" ? "0" : splitRM_Finish01) + '&nbsp;(盞)</span>';
                str += '<input type="text" class="inputex" size="12" name="' + ptno + 'RM_Finish" value="' + (splitRM_Finish01 == "" ? "0" : splitRM_Finish01) + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</div>';
                str += '<div class="OchiCellInner nowrap textcenter">本期累計申請數：</div>';
                str += '<div class="OchiCellInner width33">';
                str += '<span>' + (strcountApply01 == "" ? "0" : strcountApply01) + '&nbsp;(盞)</span>';
                str += '<input type="text" class="inputex" size="12" name="' + ptno + 'RM_Apply01" value="' + (strcountApply01 == "" ? "0" : strcountApply01) + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</div>';
                str += '<div class="OchiCellInner nowrap textcenter">本期規劃數：</div>';
                str += '<div class="OchiCellInner width50">';
                str += '<span>' + (strP_ExFinish == "" ? "0" : strP_ExFinish) + '&nbsp;(盞)</span>';
                str += '<input type="text" class="inputex" size="18" name="' + ptno + 'RM_Planning" value="' + (strP_ExFinish == "" ? "0" : strP_ExFinish) + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
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
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1Value1" id="' + ptno + 'RM_Type1Value1" value="' + xmlData.children("RM_Type1Value1").text().trim() + '" t="strint" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1Value2" id="' + ptno + 'RM_Type1Value2" value="' + xmlData.children("RM_Type1Value2").text().trim() + '" t="strint" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1Value3" id="' + ptno + 'RM_Type1Value3" value="' + xmlData.children("RM_Type1Value3").text().trim() + '" t="strint" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2Value1" id="' + ptno + 'RM_Type2Value1" value="' + xmlData.children("RM_Type2Value1").text().trim() + '" t="strint" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2Value2" id="' + ptno + 'RM_Type2Value2" value="' + xmlData.children("RM_Type2Value2").text().trim() + '" t="strint" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2Value3" id="' + ptno + 'RM_Type2Value3" value="' + xmlData.children("RM_Type2Value3").text().trim() + '" t="strint" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3Value1" id="' + ptno + 'RM_Type3Value1" value="' + xmlData.children("RM_Type3Value1").text().trim() + '" t="strflot" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3Value2" id="' + ptno + 'RM_Type3Value2" value="' + xmlData.children("RM_Type3Value2").text().trim() + '" t="strflot" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3Value3" id="' + ptno + 'RM_Type3Value3" value="' + xmlData.children("RM_Type3Value3").text().trim() + '" t="strflot" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type4Value1" id="' + ptno + 'RM_Type4Value1" value="' + xmlData.children("RM_Type4Value1").text().trim() + '" t="strflot" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type4Value2" id="' + ptno + 'RM_Type4Value2" value="' + xmlData.children("RM_Type4Value2").text().trim() + '" t="strflot" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type4Value3" id="' + ptno + 'RM_Type4Value3" value="' + xmlData.children("RM_Type4Value3").text().trim() + '" t="strflot" /></td>';
                str += '</tr>';
                str += '<tr>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">';
                str += '<span id="' + ptno + 'RM_Type1ValueSum_span">' + (xmlData.children("RM_Type1ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type1ValueSum").text().trim()) + '</span>';
                str += '<input type="text" class="inputex width100" name="' + ptno + 'RM_Type1ValueSum" id="' + ptno + 'RM_Type1ValueSum" value="' + (xmlData.children("RM_Type1ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type1ValueSum").text().trim()) + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</td>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">';
                str += '<span id="' + ptno + 'RM_Type2ValueSum_span">' + (xmlData.children("RM_Type1ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type2ValueSum").text().trim()) + '</span>';
                str += '<input type="text" class="inputex width100" name="' + ptno + 'RM_Type2ValueSum" id="' + ptno + 'RM_Type2ValueSum" value="' + (xmlData.children("RM_Type2ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type2ValueSum").text().trim()) + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</td>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">';
                str += '<span id="' + ptno + 'RM_Type3ValueSum_span">' + (xmlData.children("RM_Type1ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type3ValueSum").text().trim()) + '</span>';
                str += '<input type="text" class="inputex width100" name="' + ptno + 'RM_Type3ValueSum" id="' + ptno + 'RM_Type3ValueSum" value="' + (xmlData.children("RM_Type3ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type3ValueSum").text().trim()) + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</td>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">';
                str += '<span id="' + ptno + 'RM_Type4ValueSum_span">' + (xmlData.children("RM_Type1ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type4ValueSum").text().trim()) + '</span>';
                str += '<input type="text" class="inputex width100" name="' + ptno + 'RM_Type4ValueSum" id="' + ptno + 'RM_Type4ValueSum" value="' + (xmlData.children("RM_Type4ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type4ValueSum").text().trim()) + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</td>';
                str += '</tr>';
                str += '<tr>';
                str += '<th colspan="4">本月申請數預期年節電量(度)</th>';
                str += '<th colspan="4">本月核定數預期年節電量(度)</th>';
                str += '<th colspan="4">本月未核定數之年節電量(度) </th>';
                str += '</tr>';
                str += '<tr>';
                //如果是"其他"，因為沒有公式，所以讓使用者自己輸入計算後的數值
                //if (ptno.length > 2) {
                //    str += '<td colspan="4" style="text-align:right;"><input type="text" class="inputex width100" name="' + ptno + 'RM_PreVal" id="' + ptno + 'RM_PreVal" value="' + (xmlData.children("RM_PreVal").text().trim() == "" ? "0" : xmlData.children("RM_PreVal").text().trim()) + '" t="strflot" /></td>';
                //    str += '<td colspan="4" style="text-align:right;"><input type="text" class="inputex width100" name="' + ptno + 'RM_ChkVal" id="' + ptno + 'RM_ChkVal" value="' + (xmlData.children("RM_ChkVal").text().trim() == "" ? "0" : xmlData.children("RM_ChkVal").text().trim()) + '" t="strflot" /></td>';
                //    str += '<td colspan="4" style="text-align:right;"><input type="text" class="inputex width100" name="' + ptno + 'RM_NotChkVal" id="' + ptno + 'RM_NotChkVal" value="' + (xmlData.children("RM_NotChkVal").text().trim() == "" ? "0" : xmlData.children("RM_NotChkVal").text().trim()) + '" t="strflot" /></td>';
                //} else {
                    str += '<td colspan="4" style="text-align:right;"><span id="' + ptno + 'RM_PreVal_span">' + (xmlData.children("RM_PreVal").text().trim() == "" ? "0" : xmlData.children("RM_PreVal").text().trim()) + '</span><input type="text" class="inputex width100" name="' + ptno + 'RM_PreVal" id="' + ptno + 'RM_PreVal" value="' + (xmlData.children("RM_PreVal").text().trim() == "" ? "0" : xmlData.children("RM_PreVal").text().trim()) + '" t="strflot" readonly="readonly" style="background-color:#DDDDDD;display:none;" /></td>';
                    str += '<td colspan="4" style="text-align:right;"><span id="' + ptno + 'RM_ChkVal_span">' + (xmlData.children("RM_ChkVal").text().trim() == "" ? "0" : xmlData.children("RM_ChkVal").text().trim()) + '</span><input type="text" class="inputex width100" name="' + ptno + 'RM_ChkVal" id="' + ptno + 'RM_ChkVal" value="' + (xmlData.children("RM_ChkVal").text().trim() == "" ? "0" : xmlData.children("RM_ChkVal").text().trim()) + '" t="strflot" readonly="readonly" style="background-color:#DDDDDD;display:none;" /></td>';
                    str += '<td colspan="4" style="text-align:right;"><span id="' + ptno + 'RM_NotChkVal_span">' + (xmlData.children("RM_NotChkVal").text().trim() == "" ? "0" : xmlData.children("RM_NotChkVal").text().trim()) + '</span><input type="text" class="inputex width100" name="' + ptno + 'RM_NotChkVal" id="' + ptno + 'RM_NotChkVal" value="' + (xmlData.children("RM_NotChkVal").text().trim() == "" ? "0" : xmlData.children("RM_NotChkVal").text().trim()) + '" t="strflot" readonly="readonly" style="background-color:#DDDDDD;display:none;" /></td>';
                //}
                str += '</tr>';
                str += '<tr>';
                //if ($.getParamValue('month') == "01" || xmlData.children("RM_Formula").text().trim() == "") {
                    str += '<td colspan="2" style="text-align:right;">'+strFo+'：</td><td colspan="10"><input type="text" value="' + xmlData.children("RM_Formula").text().trim() + '" class="inputex width15" id="' + ptno + 'RM_Formula" name="' + ptno + 'RM_Formula" tp="f" /></td>';
                //} else {
                    //str += '<td colspan="2" style="text-align:right;">年節電量(度)：</td><td colspan="10"><input type="text" value="' + xmlData.children("RM_Formula").text().trim() + '" class="inputex width15" id="' + ptno + 'RM_Formula" name="' + ptno + 'RM_Formula" tp="f" readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                //}

                str += '</tr>';
                str += '</table>';
                str += '</div>';
                str += '</div>';
                str += '<div class="OchiTableInner width100 margin5TB">';
                str += '補充說明：';
                str += '<textarea rows="3" class="inputex width100" name="' + ptno + 'RM_Remark">' + xmlData.children("RM_Remark").text().trim() + '</textarea>';
                str += '</div>';
                str += '</div>';
                str += '</div>';
                str += '<div><input type="hidden" name="report_type" value="' + ptno + '" /></div>';
                str += '<div><input type="hidden" name="report_P_Guid" value="' + xmlData.children("P_Guid").text().trim() + '" /></div>';
                str += '<div><input type="hidden" name="report_Guid" value="' + strRGuiud + '" /></div>';
            }
            else {
                var splitcountFinish03 = "", countFinish03 = "";
                if (strItem == "05" || strItem == "23") {
                    //strUnit = "KW";
                    strFo = "每單位年節電量(度/kW)";
                    intflort = "strflot";
                    countFinish03 = xmlData.children("countFinish03").text().trim();
                } else if (strItem == "14") {
                    //strUnit = "組";
                    intflort = "strint";
                    splitcountFinish03 = xmlData.children("countFinish03").text().trim().split('.');
                    countFinish03 = splitcountFinish03[0];
                } else {
                    //strUnit = "台";
                    intflort = "strint";
                    splitcountFinish03 = xmlData.children("countFinish03").text().trim().split('.');
                    countFinish03 = splitcountFinish03[0];
                }
                
                
                str += '<div class="OchiRow">';
                str += '<div class="OchiCell OchiTitle TitleSetWidth">' + ItemCname + '</div>';
                str += '<div class="OchiCell width100">';
                str += '<div class="OchiTableInner width100">';
                str += '<div class="OchiCellInner nowrap textcenter">本期累計完成數:</div>';
                str += '<div class="OchiCellInner width33">';
                str += '<span>' + (countFinish03 == "" ? "0" : countFinish03) + '&nbsp;(' + strUnit + ')</span>';
                str += '<input type="text" class="inputex" size="12" name="' + ptno + 'RM_Finish" value="' + countFinish03 + '" t="' + intflort + '" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</div>';
                str += '<div class="OchiCellInner nowrap textcenter">本期累計申請數:</div>';
                str += '<div class="OchiCellInner width33"><span>' + strcountApply01 + '&nbsp;(' + strUnit + ')</span><input type="text" class="inputex" size="12" name="' + ptno + 'RM_Apply01" value="' + strcountApply01 + '" t="' + intflort + '" readonly="readonly" style="background-color:#DDDDDD;display:none;" /></div>';
                str += '<div class="OchiCellInner nowrap textcenter">本期規劃數:</div>';
                str += '<div class="OchiCellInner width33">';
                str += '<span>' + (strP_ExFinish == "" ? "0" : strP_ExFinish) + '&nbsp;(' + strUnit + ')</span>';
                str += '<input type="text" class="inputex" size="12" name="' + ptno + 'RM_Planning" value="' + strP_ExFinish + '" t="' + intflort + '" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
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
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1Value1" id="' + ptno + 'RM_Type1Value1" value="' + xmlData.children("RM_Type1Value1").text().trim() + '" t="' + intflort + '" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1Value2" id="' + ptno + 'RM_Type1Value2" value="' + xmlData.children("RM_Type1Value2").text().trim() + '" t="' + intflort + '" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1Value3" id="' + ptno + 'RM_Type1Value3" value="' + xmlData.children("RM_Type1Value3").text().trim() + '" t="' + intflort + '" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2Value1" id="' + ptno + 'RM_Type2Value1" value="' + xmlData.children("RM_Type2Value1").text().trim() + '" t="' + intflort + '" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2Value2" id="' + ptno + 'RM_Type2Value2" value="' + xmlData.children("RM_Type2Value2").text().trim() + '" t="' + intflort + '" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2Value3" id="' + ptno + 'RM_Type2Value3" value="' + xmlData.children("RM_Type2Value3").text().trim() + '" t="' + intflort + '" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3Value1" id="' + ptno + 'RM_Type3Value1" value="' + xmlData.children("RM_Type3Value1").text().trim().replace(".0","") + '" t="' + intflort + '" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3Value2" id="' + ptno + 'RM_Type3Value2" value="' + xmlData.children("RM_Type3Value2").text().trim().replace(".0","") + '" t="' + intflort + '" /></td>';
                str += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3Value3" id="' + ptno + 'RM_Type3Value3" value="' + xmlData.children("RM_Type3Value3").text().trim().replace(".0","") + '" t="' + intflort + '" /></td>';
                str += '</tr>';
                str += '<tr>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">';
                str += '<span id="' + ptno + 'RM_Type1ValueSum_span">' + (xmlData.children("RM_Type1ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type1ValueSum").text().trim()) + '</span>';
                str += '<input type="text" class="inputex width100" name="' + ptno + 'RM_Type1ValueSum" id="' + ptno + 'RM_Type1ValueSum" value="' + (xmlData.children("RM_Type1ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type1ValueSum").text().trim()) + '" t="' + intflort + '"  readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</td>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">';
                str += '<span id="' + ptno + 'RM_Type2ValueSum_span">' + (xmlData.children("RM_Type2ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type2ValueSum").text().trim()) + '</span>';
                str += '<input type="text" class="inputex width100" name="' + ptno + 'RM_Type2ValueSum" id="' + ptno + 'RM_Type2ValueSum" value="' + (xmlData.children("RM_Type2ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type2ValueSum").text().trim()) + '" t="' + intflort + '"  readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</td>';
                str += '<td align="right">合計</td>';
                str += '<td colspan="2" style="text-align:right; background-color:#FEFBC2;">';
                str += '<span id="' + ptno + 'RM_Type3ValueSum_span">' + (xmlData.children("RM_Type3ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type3ValueSum").text().trim().replace(".0","")) + '</span>';
                str += '<input type="text" class="inputex width100" name="' + ptno + 'RM_Type3ValueSum" id="' + ptno + 'RM_Type3ValueSum" value="' + (xmlData.children("RM_Type3ValueSum").text().trim() == "" ? "0" : xmlData.children("RM_Type3ValueSum").text().trim().replace(".0","")) + '" t="' + intflort + '" readonly="readonly" style="background-color:#DDDDDD;display:none;" />';
                str += '</td>';
                str += '</tr>';
                str += '<tr>';
                str += '<th colspan="3">本月申請數預期年節電量(度)</th>';
                str += '<th colspan="3">本月核定數預期年節電量(度)</th>';
                str += '<th colspan="3">本月未核定數之年節電量(度) </th>';
                str += '</tr>';
                str += '<tr>';
                //如果是"其他"，因為沒有公式，所以讓使用者自己輸入計算後的數值
                //if (ptno.length > 2) {
                //  str += '<td colspan="3"><input type="text" class="inputex width100" name="' + ptno + 'RM_PreVal" id="' + ptno + 'RM_PreVal" value="' + (xmlData.children("RM_PreVal").text().trim() == "" ? "0" : xmlData.children("RM_PreVal").text().trim()) + '" t="strfloat" /></td>';
                //  str += '<td colspan="3"><input type="text" class="inputex width100" name="' + ptno + 'RM_ChkVal" id="' + ptno + 'RM_ChkVal" value="' + (xmlData.children("RM_ChkVal").text().trim() == "" ? "0" : xmlData.children("RM_ChkVal").text().trim()) + '" t="strfloat" /></td>';
                //   str += '<td colspan="3"><input type="text" class="inputex width100" name="' + ptno + 'RM_NotChkVal" id="' + ptno + 'RM_NotChkVal" value="' + (xmlData.children("RM_NotChkVal").text().trim() == "" ? "0" : xmlData.children("RM_NotChkVal").text().trim()) + '" t="strfloat" /></td>';
                //} else {
                    str += '<td colspan="3"><span id="' + ptno + 'RM_PreVal_span">' + (xmlData.children("RM_PreVal").text().trim() == "" ? "0" : xmlData.children("RM_PreVal").text().trim()) + '</span><input type="text" class="inputex width100" name="' + ptno + 'RM_PreVal" id="' + ptno + 'RM_PreVal" value="' + (xmlData.children("RM_PreVal").text().trim() == "" ? "0" : xmlData.children("RM_PreVal").text().trim()) + '" t="strfloat" readonly="readonly" style="background-color:#DDDDDD;display:none;" /></td>';
                    str += '<td colspan="3"><span id="' + ptno + 'RM_ChkVal_span">' + (xmlData.children("RM_ChkVal").text().trim() == "" ? "0" : xmlData.children("RM_ChkVal").text().trim()) + '</span><input type="text" class="inputex width100" name="' + ptno + 'RM_ChkVal" id="' + ptno + 'RM_ChkVal" value="' + (xmlData.children("RM_ChkVal").text().trim() == "" ? "0" : xmlData.children("RM_ChkVal").text().trim()) + '" t="strfloat" readonly="readonly" style="background-color:#DDDDDD;display:none;" /></td>';
                    str += '<td colspan="3"><span id="' + ptno + 'RM_NotChkVal_span">' + (xmlData.children("RM_NotChkVal").text().trim() == "" ? "0" : xmlData.children("RM_NotChkVal").text().trim()) + '</span><input type="text" class="inputex width100" name="' + ptno + 'RM_NotChkVal" id="' + ptno + 'RM_NotChkVal" value="' + (xmlData.children("RM_NotChkVal").text().trim() == "" ? "0" : xmlData.children("RM_NotChkVal").text().trim()) + '" t="strfloat" readonly="readonly" style="background-color:#DDDDDD;display:none;" /></td>';
                //}
                str += '</tr>';
                str += '<tr>';
                //if ($.getParamValue('month') == "01" || xmlData.children("RM_Formula").text().trim() == "") {
                    str += '<td colspan="2" style="text-align:right;">'+strFo+'：</td><td colspan="10"><input type="text" value="' + xmlData.children("RM_Formula").text().trim() + '" class="inputex width15" id="' + ptno + 'RM_Formula" name="' + ptno + 'RM_Formula" tp="f" /></td>';
                //} else {
                    //str += '<td colspan="2" style="text-align:right;">年節電量(度)：</td><td colspan="10"><input type="text" value="' + xmlData.children("RM_Formula").text().trim() + '" class="inputex width15" id="' + ptno + 'RM_Formula" name="' + ptno + 'RM_Formula" tp="f" readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                //}
                
                str += '</tr>';
                str += '</table>';
                str += '</div>';
                str += '</div>';
                str += '<div class="OchiTableInner width100 margin5T">補充說明：<textarea rows="3" class="inputex width100" name="' + ptno + 'RM_Remark">' + xmlData.children("RM_Remark").text().trim() + '</textarea></div>';
                str += '</div>';
                str += '</div>';
                str += '<div><input type="hidden" name="report_type" value="' + ptno + '" /></div>';
                str += '<div><input type="hidden" name="report_P_Guid" value="' + xmlData.children("P_Guid").text().trim() + '" /></div>';
                str += '<div><input type="hidden" name="report_Guid" value="' + strRGuiud + '" /></div>';
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

        //撈提報年分
        function load_year() {
            var startyear = 107;
            var endyear = startyear + 2;
            var str = '<option value="">請選擇</option>';
            for (var i = startyear; i <= endyear; i++) {
                str += '<option value="' + (i + 1911) + '">' + i + '</option>';
            }
            $("#txt_year").empty().append(str);
        }

        //儲存  點儲存按鈕 strType="save"  自動存檔 strType="autosave"
        function saveFunc(strType) {
            if ($.getParamValue('stage') == "" || $.getParamValue('year') == "" || $.getParamValue('month') == ""|| $.getParamValue('rmtype') == "") {
                    alert("參數錯誤，請重新操作");
                    location.href = "ReportMonthList.aspx";
                    return;
                }
                var iframe = $('<iframe name="postiframe" id="postiframe" style="display: none" />');
                var mid = $('<input type="hidden" name="mid" id="mid" value="' + $.getParamValue('v') + '" />');
                var stage = $('<input type="hidden" name="stage" id="stage" value="' + $.getParamValue('stage') + '" />');
                var year = $('<input type="hidden" name="year" id="year" value="' + $.getParamValue('year') + '" />');
                var month = $('<input type="hidden" name="month" id="month" value="' + $.getParamValue('month') + '" />');
                var savetype = $('<input type="hidden" name="savetype" id="savetype" value="' + strType + '" />');

                var form = $("form")[0];

                $("#postiframe").remove();
                $("input[name='mid']").remove();
                $("input[name='stage']").remove();
                $("input[name='year']").remove();
                $("input[name='month']").remove();
                $("input[name='savetype']").remove();

                form.appendChild(iframe[0]);
                form.appendChild(mid[0]);
                form.appendChild(stage[0]);
                form.appendChild(year[0]);
                form.appendChild(month[0]);
                form.appendChild(savetype[0]);

                form.setAttribute("action", "../handler/saveReportMonthEx.ashx");
                form.setAttribute("method", "post");
                form.setAttribute("enctype", "multipart/form-data");
                form.setAttribute("encoding", "multipart/form-data");
                form.setAttribute("target", "postiframe");
                form.submit();
        }

        //儲存form post後 回傳值
        function feedback(str,strtype) {
            var form = document.body.getElementsByTagName('form')[0];
            form.target = '';
            form.method = "post";
            form.enctype = "application/x-www-form-urlencoded";
            form.encoding = "application/x-www-form-urlencoded";
            form.action = location;

            if (str.indexOf("Error") > -1)
                alert(str);

            if (str == "succeed") {
                if (strtype == "save") {
                    alert("儲存完成");
                    load_data();
                }
                //getBasicWork();
                //getPlace();
                //getSmart();
            }
        }

        //計算合計
        function sumVal(thisid) {
            var str1 = 0
            var str2 = 0
            var str3 = 0
            var strType1 = thisid.substr(0, 2), strType2 = "";
            var strSum;
            if (strType1 == "99") {
                strType1 = thisid.substr(0, 6);
                strType2 = thisid.substr(13, 1);
            } else {
                strType2 = thisid.substr(9, 1);
            }

            //如果抓到的欄位是空就補0
            if ($("#" + thisid + "").attr("t") == "strint" && $("#" + thisid + "").val() == "") {
                $("#" + thisid + "").val("0");
            }
            if ($("#" + thisid + "").attr("t") == "strflot" && $("#" + thisid + "").val() == "") {
                $("#" + thisid + "").val("0.0");
            }

            //抓到合計欄位
            if($("#" + strType1 + "RM_Type" + strType2 + "Value1").val() != "")
                str1 = $("#" + strType1 + "RM_Type" + strType2 + "Value1").val();
            if ($("#" + strType1 + "RM_Type" + strType2 + "Value2").val() != "")
                str2 = $("#" + strType1 + "RM_Type" + strType2 + "Value2").val();
            if ($("#" + strType1 + "RM_Type" + strType2 + "Value3").val() != "")
                str3 = $("#" + strType1 + "RM_Type" + strType2 + "Value3").val();

            //加總完看是整數或者是小數
            if ($("#" + thisid + "").attr("t") == "strint") {
                strSum = parseInt(str1) + parseInt(str2) + parseInt(str3);
                $("#" + strType1 + "RM_Type" + strType2 + "ValueSum").val(strSum);
                $("#" + strType1 + "RM_Type" + strType2 + "ValueSum_span").html(strSum);
            }
            if ($("#" + thisid + "").attr("t") == "strflot") {
                strSum = (parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1);//小數第一位
                $("#" + strType1 + "RM_Type" + strType2 + "ValueSum").val(strSum);
                $("#" + strType1 + "RM_Type" + strType2 + "ValueSum_span").html(strSum);
            }

            //套公式 算出節電量
            //申請:RM_PreVal_span 核定:RM_ChkVal_span 未核定:RM_NotChkVal_span
            var PreVal, ChkVal, FormulaVal;
            FormulaVal = ($("#" + strType1 + "RM_Formula").val() == "" ? "0" : $("#" + strType1 + "RM_Formula").val());//年節電量度
            //本月申請數預期年節電量(度)
            if (strType2 == "1" && strType1 != "01" && strType1.substr(0, 2) != "99") {
                PreVal = conFun(strSum, parseFloat(FormulaVal));
                $("#" + strType1 + "RM_PreVal").val(PreVal);
                $("#" + strType1 + "RM_PreVal_span").html(PreVal);
                $("#" + strType1 + "RM_NotChkVal").val(parseFloat(parseFloat(PreVal) - parseFloat($("#" + strType1 + "RM_ChkVal").val())).toFixed(1));
                $("#" + strType1 + "RM_NotChkVal_span").html(parseFloat(parseFloat(PreVal) - parseFloat($("#" + strType1 + "RM_ChkVal").val())).toFixed(1));
            }
            //本月核定數預期年節電量(度)
            if (strType2 == "2" && strType1 != "01" && strType1.substr(0, 2) != "99") {
                ChkVal = conFun(strSum, parseFloat(FormulaVal));
                $("#" + strType1 + "RM_ChkVal").val(ChkVal);
                $("#" + strType1 + "RM_ChkVal_span").html(ChkVal);
                $("#" + strType1 + "RM_NotChkVal").val(parseFloat(parseFloat($("#" + strType1 + "RM_PreVal").val()) - parseFloat(ChkVal)).toFixed(1));
                $("#" + strType1 + "RM_NotChkVal_span").html(parseFloat(parseFloat($("#" + strType1 + "RM_PreVal").val()) - parseFloat(ChkVal)).toFixed(1));
            }
            //只有風管的 申請數&核定數是用W去算
            if (strType2 == "3" && strType1 == "01") {
                PreVal = conFun(strSum, parseFloat(FormulaVal));
                $("#" + strType1 + "RM_PreVal").val(PreVal);
                $("#" + strType1 + "RM_PreVal_span").html(PreVal);
                $("#" + strType1 + "RM_NotChkVal").val(parseFloat(parseFloat(PreVal) - parseFloat($("#" + strType1 + "RM_ChkVal").val())).toFixed(1));
                $("#" + strType1 + "RM_NotChkVal_span").html(parseFloat(parseFloat(parseFloat(PreVal) - parseFloat($("#" + strType1 + "RM_ChkVal").val()))).toFixed(1));
            }
            if (strType2 == "4" && strType1 == "01") {
                ChkVal = conFun(strSum, parseFloat(FormulaVal));
                $("#" + strType1 + "RM_ChkVal").val(ChkVal);
                $("#" + strType1 + "RM_ChkVal_span").html(ChkVal);
                $("#" + strType1 + "RM_NotChkVal").val(parseFloat($("#" + strType1 + "RM_PreVal").val()) - parseFloat(parseFloat(ChkVal)).toFixed(1));
                $("#" + strType1 + "RM_NotChkVal_span").html(parseFloat(parseFloat($("#" + strType1 + "RM_PreVal").val()) - parseFloat(parseFloat(ChkVal))).toFixed(1));
            }
        }

        //計算合計 (年節電量(度)欄位)
        function sumValF(thisid) {
            //thisid + RM_Formula
            var str1 = 0;
            var str2 = 0;
            var strType1 = thisid.substr(0, 2), strType2 = "", strType3 = "",stridNo;
            var strSum;
            if (strType1 == "99") {
                strType2 = thisid.substr(2, 2);
                strType3 = thisid.substr(4, 2);
                stridNo = thisid.substr(0, 6);
            } else {
                stridNo = thisid.substr(0, 2);
            }

            //如果抓到的欄位是空就補0
            if ($("#" + thisid + "").attr("t") == "strint" && $("#" + thisid + "").val() == "") {
                $("#" + thisid + "").val("0");
            }
            if ($("#" + thisid + "").attr("t") == "strflot" && $("#" + thisid + "").val() == "") {
                $("#" + thisid + "").val("0.0");
            }

            //抓到三個要加總的欄位
            if (strType1 == "01" || (strType2 == "02" && strType3 == "01")) {
                str1 = parseFloat($("#" + stridNo + "RM_Type3ValueSum").val());
                str2 = parseFloat($("#" + stridNo + "RM_Type4ValueSum").val());
                
            }
            else {
                str1 = parseFloat($("#" + stridNo + "RM_Type1ValueSum").val());
                str2 = parseFloat($("#" + stridNo + "RM_Type2ValueSum").val());
            }
            strSum = parseFloat($("#" + thisid + "").val());

            PreVal = conFun(strSum, parseFloat(str1));
            ChkVal = conFun(strSum, parseFloat(str2));

            $("#" + stridNo + "RM_PreVal").val(PreVal);
            $("#" + stridNo + "RM_PreVal_span").html(PreVal);
            $("#" + stridNo + "RM_ChkVal").val(ChkVal);
            $("#" + stridNo + "RM_ChkVal_span").html(ChkVal);
            $("#" + stridNo + "RM_NotChkVal").val((PreVal - ChkVal).toFixed(1));
            $("#" + stridNo + "RM_NotChkVal_span").html((PreVal - ChkVal).toFixed(1));
        }

        //套公式
        function conFun(strSum, strFVal) {
            var returnStr;
            returnStr = parseFloat(strFVal * strSum).toFixed(1)
            //switch (str1) {
            //    case "01":
            //        //(接)風管空氣調節機(4KW/台)
            //        returnStr = parseFloat(1245 / 4 * (strSum)).toFixed(1);
            //        break;
            //    case "02":
            //        //老舊辦公室照明燈具(T5螢光燈具)
            //        returnStr = parseFloat(189 * (strSum)).toFixed(1);
            //        break;
            //    case "03":
            //        //室內停車照明
            //        returnStr = parseFloat(175.2 * (strSum)).toFixed(1);
            //        break;
            //    case "04":
            //        //51KW以下能管系統
            //        returnStr = parseFloat(5280.0 * (strSum)).toFixed(1);
            //        break;
            //    case "05":
            //        //家庭(冷氣670度/3.2KW)
            //        returnStr = parseFloat(670 * (strSum)).toFixed(1);
            //        break;
            //    case "06":
            //        //電冰箱
            //        returnStr = parseFloat(526 * (strSum)).toFixed(1);
            //        break;
            //    case "07":
            //        //電視
            //        returnStr = parseFloat(62 * (strSum)).toFixed(1);
            //        break;
            //    case "08":
            //        //電熱水瓶
            //        returnStr = parseFloat(251 * (strSum)).toFixed(1);
            //        break;
            //    case "09":
            //        //電熱水器(儲備型)
            //        returnStr = parseFloat(165 * (strSum)).toFixed(1);
            //        break;
            //    case "10":
            //        //電子鍋
            //        returnStr = parseFloat(172 * (strSum)).toFixed(1);
            //        break;
            //    case "11":
            //        //溫熱型開飲機
            //        returnStr = parseFloat(234 * (strSum)).toFixed(1);
            //        break;
            //    case "12":
            //        //電鍋
            //        returnStr = parseFloat(172 * (strSum)).toFixed(1);
            //        break;
            //    case "13":
            //        //吹風機
            //        returnStr = parseFloat(21 * (strSum)).toFixed(1);
            //        break;
            //    case "14":
            //        //公設LED照明(10顆)
            //        returnStr = parseFloat(200 * (strSum)).toFixed(1);
            //        break;
            //    default:
            //        returnStr = 0;
            //        break
            //}
            return returnStr;
        }

        //格式化日期
        function formatDate(date) {
            var d = new Date(date),
                month = '' + (d.getMonth() + 1),
                day = '' + d.getDate(),
                year = d.getFullYear();

            if (month.length < 2) month = '0' + month;
            if (day.length < 2) day = '0' + day;

            return [year, month, day].join('/');
        }
        //送審
        function goCheck() {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/ReportMonth.ashx",
                data: {
                    func: "goCheck",
                    //str_stage: $("#txt_stage").val(),
                    //str_year: $("#txt_year").val(),
                    //str_month: $("#txt_month").val()
                    str_stage: $.getParamValue('stage'),
                    str_year: $.getParamValue('year'),
                    str_month: $.getParamValue('month'),
                    str_rcreporttype: "03"
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if (data == "timeout") {
                        alert("請重新登入");
                        location.href = "Login.aspx";
                    } else if (data.indexOf("Error") > -1) {
                        alert(data);
                    } else if (data.indexOf("nodata") > -1) {
                        alert("請先儲存");
                    } else {
                        if (data == "success") {
                            alert("成功將月報送審");
                            disabledInputandBtn();
                            location.href = "ReportMonthList.aspx";
                        }
                    }
                }
            });//ajax end
        }

        //disabled所有input跟button
        function disabledInputandBtn() {
            $("input").attr("disabled", "disabled"); 
            $("textarea").attr("disabled", "disabled");
            $(".genbtn").attr("disabled", "disabled");
            $(".genbtn").hide();
        }
        //還原button
        function showbutton() {
            $(".genbtn").removeAttr("disabled");
            $(".genbtn").show();
        }
        //數字欄位 如果為空就給0 或 0.0
        function changeZero() {
            $("input").each(function () {
                if ($(this).attr("t") == "strint" && ($(this).val() == "" || $(this).val() == "0" || $(this).val()=="0.0")) {
                    $(this).val("0");
                }
                if ($(this).attr("t") == "strflot" && ($(this).val() == "" || $(this).val() == "0" || $(this).val()=="0.0")) {
                    $(this).val("0.0");
                }
            });

        }

        //撈出當月第一天跟最後一天
        function monthFirstLastDate(strY,strM) {
            var y = strY, m = strM;
            var firstDay = new Date(y, m - 1, 1);
            var lastDay = new Date(y, m, 0);
            //var fromatDates = new Date(firstDay);
            //var fromatDatee = new Date(lastDay);
            $("#div_nowMonthSchedule").empty().append("本月執行進度(" + formatDate(firstDay) + "~" + formatDate(lastDay) + ")");
        }
        //自動存檔 feedback
        function autofeedback(str) {
            var form = document.body.getElementsByTagName('form')[0];
            form.target = '';
            form.method = "post";
            form.enctype = "application/x-www-form-urlencoded";
            form.encoding = "application/x-www-form-urlencoded";
            form.action = location;
            if (str == "succeed") {
                load_data();
            }

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
                <div class="OchiHalf">
                    <div class="OchiCell OchiTitle TitleSetWidth">期數</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <!--<div class="OchiCellInner width100" id="txt_stage"></div>-->
                            <div class="OchiCellInner nowrap textcenter" style="text-align: left;" id="txt_datesanddatee"></div>
                        </div><!-- OchiTableInner -->
                        <!-- cell內容end -->
                    </div><!-- OchiCell -->
                </div>
                <div class="OchiHalf">
                    <div class="OchiCell OchiTitle TitleSetWidth">月報類別</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <div class="OchiCellInner width100">擴大補助</div>
                        </div><!-- OchiTableInner -->
                        <!-- cell內容end -->
                    </div><!-- OchiCell -->
                </div>
            </div><!-- OchiRow -->

            <!-- 單欄 填報月份 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">提報月份</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100">
                        <div class="OchiCellInner width100" id="txt_yearmonth"></div>
                    </div>
                    <!-- OchiTableInner -->
                    <!-- cell內容end -->
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- OchiTrasTable -->
            <div class="font-bold font-title margin10T" id="div_nowMonthSchedule" style="font-size: 16px;"></div>
            <div id="div_data"></div>


        </div>
        

        <!-- OchiTrasTable 儲存&取消按鈕 -->
        <div class="twocol margin15T margin5B">
            <div class="right">
                <a href="javascript:void(0);" class="genbtn pbtn" id="btn_goCheck" style="color: red;">送審</a>
                <a href="javascript:void(0);" class="genbtn pbtn" id="btn_save">儲存</a>
                <a href="javascript:void(0);" class="genbtn pbtn" id="btn_cancel">取消</a>
            </div>
        </div>
        <!-- twocol -->

        <input type="hidden" id="hidden_mid" />
        <input type="hidden" id="hidden_mcity" />
        <input type="hidden" id="hidden_chktype" />
        <input type="hidden" id="hidden_chkstatus" />
    </div>
    <!-- conainer -->
    <script type="text/javascript">

    </script>
</asp:Content>

