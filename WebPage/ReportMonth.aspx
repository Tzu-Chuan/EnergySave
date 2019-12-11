<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="ReportMonth.aspx.cs" Inherits="WebPage_ReportMonth" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script>
        setInterval(function () {
            var breakStatus = false;
            if ($.getParamValue('stage') == "" || $.getParamValue('year') == null || $.getParamValue('year') == "" || $.getParamValue('month') == "") {
                breakStatus = true;
            }
            if ($("#hidden_chktype").val() == "Y" || ($("#hidden_chktype").val() == "" && $("#hidden_chkstatus").val() == "A")) {
                breakStatus = true;
            }
            if (breakStatus == true)
                return;
            //判斷如果現在是定稿或送審
            if ($("#btn_save").prop("disabled") || $("#txt_Type01Real").prop("disabled"))
                return;
            var iframe = $('<iframe name="postiframe" id="postiframe" style="display: none" />');
            var mid = $('<input type="hidden" name="mid" id="mid" value="' + $.getParamValue('v') + '" />');
            var stage = $('<input type="hidden" name="stage" id="stage" value="' + $.getParamValue('stage') + '" />');
            var year = $('<input type="hidden" name="year" id="year" value="' + $.getParamValue('year') + '" />');
            var month = $('<input type="hidden" name="month" id="month" value="' + $.getParamValue('month') + '" />');
            var rmtype = $('<input type="hidden" name="rmtype" id="rmtype" value="' + $.getParamValue('rmtype') + '" />');
            var form = $("form")[0];
            $("#postiframe").remove();
            $("input[name='mid']").remove();
            $("input[name='stage']").remove();
            $("input[name='year']").remove();
            $("input[name='month']").remove();
            $("input[name='rmtype']").remove();
            form.appendChild(iframe[0]);
            form.appendChild(mid[0]);
            form.appendChild(stage[0]);
            form.appendChild(year[0]);
            form.appendChild(month[0]);
            form.setAttribute("action", "../handler/AutoSave_ReportMonth.ashx");
            form.setAttribute("method", "post");
            form.setAttribute("enctype", "multipart/form-data");
            form.setAttribute("encoding", "multipart/form-data");
            form.setAttribute("target", "postiframe");
            form.submit();
        }, 1200000);//20 minutes 1200000

        $(function () {
            //$("#txt_stage").html("第" + $.getParamValue('stage') + "期");
            $("#txt_yearmonth").html($.getParamValue('year') + " 年" + $.getParamValue('month') + "月");

            $(".pbtn").hide();
            $("#hidden_mid").val($.getParamValue('v'));
            $("#div_data").empty();
            //期數change
            load_stagedata($.getParamValue('stage'));
            load_data();
            //$("#txt_stage").change(function () {
            //    if ($(this).val() != "") {
            //        load_stagedata($(this).val());
            //        //load_year($(this).val());
            //    }
            //});

            //期數+年分+月份 change 然後去撈閱報資料
            //$("#txt_stage,#txt_year,#txt_month").change(function () {
            //    chk_selval();
            //});
            //load_year();

            load_citydatabymid();

            //儲存
            $(document).on("click", "#btn_save", function () {
                //var errMsg = "";
                //if ($("#txt_stage").val() == "") {
                //    errMsg += "[請選擇期數]\n";
                //}
                //if ($("#txt_year").val() == null || $("#txt_year").val() == "") {
                //    errMsg += "[請選擇年份]\n";
                //}
                //if ($("#txt_month").val() == "") {
                //    errMsg += "[請選擇月份]\n";
                //}
                //if (errMsg != "") {
                //    alert(errMsg);
                //    return;
                //}
                if ($.getParamValue('stage') == "" || $.getParamValue('year') == "" || $.getParamValue('month') == "") {
                    alert("參數錯誤，請重新操作");
                    location.href = "ReportMonthList.aspx";
                    return;
                }
                var iframe = $('<iframe name="postiframe" id="postiframe" style="display: none" />');
                var mid = $('<input type="hidden" name="mid" id="mid" value="' + $.getParamValue('v') + '" />');
                //var stage = $('<input type="hidden" name="stage" id="stage" value="' + $("#txt_stage").val() + '" />');
                //var year = $('<input type="hidden" name="year" id="year" value="' + $("#txt_year").val() + '" />');
                //var month = $('<input type="hidden" name="month" id="month" value="' + $("#txt_month").val() + '" />');
                var stage = $('<input type="hidden" name="stage" id="stage" value="' + $.getParamValue('stage') + '" />');
                var year = $('<input type="hidden" name="year" id="year" value="' + $.getParamValue('year') + '" />');
                var month = $('<input type="hidden" name="month" id="month" value="' + $.getParamValue('month') + '" />');

                var form = $("form")[0];

                $("#postiframe").remove();
                $("input[name='mid']").remove();
                $("input[name='stage']").remove();
                $("input[name='year']").remove();
                $("input[name='month']").remove();

                form.appendChild(iframe[0]);
                form.appendChild(mid[0]);
                form.appendChild(stage[0]);
                form.appendChild(year[0]);
                form.appendChild(month[0]);

                form.setAttribute("action", "../handler/saveReportMonth.ashx");
                form.setAttribute("method", "post");
                form.setAttribute("enctype", "multipart/form-data");
                form.setAttribute("encoding", "multipart/form-data");
                form.setAttribute("target", "postiframe");
                form.submit();
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
            //數字欄位 blur 計算合計用
            $(document).on("blur", "input", function () {
                sumVal($(this).attr("id"));
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
                                //$("#div_person").empty().append($(this).children("M_Name").text().trim());
                                //$("#div_phone").empty().append($(this).children("M_Phone").text().trim());
                                //$("#div_createdate").empty().append("");
                                //$("#div_dossname").empty().append($(this).children("Manager_name").text().trim());
                                //$("#div_chkdate").empty().append("");
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
        //檢查期程 年 月是不是都有選
        //function chk_selval() {
        //    var sel_stage = $("#txt_stage").val();
        //    var sel_year = $("#txt_year").val();
        //    var sel_month = $("#txt_month").val();
        //    $("#div_data").empty();
        //    $("#div_nowMonthSchedule").empty();
        //    disabledInputandBtn();
        //    if (sel_stage != "" && sel_year != "" && sel_month != "") {
        //        monthFirstLastDate(sel_year, sel_month);
        //        load_data();
        //    }

        //}

        //撈月報資料
        function load_data() {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/ReportMonth.ashx",
                data: {
                    func: "load_MonthReportData",//txt_stage  
                    //str_guid: $.getParamValue('v'),
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
                        data = $.parseXML(data);
                        if (data == null) {
                            //沒資料連button都不顯示
                            disabledInputandBtn();
                            $("#div_data").empty().append("<font size='5' color='red'>尚無查核點及預訂工作進度資料，請至計畫基本資料確認</font>");
                            return;
                        }
                        if ($(data).find("data_item").length > 0) {
                            //showInputandBtn();
                            //$("#div_data").empty();
                            var strCreateDate;
                            var strChkDate;
                            $(data).find("data_item").each(function (i) {
                                var strRow1Th1, strRow1Th2, strRow1Th3, strRow1Th4;
                                var strRow2Th1, strRow2Th2, strRow2Th3;
                                var strunit, strunitOne;
                                var strRM_Finish,strRM_Finish01;
                                var strRM_Planning = $(this).children("RM_Planning").text().trim();
                                var strRM_Type3Value1 = $(this).children("RM_Type3Value1").text().trim();
                                var strRM_Type3Value2 = $(this).children("RM_Type3Value2").text().trim();
                                var strRM_Type3Value3 = $(this).children("RM_Type3Value3").text().trim();
                                var strRM_Type3ValueSum = $(this).children("RM_Type3ValueSum").text().trim();

                                var strcountApplyKW = $(this).children("countApplyKW").text().trim();//累計申請數KW
                                var strcountApply01 = $(this).children("countApply01").text().trim();//累計申請數

                                var splitRM_Finish, splitRM_Finish01, splitRM_Planning, splitRM_Type3Value1, splitRM_Type3Value2, splitRM_Type3Value3, splitRM_Type3ValueSum;
                                var planValue01, planValue02, planValue03, planValue04, planValue05;
                                var ptno = $(this).children("P_ItemName").text().trim();
                                var strIntFloat;
                                var strday = new Date();
                                //99是"其他"，這是不會出現在月報上面的，但是join出來的資料會有
                                //如果不過濾掉，chk_status會變成空直，沒辦法判斷現在到底是不是送審中
                                if ($(this).children("P_ItemName").text().trim() != "99") {
                                    chk_status = $(this).children("RC_Status").text().trim();
                                    chk_type = $(this).children("RC_CheckType").text().trim();
                                }
                                //填表日 如果還沒審核過 填表日永遠是填表當下的日期
                                if ($(this).children("RC_CheckDate").text().trim() != "" && $(this).children("RC_CheckDate").text().trim() != "1900-01-01T00:00:00+08:00") {
                                    chk_date = formatDate($(this).children("RC_CheckDate").text().trim());
                                    strCreateDate = formatDate($(this).children("RM_ModDate").text().trim());
                                } else {
                                    chk_date = "";
                                    strCreateDate = strday.getFullYear() + "/" + (strday.getMonth() + 1) + "/" + strday.getDate();
                                }
                                if ($(this).children("P_ItemName").text().trim() != '99') {
                                    if ($(this).children("P_ItemName").text().trim() == '01' || $(this).children("P_ItemName").text().trim() == '02' || $(this).children("P_ItemName").text().trim() == '03') {
                                        if ($(this).children("P_ItemName").text().trim() == '01') {
                                            strRow1Th1 = "本月申請數量(台)";
                                            strRow1Th2 = "本月核定數量(台)";
                                            strRow1Th3 = "本月申請總冷氣能力(kW) ";
                                            strRow1Th4 = "本月完成總冷氣能力(kW)";
                                            strRow2Th1 = "機關";
                                            strRow2Th2 = "學校";
                                            strRow2Th3 = "服務業";
                                            strunit = "(kW)";
                                            strunitOne = "(台)";
                                            strIntFloat = "strflot";
                                            splitRM_Planning = $(this).children("RM_Planning").text().trim().split('.');
                                            //已審核或審核中的月報 累計數撈資料表存的
                                            //未審核的月報 累計數用計算的
                                            //if (chk_type == "Y" || (chk_type == "" && chk_status == "A")) {
                                            //    splitRM_Finish = $(this).children("countFinishKW").text().trim().split('.');
                                            //    splitRM_Finish01 = $(this).children("countFinish02").text().trim().split('.');
                                            //} else {
                                            //    splitRM_Finish = $(this).children("countFinishKW").text().trim().split('.');
                                            //    splitRM_Finish01 = $(this).children("countFinish02").text().trim().split('.');
                                            //}
                                            //累計的全部用算的出來，而且是不包含本月份的，且只有審核過的月報才列入累計
                                            //splitRM_Finish = $(this).children("countFinishKW").text().trim().split('.');
                                            splitRM_Finish01 = $(this).children("countFinish02").text().trim().split('.');
                                            strRM_Planning = splitRM_Planning[0];
                                            strRM_Finish = $(this).children("countFinishKW").text().trim();
                                            strRM_Finish01 = splitRM_Finish01[0];
                                            //本期規劃數  && (strRM_Planning == "" || strRM_Planning == null)
                                            if ($.getParamValue('stage') == "1")
                                                strRM_Planning = $(this).children("I_Finish_item1_1").text().trim();
                                            if ($.getParamValue('stage') == "2")
                                                strRM_Planning = $(this).children("I_Finish_item1_2").text().trim();
                                            if ($.getParamValue('stage') == "3")
                                                strRM_Planning = $(this).children("I_Finish_item1_3").text().trim();

                                        }
                                        if ($(this).children("P_ItemName").text().trim() == '02') {
                                            strRow1Th1 = "本月申請數量(具)";
                                            strRow1Th2 = "本月核定數量(具)";
                                            strRow1Th3 = "本月申請更換照明瓦數(W)";
                                            strRow1Th4 = "本月完成更換照明瓦數(W)";
                                            strRow2Th1 = "機關";
                                            strRow2Th2 = "學校";
                                            strRow2Th3 = "服務業";
                                            strunit = "(具)";
                                            strunitOne = "(具)";
                                            splitRM_Planning = $(this).children("RM_Planning").text().trim().split('.');
                                            //已審核或審核中的月報 累計數撈資料表存的
                                            //未審核的月報 累計數用計算的
                                            if (chk_type == "Y" || (chk_type == "" && chk_status == "A")) {
                                                splitRM_Finish = $(this).children("RM_Finish").text().trim().split('.');
                                            } else {
                                                splitRM_Finish = $(this).children("countFinish02").text().trim().split('.');
                                            }
                                            strRM_Planning = splitRM_Planning[0];
                                            strRM_Finish = splitRM_Finish[0];
                                            strIntFloat = "strint";
                                            //本期規劃數
                                            if ($.getParamValue('stage') == "1")
                                                strRM_Planning = $(this).children("I_Finish_item2_1").text().trim();
                                            if ($.getParamValue('stage') == "2")
                                                strRM_Planning = $(this).children("I_Finish_item2_2").text().trim();
                                            if ($.getParamValue('stage') == "3")
                                                strRM_Planning = $(this).children("I_Finish_item2_3").text().trim();
                                        }
                                        if ($(this).children("P_ItemName").text().trim() == '03') {
                                            strRow1Th1 = "本月申請數量(盞)";
                                            strRow1Th2 = "本月核定數量(盞)";
                                            strRow1Th3 = "本月申請更換照明瓦數(W)";
                                            strRow1Th4 = "本月完成更換照明瓦數(W)";
                                            strRow2Th1 = "集合住宅  ";
                                            strRow2Th2 = "辦公大樓";
                                            strRow2Th3 = "服務業";
                                            strunit = "(盞)";
                                            strunitOne = "(盞)";
                                            splitRM_Planning = $(this).children("RM_Planning").text().trim().split('.');
                                            //已審核或審核中的月報 累計數撈資料表存的
                                            //未審核的月報 累計數用計算的
                                            if (chk_type == "Y" || (chk_type == "" && chk_status == "A")) {
                                                splitRM_Finish = $(this).children("RM_Finish").text().trim().split('.');
                                            } else {
                                                splitRM_Finish = $(this).children("countFinish02").text().trim().split('.');
                                            }
                                            strRM_Planning = splitRM_Planning[0];
                                            strRM_Finish = splitRM_Finish[0];
                                            strIntFloat = "strint";
                                            //本期規劃數
                                            if ($.getParamValue('stage') == "1")
                                                strRM_Planning = $(this).children("I_Finish_item3_1").text().trim();
                                            if ($.getParamValue('stage') == "2")
                                                strRM_Planning = $(this).children("I_Finish_item3_2").text().trim();
                                            if ($.getParamValue('stage') == "3")
                                                strRM_Planning = $(this).children("I_Finish_item3_3").text().trim();
                                        }
                                        strHtml += '<div class="OchiRow">';
                                        strHtml += '<div class="OchiCell OchiTitle TitleSetWidth">' + $(this).children("C_Item_cn").text().trim() + '</div>';
                                        strHtml += '<div class="OchiCell width100">';
                                        strHtml += '<div class="OchiTableInner width100">';
                                        if ($(this).children("P_ItemName").text().trim() == '01') {
                                            
                                            strHtml += '<div class="OchiCellInner nowrap textcenter">本期累計申請數:</div>';
                                            strHtml += '<div class="OchiCellInner width50"><input type="text" class="inputex" size="18" name="' + ptno + 'RM_ApplyKW" value="' + strcountApplyKW + '" t="' + strIntFloat + '" readonly="readonly" style="background-color:#DDDDDD;" />&nbsp;' + strunit + '</div>';

                                        } else {
                                            strHtml += '<div class="OchiCellInner nowrap textcenter">本期累計核定數:</div>';
                                            strHtml += '<div class="OchiCellInner width33"><input type="text" class="inputex" size="12" name="' + ptno + 'RM_Finish" value="' + strRM_Finish + '" t="' + strIntFloat + '" readonly="readonly" style="background-color:#DDDDDD;" />&nbsp;' + strunit + '</div>';
                                            strHtml += '<div class="OchiCellInner nowrap textcenter">本期累計申請數:</div>';
                                            strHtml += '<div class="OchiCellInner width33"><input type="text" class="inputex" size="12" name="' + ptno + 'RM_Apply01" value="' + strcountApply01 + '" t="' + strIntFloat + '" readonly="readonly" style="background-color:#DDDDDD;" />&nbsp;' + strunit + '</div>';
                                        }
                                        strHtml += '<div class="OchiCellInner nowrap textcenter">本期規劃數:</div>';
                                        strHtml += '<div class="OchiCellInner width50"><input type="text" class="inputex" size="18" name="' + ptno + 'RM_Planning" value="' + strRM_Planning + '" readonly="readonly" style="background-color:#DDDDDD;" />&nbsp;' + strunit + '</div>';
                                        //strHtml += '</div>';
                                        
                                        strHtml += '</div>';
                                        if ($(this).children("P_ItemName").text().trim() == '01') {
                                            //strHtml += '<div class="OchiCell width100">';
                                            strHtml += '<div class="OchiTableInner width100">';
                                            strHtml += '<div class="OchiCellInner nowrap textcenter">本期累計完成數:</div>';
                                            strHtml += '<div class="OchiCellInner width50"><input type="text" class="inputex" size="18" name="' + ptno + 'RM_Finish" value="' + strRM_Finish + '" t="' + strIntFloat + '" readonly="readonly" style="background-color:#DDDDDD;" />&nbsp;' + strunit + '</div>';
                                            strHtml += '<div class="OchiCellInner nowrap textcenter">本期累計核定數:</div>';
                                            strHtml += '<div class="OchiCellInner width50"><input type="text" class="inputex" size="12" name="' + ptno + 'RM_Finish01" value="' + strRM_Finish01 + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;" />&nbsp;' + strunitOne + '</div>';
                                            //strHtml += '<div class="OchiCellInner nowrap textcenter" style="text-align:left;">本期累計核定數:<input type="text" class="inputex" size="8" name="' + ptno + 'RM_Finish01" value="' + strRM_Finish01 + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;" />&nbsp;' + strunitOne + '</div>';
                                            strHtml += '</div>';
                                            //strHtml += '</div>';
                                        }

                                        strHtml += '<div class="OchiTableInner width100">';
                                        strHtml += '<div class="stripeMepure font-normal margin5T">';
                                        strHtml += '<table border="0" cellspacing="0" cellpadding="0" width="100%">';
                                        strHtml += '<tr>';
                                        strHtml += '<th colspan="3">' + strRow1Th1 + '</th>';
                                        strHtml += '<th colspan="4">' + strRow1Th2 + '</th>';
                                        strHtml += '<th colspan="4">' + strRow1Th3 + '</th>';
                                        strHtml += '<th colspan="3">' + strRow1Th4 + '</th>';
                                        strHtml += '</tr>';
                                        strHtml += '<tr>';
                                        strHtml += '<td align="center">' + strRow2Th1 + '</td>';
                                        strHtml += '<td align="center">' + strRow2Th2 + '</td>';
                                        strHtml += '<td align="center"><strong>' + strRow2Th3 + '</strong></td>';
                                        strHtml += '<td align="center">' + strRow2Th1 + '</td>';
                                        strHtml += '<td colspan="2" align="center">' + strRow2Th2 + '</td>';
                                        strHtml += '<td align="center"><strong>' + strRow2Th3 + '</strong></td>';
                                        strHtml += '<td align="center">' + strRow2Th1 + '</td>';
                                        strHtml += '<td colspan="2" align="center">' + strRow2Th2 + '</td>';
                                        strHtml += '<td align="center"><strong>' + strRow2Th3 + '</strong></td>';
                                        strHtml += '<td align="center">' + strRow2Th1 + '</td>';
                                        strHtml += '<td align="center">' + strRow2Th2 + '</td>';
                                        strHtml += '<td align="center"><strong>' + strRow2Th3 + '</strong></td>';
                                        strHtml += '</tr>';
                                        strHtml += '<tr>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1Value1" id="' + ptno + 'RM_Type1Value1" value="' + $(this).children("RM_Type1Value1").text().trim() + '" t="strint" /></td>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1Value2" id="' + ptno + 'RM_Type1Value2" value="' + $(this).children("RM_Type1Value2").text().trim() + '" t="strint" /></td>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1Value3" id="' + ptno + 'RM_Type1Value3" value="' + $(this).children("RM_Type1Value3").text().trim() + '" t="strint" /></td>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2Value1" id="' + ptno + 'RM_Type2Value1" value="' + $(this).children("RM_Type2Value1").text().trim() + '" t="strint" /></td>';
                                        strHtml += '<td colspan="2"><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2Value2" id="' + ptno + 'RM_Type2Value2" value="' + $(this).children("RM_Type2Value2").text().trim() + '" t="strint" /></td>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2Value3" id="' + ptno + 'RM_Type2Value3" value="' + $(this).children("RM_Type2Value3").text().trim() + '" t="strint" /></td>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3Value1" id="' + ptno + 'RM_Type3Value1" value="' + $(this).children("RM_Type3Value1").text().trim() + '" t="strflot" /></td>';
                                        strHtml += '<td colspan="2"><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3Value2" id="' + ptno + 'RM_Type3Value2" value="' + $(this).children("RM_Type3Value2").text().trim() + '" t="strflot" /></td>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3Value3" id="' + ptno + 'RM_Type3Value3" value="' + $(this).children("RM_Type3Value3").text().trim() + '" t="strflot" /></td>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type4Value1" id="' + ptno + 'RM_Type4Value1" value="' + $(this).children("RM_Type4Value1").text().trim() + '" t="strflot" /></td>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type4Value2" id="' + ptno + 'RM_Type4Value2" value="' + $(this).children("RM_Type4Value2").text().trim() + '" t="strflot" /></td>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type4Value3" id="' + ptno + 'RM_Type4Value3" value="' + $(this).children("RM_Type4Value3").text().trim() + '" t="strflot" /></td>';
                                        strHtml += '</tr>';
                                        strHtml += '<tr>';
                                        strHtml += '<td align="right">合計 </td>';
                                        strHtml += '<td colspan="2"><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1ValueSum" id="' + ptno + 'RM_Type1ValueSum" value="' + $(this).children("RM_Type1ValueSum").text().trim() + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                                        strHtml += '<td align="right" colspan="2">合計 </td>';
                                        strHtml += '<td colspan="2"><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2ValueSum" id="' + ptno + 'RM_Type2ValueSum" value="' + $(this).children("RM_Type2ValueSum").text().trim() + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                                        strHtml += '<td align="right" colspan="2">合計 </td>';
                                        strHtml += '<td colspan="2"><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3ValueSum" id="' + ptno + 'RM_Type3ValueSum" value="' + $(this).children("RM_Type3ValueSum").text().trim() + '" t="strflot" readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                                        strHtml += '<td align="right">合計 </td>';
                                        strHtml += '<td colspan="2"><input type="text" class="inputex width100" name="' + ptno + 'RM_Type4ValueSum" id="' + ptno + 'RM_Type4ValueSum" value="' + $(this).children("RM_Type4ValueSum").text().trim() + '" t="strflot" readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                                        strHtml += '</tr>';
                                        strHtml += '<tr>';
                                        strHtml += '<th colspan="4">本月申請數預期年節電量(度)</th>';
                                        strHtml += '<th colspan="6">本月核定數預期年節電量(度)</th>';
                                        strHtml += '<th colspan="4">本月未核定數之年節電量(度) </th>';
                                        strHtml += '</tr>';
                                        strHtml += '<tr>';
                                        strHtml += '<td colspan="4"><input type="text" class="inputex width100" name="' + ptno + 'RM_PreVal" id="' + ptno + 'RM_PreVal" value="' + $(this).children("RM_PreVal").text().trim() + '" t="strflot" readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                                        strHtml += '<td colspan="6"><input type="text" class="inputex width100" name="' + ptno + 'RM_ChkVal" id="' + ptno + 'RM_ChkVal" value="' + $(this).children("RM_ChkVal").text().trim() + '" t="strflot" readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                                        strHtml += '<td colspan="4"><input type="text" class="inputex width100" name="' + ptno + 'RM_NotChkVal" id="' + ptno + 'RM_NotChkVal" value="' + $(this).children("RM_NotChkVal").text().trim() + '" t="strflot" readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                                        strHtml += '</tr>';
                                        strHtml += '</table>';
                                        strHtml += '</div>';
                                        strHtml += '</div>';
                                        var strRemark = "";
                                        if ($(this).children("P_ItemName").text().trim() == '02')//
                                        {
                                            strRemark = "(T8/T9)";
                                        }
                                        strHtml += '<div class="OchiTableInner width100 margin5T">補充說明' + strRemark + '：<textarea rows="3" class="inputex width100" name="' + ptno + 'RM_Remark">' + $(this).children("RM_Remark").text().trim() + '</textarea></div>';
                                        strHtml += '</div>';
                                        strHtml += '</div>';
                                        strHtml += '<div><input type="hidden" name="report_type" value="' + $(this).children("P_ItemName").text().trim() + '" /></div>';
                                        strHtml += '<div><input type="hidden" name="report_P_Guid" value="' + $(this).children("P_Guid").text().trim() + '" /></div>';
                                        strHtml += '<div><input type="hidden" name="report_Guid" value="' + $(this).children("RM_ReportGuid").text().trim() + '" /></div>';
                                        //$(this).children("P_ItemName").text().trim()
                                    } else {
                                        if ($(this).children("P_ItemName").text().trim() == '04') {
                                            //本期規劃數
                                            if ($.getParamValue('stage') == "1")
                                                strRM_Planning = $(this).children("I_Finish_item4_1").text().trim();
                                            if ($.getParamValue('stage') == "2" )
                                                strRM_Planning = $(this).children("I_Finish_item4_2").text().trim();
                                            if ($.getParamValue('stage') == "3")
                                                strRM_Planning = $(this).children("I_Finish_item4_3").text().trim();

                                        }
                                        if ($(this).children("P_ItemName").text().trim() == '05') {
                                            //本期規劃數
                                            if ($.getParamValue('stage') == "1")
                                                strRM_Planning = $(this).children("I_Finish_item5_1").text().trim();
                                            if ($.getParamValue('stage') == "2")
                                                strRM_Planning = $(this).children("I_Finish_item5_2").text().trim();
                                            if ($.getParamValue('stage') == "3")
                                                strRM_Planning = $(this).children("I_Finish_item5_3").text().trim();

                                        }
                                        splitRM_Planning = strRM_Planning.split('.');
                                        splitRM_Type3Value1 = $(this).children("RM_Type3Value1").text().trim().split('.');
                                        splitRM_Type3Value2 = $(this).children("RM_Type3Value2").text().trim().split('.');
                                        splitRM_Type3Value3 = $(this).children("RM_Type3Value3").text().trim().split('.');
                                        splitRM_Type3ValueSum = $(this).children("RM_Type3ValueSum").text().trim().split('.');
                                        //已審核或審核中的月報 累計數撈資料表存的
                                        //未審核的月報 累計數用計算的
                                        if (chk_type == "Y" || (chk_type == "" && chk_status == "A")) {
                                            splitRM_Finish = $(this).children("RM_Finish").text().trim().split('.');
                                        } else {
                                            splitRM_Finish = $(this).children("countFinish03").text().trim().split('.');
                                        }
                                        strRM_Planning = splitRM_Planning[0];
                                        strRM_Type3Value1 = splitRM_Type3Value1[0];
                                        strRM_Type3Value2 = splitRM_Type3Value2[0];
                                        strRM_Type3Value3 = splitRM_Type3Value3[0];
                                        strRM_Type3ValueSum = splitRM_Type3ValueSum[0];
                                        strRM_Finish = splitRM_Finish[0];
                                        strHtml += '<div class="OchiRow">';
                                        strHtml += '<div class="OchiCell OchiTitle TitleSetWidth">' + $(this).children("C_Item_cn").text().trim() + '</div>';
                                        strHtml += '<div class="OchiCell width100">';
                                        strHtml += '<div class="OchiTableInner width100">';
                                        strHtml += '<div class="OchiCellInner nowrap textcenter">本期累計完成數:</div>';
                                        strHtml += '<div class="OchiCellInner width33">';
                                        strHtml += '<input type="text" class="inputex" size="12" name="' + ptno + 'RM_Finish" value="' + strRM_Finish + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;" />&nbsp;(套)';
                                        strHtml += '</div>';
                                        strHtml += '<div class="OchiCellInner nowrap textcenter">本期累計申請數:</div>';
                                        strHtml += '<div class="OchiCellInner width33"><input type="text" class="inputex" size="12" name="' + ptno + 'RM_Finish" value="' + strcountApply01 + '" t="' + strIntFloat + '" readonly="readonly" style="background-color:#DDDDDD;" />&nbsp;(套)</div>';
                                        strHtml += '<div class="OchiCellInner nowrap textcenter">本期規劃數:</div>';
                                        strHtml += '<div class="OchiCellInner width33">';
                                        strHtml += '<input type="text" class="inputex" size="12" name="' + ptno + 'RM_Planning" value="' + strRM_Planning + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;" />&nbsp;(套)';
                                        strHtml += '</div>';
                                        strHtml += '</div>';
                                        strHtml += '<div class="OchiTableInner width100">';
                                        strHtml += '<div class="stripeMepure font-normal margin5T">';
                                        strHtml += '<table border="0" cellspacing="0" cellpadding="0" width="100%">';
                                        strHtml += '<tr>';
                                        strHtml += '<th colspan="4" valign="top">本月申請數量(套)</th>';
                                        strHtml += '<th colspan="4" valign="top">本月核定數量(套)</th>';
                                        strHtml += '<th colspan="4" valign="top">本月完成數量(套)</th>';
                                        strHtml += '</tr>';
                                        strHtml += '<tr>';
                                        strHtml += '<td align="center">機關 </td>';
                                        strHtml += '<td colspan="2" align="center">學校 </td>';
                                        strHtml += '<td align="center">服務業 </td>';
                                        strHtml += '<td align="center">機關 </td>';
                                        strHtml += '<td colspan="2" align="center">學校 </td>';
                                        strHtml += '<td align="center">服務業 </td>';
                                        strHtml += '<td align="center">機關 </td>';
                                        strHtml += '<td colspan="2" align="center">學校 </td>';
                                        strHtml += '<td align="center">服務業 </td>';
                                        strHtml += '</tr>';
                                        strHtml += '<tr>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1Value1" id="' + ptno + 'RM_Type1Value1" value="' + $(this).children("RM_Type1Value1").text().trim() + '" t="strint" /></td>';
                                        strHtml += '<td colspan="2"><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1Value2" id="' + ptno + 'RM_Type1Value2" value="' + $(this).children("RM_Type1Value2").text().trim() + '" t="strint" /></td>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1Value3" id="' + ptno + 'RM_Type1Value3" value="' + $(this).children("RM_Type1Value3").text().trim() + '" t="strint" /></td>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2Value1" id="' + ptno + 'RM_Type2Value1" value="' + $(this).children("RM_Type2Value1").text().trim() + '" t="strint" /></td>';
                                        strHtml += '<td colspan="2"><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2Value2" id="' + ptno + 'RM_Type2Value2" value="' + $(this).children("RM_Type2Value2").text().trim() + '" t="strint" /></td>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2Value3" id="' + ptno + 'RM_Type2Value3" value="' + $(this).children("RM_Type2Value3").text().trim() + '" t="strint" /></td>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3Value1" id="' + ptno + 'RM_Type3Value1" value="' + strRM_Type3Value1 + '" t="strint" /></td>';
                                        strHtml += '<td colspan="2"><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3Value2" id="' + ptno + 'RM_Type3Value2" value="' + strRM_Type3Value2 + '" t="strint" /></td>';
                                        strHtml += '<td><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3Value3" id="' + ptno + 'RM_Type3Value3" value="' + strRM_Type3Value3 + '" t="strint" /></td>';
                                        strHtml += '</tr>';
                                        strHtml += '<tr>';
                                        strHtml += '<td colspan="2" align="right">合計 </td>';
                                        strHtml += '<td colspan="2"><input type="text" class="inputex width100" name="' + ptno + 'RM_Type1ValueSum" id="' + ptno + 'RM_Type1ValueSum" value="' + $(this).children("RM_Type1ValueSum").text().trim() + '" t="strint"  readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                                        strHtml += '<td colspan="2" align="right">合計 </td>';
                                        strHtml += '<td colspan="2"><input type="text" class="inputex width100" name="' + ptno + 'RM_Type2ValueSum" id="' + ptno + 'RM_Type2ValueSum" value="' + $(this).children("RM_Type2ValueSum").text().trim() + '" t="strint"  readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                                        strHtml += '<td colspan="2" align="right">合計 </td>';
                                        strHtml += '<td colspan="2"><input type="text" class="inputex width100" name="' + ptno + 'RM_Type3ValueSum" id="' + ptno + 'RM_Type3ValueSum" value="' + strRM_Type3ValueSum + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                                        strHtml += '</tr>';
                                        strHtml += '<tr>';
                                        strHtml += '<th colspan="4">本月申請數預期年節電量(度)</th>';
                                        strHtml += '<th colspan="4">本月核定數預期年節電量(度)</th>';
                                        strHtml += '<th colspan="4">本月未核定數之年節電量(度) </th>';
                                        strHtml += '</tr>';
                                        strHtml += '<tr>';
                                        strHtml += '<td colspan="4"><input type="text" class="inputex width100" name="' + ptno + 'RM_PreVal" id="' + ptno + 'RM_PreVal" value="' + $(this).children("RM_PreVal").text().trim() + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                                        strHtml += '<td colspan="4"><input type="text" class="inputex width100" name="' + ptno + 'RM_ChkVal" id="' + ptno + 'RM_ChkVal" value="' + $(this).children("RM_ChkVal").text().trim() + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                                        strHtml += '<td colspan="4"><input type="text" class="inputex width100" name="' + ptno + 'RM_NotChkVal" id="' + ptno + 'RM_NotChkVal" value="' + $(this).children("RM_NotChkVal").text().trim() + '" t="strint" readonly="readonly" style="background-color:#DDDDDD;" /></td>';
                                        strHtml += '</tr>';
                                        strHtml += '</table>';
                                        strHtml += '</div>';
                                        strHtml += '</div>';
                                        strHtml += '<div class="OchiTableInner width100 margin5T">補充說明：<textarea rows="3" class="inputex width100" name="' + ptno + 'RM_Remark">' + $(this).children("RM_Remark").text().trim() + '</textarea></div>';
                                        strHtml += '</div>';
                                        strHtml += '</div>';
                                        strHtml += '<div><input type="hidden" name="report_type" value="' + $(this).children("P_ItemName").text().trim() + '" /></div>';
                                        strHtml += '<div><input type="hidden" name="report_P_Guid" value="' + $(this).children("P_Guid").text().trim() + '" /></div>';
                                        strHtml += '<div><input type="hidden" name="report_Guid" value="' + $(this).children("RM_ReportGuid").text().trim() + '" /></div>';
                                    }
                                }
                                
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

        //儲存form post後 回傳值
        function feedback(str) {
            var form = document.body.getElementsByTagName('form')[0];
            form.target = '';
            form.method = "post";
            form.enctype = "application/x-www-form-urlencoded";
            form.encoding = "application/x-www-form-urlencoded";
            form.action = location;

            if (str.indexOf("Error") > -1)
                alert(str);

            if (str == "succeed") {
                alert("儲存完成");
                load_data();
                //getBasicWork();
                //getPlace();
                //getSmart();
            }
        }

        //計算合計
        function sumVal(thisid) {
            //RM_Type1Value1 RM_Type1ValueSum
            var str1 = 0
            var str2 = 0
            var str3 = 0
            if ($("#" + thisid + "").attr("t") == "strint" && $("#" + thisid + "").val() == "") {
                $("#" + thisid + "").val("0");
            }
            if ($("#" + thisid + "").attr("t") == "strflot" && $("#" + thisid + "").val() == "") {
                $("#" + thisid + "").val("0.0");
            }
            switch (thisid) {
                case "01RM_Type1Value1":
                case "01RM_Type1Value2":
                case "01RM_Type1Value3":
                    if ($("#01RM_Type1Value1").val() != "")
                        str1 = $("#01RM_Type1Value1").val();
                    if ($("#01RM_Type1Value2").val() != "")
                        str2 = $("#01RM_Type1Value2").val();
                    if ($("#01RM_Type1Value3").val() != "")
                        str3 = $("#01RM_Type1Value3").val();
                    $("#01RM_Type1ValueSum").val(parseInt(str1) + parseInt(str2) + parseInt(str3));
                    break;
                case "01RM_Type2Value1":
                case "01RM_Type2Value2":
                case "01RM_Type2Value3":
                    if ($("#01RM_Type2Value1").val() != "")
                        str1 = parseInt($("#01RM_Type2Value1").val());
                    if ($("#01RM_Type2Value2").val() != "")
                        str2 = parseInt($("#01RM_Type2Value2").val());
                    if ($("#01RM_Type2Value3").val() != "")
                        str3 = parseInt($("#01RM_Type2Value3").val());
                    $("#01RM_Type2ValueSum").val(parseInt(str1) + parseInt(str2) + parseInt(str3));
                    break;
                case "01RM_Type3Value1":
                case "01RM_Type3Value2":
                case "01RM_Type3Value3":
                    if ($("#01RM_Type3Value1").val() != "")
                        str1 = $("#01RM_Type3Value1").val();
                    if ($("#01RM_Type3Value2").val() != "")
                        str2 = $("#01RM_Type3Value2").val();
                    if ($("#01RM_Type3Value3").val() != "")
                        str3 = $("#01RM_Type3Value3").val();
                    $("#01RM_Type3ValueSum").val(parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1));
                    //申請數預計年節電量 = 1245/4*(申請總冷氣能力(kW)(機關+學校+服務業))
                    //RM_PreVal RM_ChkVal RM_NotChkVal
                    var sum_01 = parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1);
                    var Val_01 = parseFloat(1245 / 4 * (sum_01)).toFixed(1);
                    $("#01RM_PreVal").val(Val_01);
                    //未核定數支年節電量(度) = 申請數預期年節電量(度) - 核定數預期年節電量(度)
                    $("#01RM_NotChkVal").val(parseFloat(parseFloat(Val_01) - parseFloat($("#01RM_ChkVal").val())).toFixed(1));
                    break;
                case "01RM_Type4Value1":
                case "01RM_Type4Value2":
                case "01RM_Type4Value3":
                    if ($("#01RM_Type4Value1").val() != "")
                        str1 = $("#01RM_Type4Value1").val();
                    if ($("#01RM_Type4Value2").val() != "")
                        str2 = $("#01RM_Type4Value2").val();
                    if ($("#01RM_Type4Value3").val() != "")
                        str3 = $("#01RM_Type4Value3").val();
                    $("#01RM_Type4ValueSum").val(parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1));
                    //核定數預計年節電量 = 1245/4*(完成總冷氣能力(kW)(機關+學校+服務業))
                    var sum_01 = parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1);
                    var Val_01 = parseFloat(1245 / 4 * (sum_01)).toFixed(1);
                    $("#01RM_ChkVal").val(Val_01);
                    //未核定數支年節電量(度) = 申請數預期年節電量(度) - 核定數預期年節電量(度)
                    $("#01RM_NotChkVal").val(parseFloat(parseFloat($("#01RM_PreVal").val()) - parseFloat(Val_01)).toFixed(1));
                    break;

                case "02RM_Type1Value1":
                case "02RM_Type1Value2":
                case "02RM_Type1Value3":
                    if ($("#02RM_Type1Value1").val() != "")
                        str1 = $("#02RM_Type1Value1").val();
                    if ($("#02RM_Type1Value2").val() != "")
                        str2 = $("#02RM_Type1Value2").val();
                    if ($("#02RM_Type1Value3").val() != "")
                        str3 = $("#02RM_Type1Value3").val();
                    $("#02RM_Type1ValueSum").val(parseInt(str1) + parseInt(str2) + parseInt(str3));
                    //申請數預期年節電量(度) = 189*(申請數量(具)(機關+學校+服務業))
                    var sum_02 = parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1);
                    var Val_02 = parseFloat(189 * (sum_02)).toFixed(1);
                    $("#02RM_PreVal").val(Val_02);
                    //未核定數支年節電量(度) = 申請數預期年節電量(度) - 核定數預期年節電量(度)
                    $("#02RM_NotChkVal").val(parseFloat(parseFloat(Val_02) - parseFloat($("#02RM_ChkVal").val())).toFixed(1));
                    break;
                case "02RM_Type2Value1":
                case "02RM_Type2Value2":
                case "02RM_Type2Value3":
                    if ($("#02RM_Type2Value1").val() != "")
                        str1 = $("#02RM_Type2Value1").val();
                    if ($("#02RM_Type2Value2").val() != "")
                        str2 = $("#02RM_Type2Value2").val();
                    if ($("#02RM_Type2Value3").val() != "")
                        str3 = $("#02RM_Type2Value3").val();
                    $("#02RM_Type2ValueSum").val(parseInt(str1) + parseInt(str2) + parseInt(str3));
                    //核定數預計年節電量 = 189*(申請數量(具)(機關+學校+服務業))
                    var sum_02 = parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1);
                    var Val_02 = parseFloat(189 * (sum_02)).toFixed(1);
                    $("#02RM_ChkVal").val(Val_02);
                    //未核定數支年節電量(度) = 申請數預期年節電量(度) - 核定數預期年節電量(度)
                    $("#02RM_NotChkVal").val(parseFloat(parseFloat($("#02RM_PreVal").val()) - parseFloat(Val_02)).toFixed(1));
                    break;
                case "02RM_Type3Value1":
                case "02RM_Type3Value2":
                case "02RM_Type3Value3":
                    if ($("#02RM_Type3Value1").val() != "")
                        str1 = $("#02RM_Type3Value1").val();
                    if ($("#02RM_Type3Value2").val() != "")
                        str2 = $("#02RM_Type3Value2").val();
                    if ($("#02RM_Type3Value3").val() != "")
                        str3 = $("#02RM_Type3Value3").val();
                    $("#02RM_Type3ValueSum").val(parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1));
                    break;
                case "02RM_Type4Value1":
                case "02RM_Type4Value2":
                case "02RM_Type4Value3":
                    if ($("#02RM_Type4Value1").val() != "")
                        str1 = $("#02RM_Type4Value1").val();
                    if ($("#02RM_Type4Value2").val() != "")
                        str2 = $("#02RM_Type4Value2").val();
                    if ($("#02RM_Type4Value3").val() != "")
                        str3 = $("#02RM_Type4Value3").val();
                    $("#02RM_Type4ValueSum").val(parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1));
                    break;

                case "03RM_Type1Value1":
                case "03RM_Type1Value2":
                case "03RM_Type1Value3":
                    if ($("#03RM_Type1Value1").val() != "")
                        str1 = $("#03RM_Type1Value1").val();
                    if ($("#03RM_Type1Value2").val() != "")
                        str2 = $("#03RM_Type1Value2").val();
                    if ($("#03RM_Type1Value3").val() != "")
                        str3 = $("#03RM_Type1Value3").val();
                    $("#03RM_Type1ValueSum").val(parseInt(str1) + parseInt(str2) + parseInt(str3));
                    //核定數預計年節電量 = 175*(申請數量(盞))(集合住宅+辦公大樓+服務業))
                    var sum_03 = parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1);
                    var Val_03 = parseFloat(175 * (sum_03)).toFixed(1);
                    $("#03RM_PreVal").val(Val_03);
                    //未核定數支年節電量(度) = 申請數預期年節電量(度) - 核定數預期年節電量(度)
                    $("#03RM_NotChkVal").val(parseFloat(parseFloat(Val_03) - parseFloat($("#03RM_ChkVal").val())).toFixed(1));
                    break;
                case "03RM_Type2Value1":
                case "03RM_Type2Value2":
                case "03RM_Type2Value3":
                    if ($("#03RM_Type2Value1").val() != "")
                        str1 = $("#03RM_Type2Value1").val();
                    if ($("#03RM_Type2Value2").val() != "")
                        str2 = $("#03RM_Type2Value2").val();
                    if ($("#03RM_Type2Value3").val() != "")
                        str3 = $("#03RM_Type2Value3").val();
                    $("#03RM_Type2ValueSum").val(parseInt(str1) + parseInt(str2) + parseInt(str3));
                    //核定數預計年節電量 = 175*(申請數量(盞))(集合住宅+辦公大樓+服務業))
                    var sum_03 = parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1);
                    var Val_03 = parseFloat(175 * (sum_03)).toFixed(1);
                    $("#03RM_ChkVal").val(Val_03);
                    //未核定數支年節電量(度) = 申請數預期年節電量(度) - 核定數預期年節電量(度)
                    $("#03RM_NotChkVal").val(parseFloat(parseFloat($("#03RM_PreVal").val()) - parseFloat(Val_03)).toFixed(1));
                    break;
                case "03RM_Type3Value1":
                case "03RM_Type3Value2":
                case "03RM_Type3Value3":
                    if ($("#03RM_Type3Value1").val() != "")
                        str1 = $("#03RM_Type3Value1").val();
                    if ($("#03RM_Type3Value2").val() != "")
                        str2 = $("#03RM_Type3Value2").val();
                    if ($("#03RM_Type3Value3").val() != "")
                        str3 = $("#03RM_Type3Value3").val();
                    $("#03RM_Type3ValueSum").val(parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1));
                    break;
                case "03RM_Type4Value1":
                case "03RM_Type4Value2":
                case "03RM_Type4Value3":
                    if ($("#03RM_Type4Value1").val() != "")
                        str1 = $("#03RM_Type4Value1").val();
                    if ($("#03RM_Type4Value2").val() != "")
                        str2 = $("#03RM_Type4Value2").val();
                    if ($("#03RM_Type4Value3").val() != "")
                        str3 = $("#03RM_Type4Value3").val();
                    $("#03RM_Type4ValueSum").val(parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1));
                    break;

                case "04RM_Type1Value1":
                case "04RM_Type1Value2":
                case "04RM_Type1Value3":
                    if ($("#04RM_Type1Value1").val() != "")
                        str1 = $("#04RM_Type1Value1").val();
                    if ($("#04RM_Type1Value2").val() != "")
                        str2 = $("#04RM_Type1Value2").val();
                    if ($("#04RM_Type1Value3").val() != "")
                        str3 = $("#04RM_Type1Value3").val();
                    $("#04RM_Type1ValueSum").val(parseInt(str1) + parseInt(str2) + parseInt(str3));
                    //核定數預計年節電量 = 40*100*(申請數量(套))(機關 +學校 +服務業))
                    var sum_04 = parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1);
                    var Val_04 = parseFloat(40 * 100 * (sum_04)).toFixed(1);
                    $("#04RM_PreVal").val(Val_04);
                    //未核定數支年節電量(度) = 申請數預期年節電量(度) - 核定數預期年節電量(度)
                    $("#04RM_NotChkVal").val(parseFloat(parseFloat(Val_04) - parseFloat($("#04RM_ChkVal").val())).toFixed(1));
                    break;
                case "04RM_Type2Value1":
                case "04RM_Type2Value2":
                case "04RM_Type2Value3":
                    if ($("#04RM_Type2Value1").val() != "")
                        str1 = $("#04RM_Type2Value1").val();
                    if ($("#04RM_Type2Value2").val() != "")
                        str2 = $("#04RM_Type2Value2").val();
                    if ($("#04RM_Type2Value3").val() != "")
                        str3 = $("#04RM_Type2Value3").val();
                    $("#04RM_Type2ValueSum").val(parseInt(str1) + parseInt(str2) + parseInt(str3));
                    //核定數預計年節電量 = 40*100*(申請數量(套))(機關 +學校 +服務業))
                    var sum_04 = parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1);
                    var Val_04 = parseFloat(40 * 100 * (sum_04)).toFixed(1);
                    $("#04RM_ChkVal").val(Val_04);
                    //未核定數支年節電量(度) = 申請數預期年節電量(度) - 核定數預期年節電量(度)
                    $("#04RM_NotChkVal").val(parseFloat(parseFloat($("#04RM_PreVal").val()) - parseFloat(Val_04)).toFixed(1));
                    break;
                case "04RM_Type3Value1":
                case "04RM_Type3Value2":
                case "04RM_Type3Value3":
                    if ($("#04RM_Type3Value1").val() != "")
                        str1 = $("#04RM_Type3Value1").val();
                    if ($("#04RM_Type3Value2").val() != "")
                        str2 = $("#04RM_Type3Value2").val();
                    if ($("#04RM_Type3Value3").val() != "")
                        str3 = $("#04RM_Type3Value3").val();
                    $("#04RM_Type3ValueSum").val(parseInt(str1) + parseInt(str2) + parseInt(str3));
                    break;

                case "05RM_Type1Value1":
                case "05RM_Type1Value2":
                case "05RM_Type1Value3":
                    if ($("#05RM_Type1Value1").val() != "")
                        str1 = $("#05RM_Type1Value1").val();
                    if ($("#05RM_Type1Value2").val() != "")
                        str2 = $("#05RM_Type1Value2").val();
                    if ($("#05RM_Type1Value3").val() != "")
                        str3 = $("#05RM_Type1Value3").val();
                    $("#05RM_Type1ValueSum").val(parseInt(str1) + parseInt(str2) + parseInt(str3));
                    //核定數預計年節電量 = 312*1000*(申請數量(套))(機關 +學校 +服務業))
                    var sum_05= parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1);
                    var Val_05 = parseFloat(312 * 1000 * (sum_05)).toFixed(1);
                    $("#05RM_PreVal").val(Val_05);
                    //未核定數支年節電量(度) = 申請數預期年節電量(度) - 核定數預期年節電量(度)
                    $("#05RM_NotChkVal").val(parseFloat(parseFloat(parseFloat(Val_05) - $("#05RM_ChkVal").val())).toFixed(1));
                    break;
                case "05RM_Type2Value1":
                case "05RM_Type2Value2":
                case "05RM_Type2Value3":
                    if ($("#05RM_Type2Value1").val() != "")
                        str1 = $("#05RM_Type2Value1").val();
                    if ($("#05RM_Type2Value2").val() != "")
                        str2 = $("#05RM_Type2Value2").val();
                    if ($("#05RM_Type2Value3").val() != "")
                        str3 = $("#05RM_Type2Value3").val();
                    $("#05RM_Type2ValueSum").val(parseInt(str1) + parseInt(str2) + parseInt(str3));
                    //核定數預計年節電量 = 312*1000*(申請數量(套))(機關 +學校 +服務業))
                    var sum_05= parseFloat(parseFloat(str1) + parseFloat(str2) + parseFloat(str3)).toFixed(1);
                    var Val_05 = parseFloat(312 * 1000 * (sum_05)).toFixed(1);
                    $("#05RM_ChkVal").val(Val_05);
                    //未核定數支年節電量(度) = 申請數預期年節電量(度) - 核定數預期年節電量(度)
                    $("#05RM_NotChkVal").val(parseFloat(parseFloat($("#05RM_PreVal").val()) - parseFloat(Val_05)).toFixed(1));
                    break;
                case "05RM_Type3Value1":
                case "05RM_Type3Value2":
                case "05RM_Type3Value3":
                    if ($("#05RM_Type3Value1").val() != "")
                        str1 = $("#05RM_Type3Value1").val();
                    if ($("#05RM_Type3Value2").val() != "")
                        str2 = $("#05RM_Type3Value2").val();
                    if ($("#05RM_Type3Value3").val() != "")
                        str3 = $("#05RM_Type3Value3").val();
                    $("#05RM_Type3ValueSum").val(parseInt(str1) + parseInt(str2) + parseInt(str3));
                    break;
            }
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
                    str_rcreporttype: "01"
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
                <div class="OchiCell OchiTitle TitleSetWidth">期數</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100">
                        <%--<select id="txt_stage">
                            <option value="">請選擇</option>
                            <option value="1">第一期</option>
                            <option value="2">第二期</option>
                            <option value="3">第三期</option>
                        </select>--%>
                        <div class="OchiCellInner nowrap textcenter" style="text-align: left;" id="txt_datesanddatee"></div>
                    </div>
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
                        <div class="OchiCellInner nowrap textcenter" style="text-align: left;" id="txt_yearmonth">
                            <!--<select id="txt_year"></select>年
                            <select id="txt_month">
                                <option value="">請選擇</option>
                                <option value="01">1月</option>
                                <option value="02">2月</option>
                                <option value="03">3月</option>
                                <option value="04">4月</option>
                                <option value="05">5月</option>
                                <option value="06">6月</option>
                                <option value="07">7月</option>
                                <option value="08">8月</option>
                                <option value="09">9月</option>
                                <option value="10">10月</option>
                                <option value="11">11月</option>
                                <option value="12">12月</option>
                            </select>月-->
                        </div>
                    </div>
                    <!-- OchiTableInner -->
                    <!-- cell內容end -->
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->


            <div class="font-bold font-title margin10T" id="div_nowMonthSchedule" style="font-size:16px;"></div>
            <div id="div_data">
                
            </div>
        </div>
        <!-- OchiTrasTable 儲存&取消按鈕 -->
        <div class="twocol margin15T margin5B">
            <div class="right">
                <a href="javascript:void(0);" class="genbtn pbtn" id="btn_goCheck" style="color:red;">送審</a>
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
    
</asp:Content>

