<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="CheckPoint.aspx.cs" Inherits="WebPage_CheckPoint" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <%--common--%>
    <script type="text/javascript">
        $(document).ready(function () {
            $("#loadimg").show();
            getData("01");
            getData("02");
            getData("03");
            getData("04");

            $(document).on("change", "#period", function () {
                if (confirm('請確認資料已儲存，是否切換期別？')) {
                    $("#loadimg").show();
                    $("#gifimg").show();
                    getData("01");
                    getData("02");
                    getData("03");
                    getData("04");
                }
            });

            // 上移
            $(document).on("click", "a[name='sortUp']", function () {
                var PrevID = $(this).closest("tr").prev().attr("class");
                var thisTr = '';
                if ($(this).closest("tr").prev().length > 0) {
                    $("." + $(this).closest("tr").attr("class")).each(function (i) {
                        thisTr += this.outerHTML;
                    });
                    $("." + $(this).closest("tr").attr("class")).remove();
                    $("#" + PrevID).before(thisTr);
                }
            });

            // 下移
            $(document).on("click", "a[name='sortDown']", function () {
                var thisClass = $(this).closest("tr").attr("class");
                var NextID = '';
                var thisTr = '';
                if ($(this).closest("tr").next().length > 0) {
                    $("." + thisClass).each(function (i) {
                        thisTr += this.outerHTML;
                        // Get Next Group ID
                        if ((i + 1) == $("." + thisClass).length) {
                            if ($(this).closest("tr").next().length > 0)
                                NextID = $(this).closest("tr").next().attr("id");
                        }
                    });
                }
                if (NextID != "") {
                    $("." + $(this).closest("tr").attr("class")).remove();

                    $("." + NextID).each(function (i) {
                        if ((i + 1) == $("." + NextID).length)
                            $(this).closest("tr").after(thisTr);
                    });
                }
            });

            // 上下移預設的Value 不會變,改過未存檔變換位子會回預設值
            $(document).on("change", "input", function () {
                $(this).attr("value", this.value);
            });
        });

        function randomId() {
            return 'xxxxxxxx'.replace(/[x]/g, function (c) {
                var r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
                return v.toString(16);
            });
        }

        function feedback(str,type) {
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
                if (type == "01")
                    getData("01");
                else if (type == "02")
                    getData("02");
                else if (type == "03")
                    getData("03");
                else
                    getData("04");
            }
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
                
            }
        }

        function getData(tp) {
            $.ajax({
                type: "POST",
                async: true, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getCheckPoint.aspx",
                data: {
                    period: $("#period").val(),
                    type: tp,
                    person_id: $.getParamValue('v')
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if ($(data).find("Error").length > 0) {
                        alert($(data).find("Error").attr("Message"));
                    }
                    else {
                        var TagName = "";
                        var InputCode = "";
                        switch (tp) {
                            case "01":
                                TagName = "#basicworkTab";
                                InputCode = "";
                                break;
                            case "02":
                                TagName = "#placeTab";
                                InputCode = "p_";
                                break;
                            case "03":
                                TagName = "#smartTab";
                                InputCode = "s_";
                                break;
                            case "04":
                                TagName = "#allowanceTab";
                                InputCode = "a_";
                                break;
                        }

                        $(TagName + " tbody").empty();
                        var tabstr = '';
                        if ($(data).find("PushItem[P_Type='" + tp + "']").length > 0) {
                            $(data).find("PushItem[P_Type='" + tp + "']").each(function (i) {
                                var rid = randomId();
                                //tabstr += '<tbody>';
                                //First row
                                tabstr += '<tr id="' + rid + '" class="' + rid + '">';
                                tabstr += '<td nowrap="nowrap" rowspan="' + $(this).children().length + '">';
                                switch (tp) {
                                    default:
                                        tabstr += '<input type="text" name="' + InputCode + 'pushitem" class="inputex width85" maxlength="50" value=' + $(this).attr("P_ItemName") + ' />&nbsp;';
                                        break;
                                    case "03":
                                        tabstr += getddl_func("07", $(this).attr("P_ItemNameCode"));
                                        break;
                                    case "04":
                                        tabstr += getddl_func("09", $(this).attr("P_ItemNameCode"));
                                        break;
                                }
                                tabstr += '<input type="hidden" name="' + InputCode + 'pgid" value="' + $(this).attr("P_Guid") + '" />';
                                tabstr += '<a href="javascript:void(0);" name="' + InputCode + 'delpibtn" aid="' + rid + '"><img src= "../App_Themes/images/icon-delete-new.png" /></a >';
                                tabstr += '<br>';
                                tabstr += '<a href="javascript:void(0);" name="sortUp">上移</a>&nbsp;&nbsp;<a href="javascript:void(0);" name="sortDown">下移</a>';
                                tabstr += '</td>';
                                if (tp == "04") {
                                    tabstr += '<td rowspan="' + $(this).children().length + '">';
                                    tabstr += '<input class="inputex" type="text" name="exFinish" value="' + $(this).attr("P_ExFinish") + '" style="width:100%;" />';
                                    tabstr += '</td>';
                                }
                                tabstr += '<td><input type="text" name="' + InputCode + 'cp_no" class="inputex width100" maxlength="10" value="' + $(this).children().find("CP_Point")[0].textContent.trim() + '" /></td>';
                                tabstr += '<td align="center">' + ddlYearAndMonth($(this).children().find("CP_ReserveYear")[0].textContent.trim(), $(this).children().find("CP_ReserveMonth")[0].textContent.trim(), InputCode) + '</td>';
                                tabstr += '<td><input type="text" name="' + InputCode + 'cp_desc" class="inputex width86" value="' + $(this).children().find("CP_Desc")[0].textContent.trim() + '" />&nbsp;';
                                var LastStatus = "";
                                if ($(this).children().length == 1) {
                                    tabstr += '<a href="javascript:void(0);" name="' + InputCode + 'addcpbtn" aid="' + rid + '" style="margin-right:5px;"><img src="../App_Themes/images/icon-add-new.png" /></a>';
                                    LastStatus = "Y";
                                }
                                else {
                                    tabstr += '<a href="javascript:void(0);" name="' + InputCode + 'delcpbtn" aid="' + rid + '" style="margin-right:5px;"><img src="../App_Themes/images/icon-delete-new.png" /></a>';
                                }
                                tabstr += '<input name="' + InputCode + 'lastitem" type="hidden" value="' + LastStatus + '" />';
                                tabstr += '<input type="hidden" name="' + InputCode + 'cpgid" value="' + $(this).children().find("CP_Guid")[0].textContent.trim() + '" /></td>';
                                tabstr += '</tr>';
                                $(this).children().each(function (i) {
                                    // 跳過第 1 筆 Row
                                    if (i != 0) {
                                        // 第 2 ~ 倒數第 2 筆
                                        if (($(this).parent().children().length - 1) != i) {
                                            tabstr += '<tr class="' + rid + '">';
                                            tabstr += '<td><input type="text" name="' + InputCode + 'cp_no" class="inputex width100" maxlength="10" value="' + $(this).children("CP_Point").text().trim() + '" /></td>';
                                            tabstr += '<td align="center">' + ddlYearAndMonth($(this).children("CP_ReserveYear").text().trim(), $(this).children("CP_ReserveMonth").text().trim(), InputCode) + '</td>';
                                            tabstr += '<td><input type="text" name="' + InputCode + 'cp_desc" class="inputex width86" value="' + $(this).children("CP_Desc").text().trim() + '" />&nbsp;';
                                            tabstr += '<a href="javascript:void(0);" name="' + InputCode + 'delcpbtn" aid="' + rid + '" style="margin-right:5px;"><img src="../App_Themes/images/icon-delete-new.png" /></a>';
                                            tabstr += '<input name="' + InputCode + 'lastitem" type="hidden" value="" /><input type="hidden" name="' + InputCode + 'cpgid" value="' + $(this).children("CP_Guid").text().trim() + '" /></td>';
                                            tabstr += '</tr>';
                                        }
                                        // 最後 1 筆
                                        else {
                                            tabstr += '<tr class="' + rid + '">';
                                            tabstr += '<td><input type="text" name="' + InputCode + 'cp_no" class="inputex width100" maxlength="10" value="' + $(this).children("CP_Point").text().trim() + '" /></td>';
                                            tabstr += '<td align="center">' + ddlYearAndMonth($(this).children("CP_ReserveYear").text().trim(), $(this).children("CP_ReserveMonth").text().trim(), InputCode) + '</td>';
                                            tabstr += '<td><input type="text" name="' + InputCode + 'cp_desc" class="inputex width86" value="' + $(this).children("CP_Desc").text().trim() + '" />&nbsp;';
                                            tabstr += '<a href="javascript:void(0);" name="' + InputCode + 'delcpbtn" aid="' + rid + '" style="margin-right:5px;"><img src="../App_Themes/images/icon-delete-new.png" /></a>';
                                            tabstr += '<a href="javascript:void(0);" name="' + InputCode + 'addcpbtn" aid="' + rid + '" style="margin-right:5px;"><img src="../App_Themes/images/icon-add-new.png" /></a>';
                                            tabstr += '<input name="' + InputCode + 'lastitem" type="hidden" value="Y" /><input type="hidden" name="' + InputCode + 'cpgid" value="' + $(this).children("CP_Guid").text().trim() + '" /></td>';
                                            tabstr += '</tr>';
                                        }
                                    }
                                });
                                //tabstr += '</tbody>';
                            });
                            $(TagName).append(tabstr);
                            // 權限
                            if ($(data).find("unVisiable").text() == "Y") {
                                $(".inputex").attr("disabled", "disabled");
                                $(".pbtn").hide();
                                $("#content img").hide();
                                $("#autoStatus").val("true"); // 關閉自動存檔
                                $("#previouspage").show();
                                $("#nextpage").show();
                            }
                            else {
                                $("#previousstep").show();
                                $("#nextstep").show();
                            }
                        }
                        else {
                            addNewRow(TagName, InputCode); // 無任何查核點時 + 1 筆空的Row
                            if ($("unVisiable", data).text() == "Y") {
                                $(".inputex").attr("disabled", "disabled");
                                $(".pbtn").hide();
                                $("#content img").hide();
                                $("#autoStatus").val("true"); // 關閉自動存檔
                                $("#previouspage").show();
                                $("#nextpage").show();
                            }
                            else {
                                $("#previousstep").show();
                                $("#nextstep").show();
                            }
                        }
                    }
                    $("#loadimg").hide();
                }
            });
        }

        function addNewRow(TabName, NameCode) {
            var rid = randomId();
            var tabstr = '<tr id="' + rid + '" class="' + rid + '">';
            tabstr += '<td nowrap="nowrap" rowspan="1">';
            switch (TabName) {
                default:
                    tabstr += '<input type="text" name="' + NameCode + 'pushitem" class="inputex width85" maxlength="50" />&nbsp;';
                    break;
                case "#smartTab":
                    tabstr += getddl_func("07", "");
                    break;
                case "#allowanceTab":
                    tabstr += getddl_func("09", "");
                    break;
            }
            tabstr += '<input type="hidden" name="' + NameCode + 'pgid" value="" />';
            tabstr += '<a href="javascript:void(0);" name="' + NameCode + 'delpibtn" aid="' + rid + '"><img src= "../App_Themes/images/icon-delete-new.png" /></a >';
            //if (TabName == "#allowanceTab") {
            //    tabstr += '<div name="OtherPanel" style="margin-top:5px; display:none;">';
            //    tabstr += '<div style="margin-top:5px;">類別：' + ddl_ExType() + '</div>';
            //    tabstr += '<div name="aSubPanel" style="margin-top:5px; display:none;">';
            //    tabstr += '設備子分類：' + ddl_ExDeType() + '</div>';
            //    tabstr += '</div>';
            //}
            tabstr += '<br>';
            tabstr += '<a href="javascript:void(0);" name="sortUp">上移</a>&nbsp;&nbsp;<a href="javascript:void(0);" name="sortDown">下移</a>';
            tabstr += '</td>';
            if (TabName == "#allowanceTab") {
                tabstr += '<td rowspan="1">';
                tabstr += '<input class="inputex" type="text" name="exFinish" style="width:100%;" />';
                tabstr += '</td>';
            }
            tabstr += '<td><input type="text" name="' + NameCode + 'cp_no" class="inputex width100" maxlength="10" /></td>';
            tabstr += '<td align="center">' + ddlYearAndMonth(null, null, NameCode) + '</td>';
            tabstr += '<td><input type="text" name="' + NameCode + 'cp_desc" class="inputex width86" />&nbsp;';
            tabstr += '<a href="javascript:void(0);" name="' + NameCode + 'addcpbtn" aid="' + rid + '" style="margin-right:5px;"><img src="../App_Themes/images/icon-add-new.png" /></a>';
            tabstr += '<input name="' + NameCode + 'lastitem" type="hidden" value="Y" /><input type="hidden" name="' + NameCode + 'cpgid" value="" /></td>';
            tabstr += '</tr>';
            $(TabName).append(tabstr);
        }

        function ddlYearAndMonth(y, m, NameCode) {
            var thisyear = parseInt(new Date().getFullYear()) - 1911;
            y = (y == null) ? thisyear : y;
            var startyear = 107;
			//var endyear = thisyear + 2;
			var endyear = 112;
            var str = '<select name="' + NameCode + 'cp_year" class="inputex">';
            for (var i = startyear; i <= endyear; i++) {
                if (y != null) {
                    if (parseInt(y) == i)
                        str += '<option value="' + i + '" selected="selected">' + i + '</option>';
                    else
                        str += '<option value="' + i + '">' + i + '</option>';
                }
                else
                    str += '<option value="' + i + '">' + i + '</option>';
            }
            str += '</select>';
            str += '&nbsp;年&nbsp;';
            str += '<select  name="' + NameCode + 'cp_month" class="inputex">';
            if (parseInt(m) == 3)
                str += '<option value="3" selected="selected">3</option>';
            else
                str += '<option value="3">3</option>';

            if (parseInt(m) == 6)
                str += '<option value="6" selected="selected">6</option>';
            else
                str += '<option value="6">6</option>';

            if (parseInt(m) == 9)
                str += '<option value="9" selected="selected">9</option>';
            else
                str += '<option value="9">9</option>';

            if (parseInt(m) == 12)
                str += '<option value="12" selected="selected">12</option>';
            else
                str += '<option value="12">12</option>';
            str += '</select>&nbsp;月';
            return str;
        }

        function getddl_func(g,v) {
            var str = '';
            if (g == "07")
                str = '<select name="s_pushitem" class="inputex">';
            else
                str = '<select name="a_pushitem" class="inputex">';

            str += '<option value="">-----------------請 選 擇-----------------</option>';

            $.ajax({
                 type: "POST",
                 async: false, //在沒有返回值之前,不會執行下一步動作
                 url: "../handler/GetDDL.aspx",
                 data: {
                     group: g
                 },
                 error: function (xhr) {
                     alert(xhr.responseText);
                 },
                 success: function (data) {
                     if ($(data).find("Error").length > 0) {
                         alert($(data).find("Error").attr("Message"));
                     }
                     else {
                         if ($(data).find("data_item").length > 0) {
                             $(data).find("data_item").each(function () {
                                 if ($(this).children("C_Item").text().trim() == v)
                                     str += '<option value="' + $(this).children("C_Item").text().trim() + '" selected="selected">' + $(this).children("C_Item_cn").text().trim() + '</option>';
                                 else
                                     str += '<option value="' + $(this).children("C_Item").text().trim() + '">' + $(this).children("C_Item_cn").text().trim() + '</option>';
                             });
                         }
                     }
                 }
            });

            str += '</select>&nbsp;';
            return str;
       }
    </script>
    <%--節電基礎工作--%>
    <script type="text/javascript">
        $(document).ready(function () {
            $("#person_id").val($.getParamValue('v'));
            $('#main_tabs').easytabs();

            //新增推動項目 button
            $(document).on("click", "#bw_addbtn", function () {
                addNewRow("#basicworkTab","");
            });

            //(上)下一步/(上)下一頁 button
            $(document).on("click", "#nextstep,#previousstep,#nextpage,#previouspage", function () {
                switch (this.id) {
                    case "previousstep":
                        if (confirm("請確認資料已儲存，是否回到上一步？")) {
                            location.href = "ProjectInfo.aspx?v=" + $.getParamValue('v');
                        }
                        break;
                    case "nextstep":
                        if (confirm("請確認資料已儲存，是否進入下一步？")) {
                            location.href = "Progress.aspx?v=" + $.getParamValue('v');
                        }
                        break;
                    case "previouspage":
                        location.href = "ProjectInfo.aspx?v=" + $.getParamValue('v');
                        break;
                    case "nextpage":
                        location.href = "Progress.aspx?v=" + $.getParamValue('v');
                        break;
                }
            });

            //刪除推動項目整個 group
            $(document).on("click", "a[name='delpibtn']", function () {
                $("." + $(this).attr("aid")).remove();

                var tmpstr = $("#del_pguid").val();
                if ($(this).closest('tr').find("input[name='pgid']").val() != "") {
                    if (tmpstr != "") tmpstr += ",";
                    tmpstr += $(this).closest('tr').find("input[name='pgid']").val();
                }
                $("#del_pguid").val(tmpstr);

                if ($("#basicworkTab tbody tr").length == 0)
                    addNewRow("#basicworkTab","");
            });

            //新增 row
            $(document).on("click", "a[name='addcpbtn']", function () {
                //Get This Row
                var thisRow = $(this).closest('tr');
                //rowspan + 1
                var rsnum = $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan");
                $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan", parseInt(rsnum) + 1);
                //尾項標記設空
                $(this).closest('tr').find("input[name='lastitem']").val("");
                //圖片改為刪除
                $(this).attr("name", "delcpbtn");
                $(this).children().attr("src", "../App_Themes/images/icon-delete-new.png");
                //若有兩個按鈕要刪除一個
                if ($(this).closest('tr').find("a[name='delcpbtn']").length == 2)
                    $(this).remove();
                //新增一筆row
                var tabstr = '<tr class="' + $(this).attr("aid") +'">';
                tabstr += '<td><input type="text" name="cp_no" class="inputex width100" maxlength="10" /></td>';
                tabstr += '<td align="center">' + ddlYearAndMonth(null, null, "") + '</td>';
                tabstr += '<td><input type="text" name="cp_desc" class="inputex width86" />&nbsp;';
                tabstr += '<a href="javascript:void(0);" name="delcpbtn" aid="' + $(this).attr("aid") + '" style="margin-right:5px;"><img src= "../App_Themes/images/icon-delete-new.png" /></a>';
                tabstr += '<a href="javascript:void(0);" name="addcpbtn" aid="' + $(this).attr("aid") + '"><img src= "../App_Themes/images/icon-add-new.png" /></a>';
                tabstr += '<input name="lastitem" type="hidden" value="Y" /><input type="hidden" name="cpgid" value="" /></td >';
                tabstr += '</tr>';
                thisRow.after(tabstr);
            });

            //刪除 row
            $(document).on("click", "a[name='delcpbtn']", function () {
                //get 推動項目的 rowspan
                var rsnum = $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan");
                //如果要刪的是推動項目的第一個Row
                if ($(this).closest('tr').attr("id") != null) {
                    var PushItemStr = '<td nowrap="nowrap" rowspan="' + (parseInt(rsnum) - 1) + '"><input type="text" name="pushitem" value="' + $(this).closest('tr').find("input[name='pushitem']").val() + '" class="inputex width85" />&nbsp;';
                    PushItemStr += '<input type="hidden" name="pgid" value="' + $(this).closest('tr').find("input[name='pgid']").val() + '" />';
                    PushItemStr += '<a href="javascript:void(0);" name="delpibtn" aid="' + $(this).closest('tr').attr("id") + '" style="margin-right:5px;"><img src= "../App_Themes/images/icon-delete-new.png" /></a ></td>';

                    $(this).closest('tr').next('tr').attr("id", $(this).closest('tr').attr("id"));
                    $(this).closest('tr').next('tr').find('td:first').before(PushItemStr);

                    //判斷next row是否有新增按鈕
                    if ($(this).closest('tr').next('tr').find("a[name='addcpbtn']").length > 0)
                        $(this).closest('tr').next('tr').find("a[name='delcpbtn']").remove();
                }
                else {
                    $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan", parseInt(rsnum) - 1);

                    if ($(this).closest('tr').find("a[name='addcpbtn']").length > 0) {
                        //判斷是否為該推動項目的第一項查核點
                        if ($(this).closest('tr').prev('tr').attr("id") == null) {
                            $(this).closest('tr').prev('tr').find('td:last').append('<a href="javascript:void(0);" name="addcpbtn" aid="' + $(this).attr("aid") + '"><img src= "../App_Themes/images/icon-add-new.png" /></a>');
                            $(this).closest('tr').prev('tr').find('td:last input[name="lastitem"]').val("Y");
                        }
                        else {
                            //如果上一層是第一項，最後一個td重湊
                            var descStr = $(this).closest('tr').prev('tr').find('td:last input[name="cp_desc"]').val();
                            var CBstr = '<input type="text" name="cp_desc" class="inputex width86" value="' + descStr + '" />&nbsp;';
                            CBstr += '<a href="javascript:void(0);" name="addcpbtn" aid="' + $(this).attr("aid") + '" style="margin-right:5px;"><img src="../App_Themes/images/icon-add-new.png" /></a>';
                            CBstr += '<input name="lastitem" type="hidden" value="Y" /><input type="hidden" name="cpgid" value="" />';
                            $(this).closest('tr').prev('tr').find('td:last').html(CBstr);
                        }
                    }
                }
                //Remove Row
                $(this).parent().parent().remove();

                //暫存刪除GUID給後臺做資料庫處理
                var tmpstr = $("#del_cpguid").val();
                if ($(this).closest('tr').find("input[name='cpgid']").val() != "") {
                    if (tmpstr != "") tmpstr += ",";
                    tmpstr += $(this).closest('tr').find("input[name='cpgid']").val();
                }
                $("#del_cpguid").val(tmpstr);
            });

            //儲存 button
            $(document).on("click", "#bw_savebtn", function () {
                var msg = "";
                var tmpary = [];
                $("#basicworkTab tbody").find("tr").each(function () {
                    var str = $(this).attr("class") + $(this).find("select[name='cp_year']").val() + $(this).find("select[name='cp_month']").val();
                    if ($.inArray(str, tmpary) == -1)
                        tmpary.push(str);
                    else
                        msg = "訊息：同一筆推動項目的預定時間不得重複";
                });

                var pi_null = false;
                $("input[name='pushitem']").each(function () {
                    if (this.value.trim() == "")
                        pi_null = true;
                });

                if (pi_null == true)
                    msg = "訊息：推動項目不得為空";

                $("#bw_errorMsg").html(msg);
                if (msg != "")
                    return false;

                //查核點&查核點概述 convert to XML post to backend
                var xmldoc = document.createElement("root");
                $("input[name='cp_no']").each(function (i) {
                    var xNode = document.createElement("cpItem");
                    var Node = document.createElement("no");
                    Node.textContent = this.value;
                    var Node2 = document.createElement("desc");
                    Node2.textContent = $("input[name='cp_desc']")[i].value;
                    xNode.appendChild(Node);
                    xNode.appendChild(Node2);
                    xmldoc.appendChild(xNode);
                });

                var iframe = $('<iframe name="postiframe" id="postiframe" style="display: none" />');
                var category = $('<input type="hidden" name="category" id="category" value="01" />');
                var tmpXML = $('<input type="hidden" name="tmpXML" id="tmpXML" value="' + encodeURIComponent(xmldoc.outerHTML) + '" />');

                var form = $("form")[0];

                $("#postiframe").remove();
                $("input[name='category']").remove();
                $("input[name='tmpXML']").remove();

                form.appendChild(iframe[0]);
                form.appendChild(category[0]);
                form.appendChild(tmpXML[0]);

                form.setAttribute("action", "../handler/addCheckPoint.ashx");
                form.setAttribute("method", "post");
                form.setAttribute("enctype", "multipart/form-data");
                form.setAttribute("encoding", "multipart/form-data");
                form.setAttribute("target", "postiframe");
                form.submit();
            });
        });//js end
    </script>
    <%--因地制宜--%>
    <script type="text/javascript">
        $(document).ready(function () {
            //新增推動項目 button
            $(document).on("click", "#p_addbtn", function () {
                addNewRow("#placeTab", "p_");
            });

            //刪除推動項目整個 group
            $(document).on("click", "a[name='p_delpibtn']", function () {
                $("." + $(this).attr("aid")).remove();

                var tmpstr = $("#p_del_pguid").val();
                if ($(this).closest('tr').find("input[name='p_pgid']").val() != "") {
                    if (tmpstr != "") tmpstr += ",";
                    tmpstr += $(this).closest('tr').find("input[name='p_pgid']").val();
                }
                $("#p_del_pguid").val(tmpstr);

                if ($("#placeTab tbody tr").length == 0)
                    addNewRow("#placeTab", "p_");
            });

            //新增 row
            $(document).on("click", "a[name='p_addcpbtn']", function () {
                //Get This Row
                var thisRow = $(this).closest('tr');
                //rowspan + 1
                var rsnum = $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan");
                $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan", parseInt(rsnum) + 1);
                //尾項標記設空
                $(this).closest('tr').find("input[name='p_lastitem']").val("");
                //圖片改為刪除
                $(this).attr("name", "p_delcpbtn");
                $(this).children().attr("src", "../App_Themes/images/icon-delete-new.png");
                //若有兩個按鈕要刪除一個
                if ($(this).closest('tr').find("a[name='p_delcpbtn']").length == 2)
                    $(this).remove();
                var tabstr = '<tr class="' + $(this).attr("aid") + '">';
                tabstr += '<td><input type="text" name="p_cp_no" class="inputex width100" maxlength="10" /></td>';
                tabstr += '<td align="center">' +  ddlYearAndMonth(null, null, "p_") + '</td>';
                tabstr += '<td><input type="text" name="p_cp_desc" class="inputex width86" />&nbsp;';
                tabstr += '<a href="javascript:void(0);" name="p_delcpbtn" aid="' + $(this).attr("aid") + '" style="margin-right:5px;"><img src= "../App_Themes/images/icon-delete-new.png" /></a>';
                tabstr += '<a href="javascript:void(0);" name="p_addcpbtn" aid="' + $(this).attr("aid") + '"><img src= "../App_Themes/images/icon-add-new.png" /></a >';
                tabstr += '<input type="hidden" name="p_lastitem" value="Y" /><input type="hidden" name="p_cpgid" value="" /></td > ';
                tabstr += '</tr>';
                thisRow.after(tabstr);
            });

            //刪除 row
            $(document).on("click", "a[name='p_delcpbtn']", function () {
                //get 推動項目的 rowspan
                var rsnum = $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan");
                //如果要刪的是推動項目的第一個Row
                if ($(this).closest('tr').attr("id") != null) {
                    var PushItemStr = '<td nowrap="nowrap" rowspan="' + (parseInt(rsnum) - 1) + '"><input type="text" name="p_pushitem" value="' + $(this).closest('tr').find("input[name='p_pushitem']").val() + '" class="inputex width85" />&nbsp;';
                    PushItemStr += '<input type="hidden" name="p_pgid" value="' + $(this).closest('tr').find("input[name='p_pgid']").val() + '" />';
                    PushItemStr += '<a href="javascript:void(0);" name="p_delpibtn" aid="' + $(this).closest('tr').attr("id") + '"><img src= "../App_Themes/images/icon-delete-new.png" /></a ></td>';

                    $(this).closest('tr').next('tr').attr("id", $(this).closest('tr').attr("id"));
                    $(this).closest('tr').next('tr').find('td:first').before(PushItemStr);

                    //判斷next row是否有新增按鈕
                    if ($(this).closest('tr').next('tr').find("a[name='p_addcpbtn']").length > 0)
                        $(this).closest('tr').next('tr').find("a[name='p_delcpbtn']").remove();
                }
                else {
                    $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan", parseInt(rsnum) - 1);

                    if ($(this).closest('tr').find("a[name='p_addcpbtn']").length > 0) {
                        //判斷是否為該推動項目的第一項查核點
                        if ($(this).closest('tr').prev('tr').attr("id") == null) {
                            $(this).closest('tr').prev('tr').find('td:last').append('<a href="javascript:void(0);" name="p_addcpbtn" aid="' + $(this).attr("aid") + '"><img src= "../App_Themes/images/icon-add-new.png" /></a>');
                            $(this).closest('tr').prev('tr').find('td:last input[name="p_lastitem"]').val("Y");
                        }
                        else {
                            //如果上一層是第一項，最後一個td重湊
                            var descStr = $(this).closest('tr').prev('tr').find('td:last input[name="p_cp_desc"]').val();
                            var CBstr = '<input type="text" name="p_cp_desc" class="inputex width86" value="' + descStr + '" />&nbsp;';
                            CBstr += '<a href="javascript:void(0);" name="p_addcpbtn" aid="' + $(this).attr("aid") + '" style="margin-right:5px;"><img src="../App_Themes/images/icon-add-new.png" /></a>';
                            CBstr += '<input type="hidden" name="p_lastitem" value="Y" /><input type="hidden" name="p_cpgid" value="" />';
                            $(this).closest('tr').prev('tr').find('td:last').html(CBstr);
                        }
                    }
                }
                $(this).parent().parent().remove();

                var tmpstr = $("#p_del_cpguid").val();
                if ($(this).closest('tr').find("input[name='p_cpgid']").val() != "") {
                    if (tmpstr != "") tmpstr += ",";
                    tmpstr += $(this).closest('tr').find("input[name='p_cpgid']").val();
                }
                $("#p_del_cpguid").val(tmpstr);
            });

            //儲存 button
            $(document).on("click", "#p_savebtn", function () {
                var msg = "";
                var tmpary = [];
                $("#placeTab tbody").find("tr").each(function () {
                    var str = $(this).attr("class") + $(this).find("select[name='p_cp_year']").val() + $(this).find("select[name='p_cp_month']").val();
                    if ($.inArray(str, tmpary) == -1)
                        tmpary.push(str);
                    else
                        msg = "訊息：同一筆推動項目的預定時間不得重複";
                });
                $("#p_errorMsg").html(msg);

                if (msg != "")
                    return false;

                //查核點&查核點概述 convert to XML post to backend
                var xmldoc = document.createElement("root");
                $("input[name='p_cp_no']").each(function (i) {
                    var xNode = document.createElement("cpItem");
                    var Node = document.createElement("no");
                    Node.textContent = this.value;
                    var Node2 = document.createElement("desc");
                    Node2.textContent = $("input[name='p_cp_desc']")[i].value;
                    xNode.appendChild(Node);
                    xNode.appendChild(Node2);
                    xmldoc.appendChild(xNode);
                });

                var iframe = $('<iframe name="postiframe" id="postiframe" style="display: none" />');
                var category = $('<input type="hidden" name="category" id="category" value="02" />');
                var tmpXML = $('<input type="hidden" name="tmpXML" id="tmpXML" value="' + encodeURIComponent(xmldoc.outerHTML) + '" />');

                var form = $("form")[0];

                $("#postiframe").remove();
                $("input[name='category']").remove();
                $("input[name='tmpXML']").remove();

                form.appendChild(iframe[0]);
                form.appendChild(category[0]);
                form.appendChild(tmpXML[0]);

                form.setAttribute("action", "../handler/addCheckPoint.ashx");
                form.setAttribute("method", "post");
                form.setAttribute("enctype", "multipart/form-data");
                form.setAttribute("encoding", "multipart/form-data");
                form.setAttribute("target", "postiframe");
                form.submit();
            });
        });//js end
    </script>
    <%--設備汰換與智慧用電--%>
    <script type="text/javascript">
        $(document).ready(function () {
            //新增推動項目 button
            $(document).on("click", "#s_addbtn", function () {
                addNewRow("#smartTab", "s_");
            });

            //刪除推動項目整個 group
            $(document).on("click", "a[name='s_delpibtn']", function () {
                $("." + $(this).attr("aid")).remove();

                var tmpstr = $("#s_del_pguid").val();
                if ($(this).closest('tr').find("input[name='s_pgid']").val() != "") {
                    if (tmpstr != "") tmpstr += ",";
                    tmpstr += $(this).closest('tr').find("input[name='s_pgid']").val();
                }
                $("#s_del_pguid").val(tmpstr);

                if ($("#smartTab tbody tr").length == 0)
                    addNewRow("#smartTab", "s_");
            });

            //新增 row
            $(document).on("click", "a[name='s_addcpbtn']", function () {
                //Get This Row
                var thisRow = $(this).closest('tr');
                //rowspan + 1
                var rsnum = $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan");
                $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan", parseInt(rsnum) + 1);
                //尾項標記設空
                $(this).closest('tr').find("input[name='s_lastitem']").val("");
                //圖片改為刪除
                $(this).attr("name", "s_delcpbtn");
                $(this).children().attr("src", "../App_Themes/images/icon-delete-new.png");
                //若有兩個按鈕要刪除一個
                if ($(this).closest('tr').find("a[name='s_delcpbtn']").length == 2)
                    $(this).remove();
                var tabstr = '<tr class="' + $(this).attr("aid") + '">';
                tabstr += '<td><input type="text" name="s_cp_no" class="inputex width100" maxlength="10" /></td>';
                tabstr += '<td nowrap="nowrap" align="center">' + ddlYearAndMonth(null, null, "s_") + '</td>';
                tabstr += '<td><input type="text" name="s_cp_desc" class="inputex width86" />&nbsp;';
                tabstr += '<a href="javascript:void(0);" name="s_delcpbtn" aid="' + $(this).attr("aid") + '" style="margin-right:5px;"><img src= "../App_Themes/images/icon-delete-new.png" /></a>';
                tabstr += '<a href="javascript:void(0);" name="s_addcpbtn" aid="' + $(this).attr("aid") + '"><img src= "../App_Themes/images/icon-add-new.png" /></a >';
                tabstr += '<input type="hidden" name="s_lastitem" value="Y" /><input type="hidden" name="s_cpgid" value="" /></td > ';
                tabstr += '</tr>';
                thisRow.after(tabstr);
            });

            //刪除 row
            $(document).on("click", "a[name='s_delcpbtn']", function () {
                //get 推動項目的 rowspan
                var rsnum = $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan");
                //如果要刪的是推動項目的第一個Row
                if ($(this).closest('tr').attr("id") != null) {
                    var PushItemStr = '<td nowrap="nowrap" rowspan="' + (parseInt(rsnum) - 1) + '">';
                    PushItemStr += getddl_func("07", $(this).closest('tr').find("select[name='s_pushitem']").val());
                    PushItemStr += '<input type="hidden" name="s_pgid" value="' + $(this).closest('tr').find("input[name='s_pgid']").val() + '" />';
                    PushItemStr += '<a href="javascript:void(0);" name="s_delpibtn" aid="' + $(this).closest('tr').attr("id") + '"><img src= "../App_Themes/images/icon-delete-new.png" /></a ></td>';

                    $(this).closest('tr').next('tr').attr("id", $(this).closest('tr').attr("id"));
                    $(this).closest('tr').next('tr').find('td:first').before(PushItemStr);

                    //判斷next row是否有新增按鈕
                    if ($(this).closest('tr').next('tr').find("a[name='s_addcpbtn']").length > 0)
                        $(this).closest('tr').next('tr').find("a[name='s_delcpbtn']").remove();
                }
                else {
                    $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan", parseInt(rsnum) - 1);
                    
                    if ($(this).closest('tr').find("a[name='s_addcpbtn']").length > 0) {
                        //判斷是否為該推動項目的第一項查核點
                        if ($(this).closest('tr').prev('tr').attr("id") == null) {
                            $(this).closest('tr').prev('tr').find('td:last').append('<a href="javascript:void(0);" name="s_addcpbtn" aid="' + $(this).attr("aid") + '"><img src= "../App_Themes/images/icon-add-new.png" /></a>');
                            $(this).closest('tr').prev('tr').find('td:last input[name="s_lastitem"]').val("Y");
                        }
                        else {
                            //如果上一層是第一項，最後一個td重湊
                            var descStr = $(this).closest('tr').prev('tr').find('td:last input[name="s_cp_desc"]').val();
                            var CBstr = '<input type="text" name="s_cp_desc" class="inputex width85" value="' + descStr + '" />&nbsp;';
                            CBstr += '<a href="javascript:void(0);" name="s_addcpbtn" aid="' + $(this).attr("aid") + '" style="margin-right:5px;"><img src="../App_Themes/images/icon-add-new.png" /></a>';
                            CBstr += '<input type="hidden" name="s_lastitem" value="Y" /><input type="hidden" name="s_cpgid" value="" />';
                            $(this).closest('tr').prev('tr').find('td:last').html(CBstr);
                        }
                    }
                }
                //Remove Row
                $(this).parent().parent().remove();

                var tmpstr = $("#s_del_cpguid").val();
                if ($(this).closest('tr').find("input[name='s_cpgid']").val() != "") {
                    if (tmpstr != "") tmpstr += ",";
                    tmpstr += $(this).closest('tr').find("input[name='s_cpgid']").val();
                }
                $("#s_del_cpguid").val(tmpstr);
            });

            //儲存 button
            $(document).on("click", "#s_savebtn", function () {
                var msg = "";
                var tmpary = [];
                $("#smartTab tbody").find("tr").each(function () {
                    var str = $(this).attr("class") + $(this).find("select[name='s_cp_year']").val() + $(this).find("select[name='s_cp_month']").val();
                    if ($.inArray(str, tmpary) == -1)
                        tmpary.push(str);
                    else
                        msg = "訊息：同一筆推動項目的預定時間不得重複";
                });
                $("#s_errorMsg").html(msg);

                if (msg != "")
                    return false;

                //查核點&查核點概述 convert to XML post to backend
                var xmldoc = document.createElement("root");
                $("input[name='s_cp_no']").each(function (i) {
                    var xNode = document.createElement("cpItem");
                    var Node = document.createElement("no");
                    Node.textContent = this.value;
                    var Node2 = document.createElement("desc");
                    Node2.textContent = $("input[name='s_cp_desc']")[i].value;
                    xNode.appendChild(Node);
                    xNode.appendChild(Node2);
                    xmldoc.appendChild(xNode);
                });

                var iframe = $('<iframe name="postiframe" id="postiframe" style="display: none" />');
                var category = $('<input type="hidden" name="category" id="category" value="03" />');
                var tmpXML = $('<input type="hidden" name="tmpXML" id="tmpXML" value="' + encodeURIComponent(xmldoc.outerHTML) + '" />');

                var form = $("form")[0];

                $("#postiframe").remove();
                $("input[name='category']").remove();
                $("input[name='tmpXML']").remove();

                form.appendChild(iframe[0]);
                form.appendChild(category[0]);
                form.appendChild(tmpXML[0]);

                form.setAttribute("action", "../handler/addCheckPoint.ashx");
                form.setAttribute("method", "post");
                form.setAttribute("enctype", "multipart/form-data");
                form.setAttribute("encoding", "multipart/form-data");
                form.setAttribute("target", "postiframe");
                form.submit();
            });
        });//js end

        
    </script>
    <%--擴大補助--%>
    <script type="text/javascript">
        $(document).ready(function () {
            //新增推動項目 button
            $(document).on("click", "#a_addbtn", function () {
                addNewRow("#allowanceTab", "a_");
            });

            //刪除推動項目整個 group
            $(document).on("click", "a[name='a_delpibtn']", function () {
                $("." + $(this).attr("aid")).remove();

                var tmpstr = $("#a_del_pguid").val();
                if ($(this).closest('tr').find("input[name='a_pgid']").val() != "") {
                    if (tmpstr != "") tmpstr += ",";
                    tmpstr += $(this).closest('tr').find("input[name='a_pgid']").val();
                }
                $("#a_del_pguid").val(tmpstr);

                if ($("#allowanceTab tbody tr").length == 0)
                    addNewRow("#allowanceTab", "a_");
            });

            //新增 row
            $(document).on("click", "a[name='a_addcpbtn']", function () {
                //Get This Row
                var thisRow = $(this).closest('tr');
                //rowspan + 1
                var rsnum = $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan");
                $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan", parseInt(rsnum) + 1);
                $("#" + $(this).attr("aid") + "").find('td:nth-child(2)').attr("rowspan", parseInt(rsnum) + 1);
                //尾項標記設空
                $(this).closest('tr').find("input[name='a_lastitem']").val("");
                //圖片改為刪除
                $(this).attr("name", "a_delcpbtn");
                $(this).children().attr("src", "../App_Themes/images/icon-delete-new.png");
                //若有兩個按鈕要刪除一個
                if ($(this).closest('tr').find("a[name='a_delcpbtn']").length == 2)
                    $(this).remove();
                var tabstr = '<tr class="' + $(this).attr("aid") + '">';
                tabstr += '<td><input type="text" name="a_cp_no" class="inputex width100" maxlength="10" /></td>';
                tabstr += '<td align="center">' +  ddlYearAndMonth(null, null, "a_") + '</td>';
                tabstr += '<td><input type="text" name="a_cp_desc" class="inputex width86" />&nbsp;';
                tabstr += '<a href="javascript:void(0);" name="a_delcpbtn" aid="' + $(this).attr("aid") + '" style="margin-right:5px;"><img src= "../App_Themes/images/icon-delete-new.png" /></a>';
                tabstr += '<a href="javascript:void(0);" name="a_addcpbtn" aid="' + $(this).attr("aid") + '"><img src= "../App_Themes/images/icon-add-new.png" /></a >';
                tabstr += '<input type="hidden" name="a_lastitem" value="Y" /><input type="hidden" name="a_cpgid" value="" /></td > ';
                tabstr += '</tr>';
                thisRow.after(tabstr);
            });

            //刪除 row
            $(document).on("click", "a[name='a_delcpbtn']", function () {
                //get 推動項目的 rowspan
                var rsnum = $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan");
                //如果要刪的是推動項目的第一個Row
                if ($(this).closest('tr').attr("id") != null) {
                    var PushItemStr = '<td nowrap="nowrap" rowspan="' + (parseInt(rsnum) - 1) + '">';
                    PushItemStr += getddl_func("09", $(this).closest('tr').find("select[name='a_pushitem']").val());
                    PushItemStr += '<input type="hidden" name="a_pgid" value="' + $(this).closest('tr').find("input[name='a_pgid']").val() + '" />';
                    PushItemStr += '<a href="javascript:void(0);" name="a_delpibtn" aid="' + $(this).closest('tr').attr("id") + '"><img src= "../App_Themes/images/icon-delete-new.png" /></a ></td>';
                    PushItemStr += '<td rowspan="' + (parseInt(rsnum) - 1) + '">';
                    PushItemStr += '<input class="inputex" type="text" name="exFinish" style="width:100%;" />';
                    PushItemStr += '</td>';

                    $(this).closest('tr').next('tr').attr("id", $(this).closest('tr').attr("id"));
                    $(this).closest('tr').next('tr').find('td:first').before(PushItemStr);

                    //判斷next row是否有新增按鈕
                    if ($(this).closest('tr').next('tr').find("a[name='a_addcpbtn']").length > 0)
                        $(this).closest('tr').next('tr').find("a[name='a_delcpbtn']").remove();
                }
                else {
                    $("#" + $(this).attr("aid") + "").find('td:first').attr("rowspan", parseInt(rsnum) - 1);
                    $("#" + $(this).attr("aid") + "").find('td:nth-child(2)').attr("rowspan", parseInt(rsnum) - 1);

                    if ($(this).closest('tr').find("a[name='a_addcpbtn']").length > 0) {
                        //判斷是否為該推動項目的第一項查核點
                        if ($(this).closest('tr').prev('tr').attr("id") == null) {
                            $(this).closest('tr').prev('tr').find('td:last').append('<a href="javascript:void(0);" name="a_addcpbtn" aid="' + $(this).attr("aid") + '"><img src= "../App_Themes/images/icon-add-new.png" /></a>');
                            var descStr = $(this).closest('tr').prev('tr').find('td:last input[name="a_lastitem"]').val("Y");
                        }
                        else {
                            //如果上一層是第一項，最後一個td重湊
                            var descStr = $(this).closest('tr').prev('tr').find('td:last input[name="a_cp_desc"]').val();
                            var CBstr = '<input type="text" name="a_cp_desc" class="inputex width86" value="' + descStr + '" />&nbsp;';
                            CBstr += '<a href="javascript:void(0);" name="a_addcpbtn" aid="' + $(this).attr("aid") + '" style="margin-right:5px;"><img src="../App_Themes/images/icon-add-new.png" /></a>';
                            CBstr += '<input type="hidden" name="a_lastitem" value="Y" /><input type="hidden" name="a_cpgid" value="" />';
                            $(this).closest('tr').prev('tr').find('td:last').html(CBstr);
                        }
                    }
                }
                $(this).parent().parent().remove();

                var tmpstr = $("#a_del_cpguid").val();
                if ($(this).closest('tr').find("input[name='a_cpgid']").val() != "") {
                    if (tmpstr != "") tmpstr += ",";
                    tmpstr += $(this).closest('tr').find("input[name='a_cpgid']").val();
                }
                $("#a_del_cpguid").val(tmpstr);
            });

            //儲存 button
            $(document).on("click", "#a_savebtn", function () {
                var msg = "";
                var tmpary = [];
                $("#allowanceTab tbody").find("tr").each(function () {
                    var str = $(this).attr("class") + $(this).find("select[name='a_cp_year']").val() + $(this).find("select[name='a_cp_month']").val();
                    if ($.inArray(str, tmpary) == -1)
                        tmpary.push(str);
                    else
                        msg = "訊息：同一筆推動項目的預定時間不得重複";
                });
                $("#a_errorMsg").html(msg);

                if (msg != "")
                    return false;

                //查核點&查核點概述 convert to XML post to backend
                var xmldoc = document.createElement("root");
                $("input[name='a_cp_no']").each(function (i) {
                    var xNode = document.createElement("cpItem");
                    var Node = document.createElement("no");
                    Node.textContent = this.value;
                    var Node2 = document.createElement("desc");
                    Node2.textContent = $("input[name='a_cp_desc']")[i].value;
                    xNode.appendChild(Node);
                    xNode.appendChild(Node2);
                    xmldoc.appendChild(xNode);
                });

                var iframe = $('<iframe name="postiframe" id="postiframe" style="display: none" />');
                var category = $('<input type="hidden" name="category" id="category" value="04" />');
                var tmpXML = $('<input type="hidden" name="tmpXML" id="tmpXML" value="' + encodeURIComponent(xmldoc.outerHTML) + '" />');

                var form = $("form")[0];

                $("#postiframe").remove();
                $("input[name='category']").remove();
                $("input[name='tmpXML']").remove();

                form.appendChild(iframe[0]);
                form.appendChild(category[0]);
                form.appendChild(tmpXML[0]);

                form.setAttribute("action", "../handler/addCheckPoint.ashx");
                form.setAttribute("method", "post");
                form.setAttribute("enctype", "multipart/form-data");
                form.setAttribute("encoding", "multipart/form-data");
                form.setAttribute("target", "postiframe");
                form.submit();
            });
        });//js end
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <input type="hidden" id="person_id" name="person_id" />
    <input type="hidden" id="del_pguid" name="del_pguid" />
    <input type="hidden" id="del_cpguid" name="del_cpguid" />
    <input type="hidden" id="p_del_pguid" name="p_del_pguid" />
    <input type="hidden" id="p_del_cpguid" name="p_del_cpguid" />
    <input type="hidden" id="s_del_pguid" name="s_del_pguid" />
    <input type="hidden" id="s_del_cpguid" name="s_del_cpguid" />
    <input type="hidden" id="a_del_pguid" name="a_del_pguid" />
    <input type="hidden" id="a_del_cpguid" name="a_del_cpguid" />
    <input type="hidden" id="autoStatus" value="false" />
    <div id="content" class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">計畫查核點</span></div>
            <div class="right">基本資料 / 計畫查核點</div>
        </div><!-- twocol -->
        <div class="font-size3 margin10T">期別：
            <select id="period" name="period">
                <option value="1">第一期</option>
                <option value="2">第二期</option>
                <option value="3">第三期</option>
            </select>
            <span id="loadimg" style="display:none;"><img id="gifimg" src="../App_Themes/images/loading.gif" height="25" />讀取中...</span>
        </div>
        <div id="main_tabs" class='easytabH margin10T'>
            <ul  class="menubar">
                <li><a href="#tabs1">節電基礎工作</a></li>
                <li><a href="#tabs2">設備汰換與智慧用電</a></li>
                <li><a href="#tabs3">因地制宜</a></li>
                <li><a href="#tabs4">擴大補助</a></li>
            </ul>
            <div class='panel-container'>
                <div id="tabs1">
                    <div style="text-align:right; margin-bottom:10px;"><input type="button" id="bw_addbtn" class="genbtn  pbtn" value="新增推動項目" /></div>
                    <div class="stripecomplex margin5T font-normal">
                        <table id="basicworkTab" width="100%" border="0" cellspacing="0" cellpadding="0">
                            <thead>
                                <tr>
                                    <th nowrap="nowrap">推動項目</th>
                                    <th nowrap="nowrap" class="width5">查核點</th>
                                    <th nowrap="nowrap" class="width20">預定時間</th>
                                    <th nowrap="nowrap" class="width50">查核點概述</th>
                                </tr>
                            </thead>
                        </table>
                    </div>
                    <div id="bw_errorMsg" style="color:red;"></div>
                    <div style="text-align:right; margin-top:10px;"><input type="button" id="bw_savebtn" class="genbtn pbtn" value="儲存" /></div>
                </div>
                <div id="tabs2">
                    <div style="text-align:right; margin-bottom:10px;"><input type="button" id="s_addbtn" class="genbtn pbtn" value="新增推動項目" /></div>
                    <div class="stripecomplex margin5T font-normal">
                        <table id="smartTab" width="100%" border="0" cellspacing="0" cellpadding="0">
                            <thead>
                                <tr>
                                    <th nowrap="nowrap">推動項目</th>
                                    <th nowrap="nowrap">查核點</th>
                                    <th nowrap="nowrap" class="width20">預定時間</th>
                                    <th nowrap="nowrap" class="width50">查核點概述</th>
                                </tr>
                            </thead>
                        </table>
                    </div>
                    <div id="s_errorMsg" style="color:red;"></div>
                    <div style="text-align:right; margin-top:10px;"><input type="button" id="s_savebtn" class="genbtn pbtn" value="儲存" /></div>
                </div>
                <div id="tabs3">
                    <div style="text-align:right; margin-bottom:10px;"><input type="button" id="p_addbtn" class="genbtn pbtn" value="新增推動項目" /></div>
                    <div class="stripecomplex margin5T font-normal">
                        <table id="placeTab" width="100%" border="0" cellspacing="0" cellpadding="0">
                            <thead>
                                <tr>
                                    <th nowrap="nowrap">推動項目</th>
                                    <th nowrap="nowrap" class="width10">查核點</th>
                                    <th nowrap="nowrap" class="width20">預定時間</th>
                                    <th nowrap="nowrap" class="width50">查核點概述</th>
                                </tr>
                            </thead>
                        </table>
                    </div>
                    <div id="p_errorMsg" style="color:red;"></div>
                    <div style="text-align:right; margin-top:10px;"><input type="button" id="p_savebtn" class="genbtn pbtn" value="儲存" /></div>
                </div>
                <div id="tabs4">
                    <div style="text-align:right; margin-bottom:10px;"><input type="button" id="a_addbtn" class="genbtn pbtn" value="新增推動項目" /></div>
                    <div class="stripecomplex margin5T font-normal">
                        <table id="allowanceTab" width="100%" border="0" cellspacing="0" cellpadding="0">
                            <thead>
                                <tr>
                                    <th nowrap="nowrap">推動項目</th>
                                    <th nowrap="nowrap" class="width10">預計完成數</th>
                                    <th nowrap="nowrap" class="width10">查核點</th>
                                    <th nowrap="nowrap" class="width25">預定時間</th>
                                    <th nowrap="nowrap" class="width50">查核點概述</th>
                                </tr>
                            </thead>
                        </table>
                    </div>
                    <div id="a_errorMsg" style="color:red;"></div>
                    <div style="text-align:right; margin-top:10px;"><input type="button" id="a_savebtn" class="genbtn pbtn" value="儲存" /></div>
                </div>
            </div>
        </div>
        <div style="text-align:right; margin-top:10px; margin-bottom:10px">
            <input id="previousstep" type="button" class="genbtn" value="上一步" style="display:none;" />&nbsp;
            <input id="nextstep" type="button" class="genbtn" value="下一步" style="display:none;" />
            <input id="previouspage" type="button" class="genbtn" value="上一頁" style="display:none;" />&nbsp;
            <input id="nextpage" type="button" class="genbtn" value="下一頁" style="display:none;" />
        </div>
    </div>
    <script type="text/javascript">
        //AutoSave Function - Auto run every 20 minutes
        //$(document).ready(function () {
        //    setInterval(function () {
        //        var breakStatus = false;
        //        var tmpary = [];
        //        $("#basicworkTab tbody").find("tr").each(function () {
        //            var str = $(this).attr("class") + $(this).find("select[name='cp_year']").val() + $(this).find("select[name='cp_month']").val();
        //            if ($.inArray(str, tmpary) == -1)
        //                tmpary.push(str);
        //            else
        //                breakStatus = true;
        //        });

        //        if (breakStatus == true || $("#autoStatus").val() == "true")
        //            return;

        //        var iframe = $('<iframe name="postiframe" id="postiframe" style="display: none" />');
        //        var form = $("form")[0];

        //        $("#postiframe").remove();

        //        form.appendChild(iframe[0]);

        //        form.setAttribute("action", "../handler/AutoSave_CP.ashx");
        //        form.setAttribute("method", "post");
        //        form.setAttribute("enctype", "multipart/form-data");
        //        form.setAttribute("encoding", "multipart/form-data");
        //        form.setAttribute("target", "postiframe");
        //        form.submit();
        //    }, 1200000);//20 minutes
        //});
    </script>
</asp:Content>

