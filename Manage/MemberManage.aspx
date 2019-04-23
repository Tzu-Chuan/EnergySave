<%@ Page Title="" Language="C#" MasterPageFile="~/Manage/Admin.master" AutoEventWireup="true" CodeFile="MemberManage.aspx.cs" Inherits="Manage_MemberManage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {

            
            getData(0);
            getddl("02", "#ddlCity");
            getddl("03", "#ddlCompetence");

            $(document).on("click", "a[name='sortbtn']", function () {
                $("a[name='sortbtn']").removeClass("asc desc")
                $("#sortName").val($(this).attr("atp"));
                if ($("#sortMethod").val() == "desc") {
                    $("#sortMethod").val("asc");
                    $(this).addClass('asc');
                }
                else {
                    $("#sortMethod").val("desc");
                    $(this).addClass('desc');
                }
                getData(0);
            });

            //限制只能輸入數字
            $(document).on("keyup", ".num", function () {
                if (/[^0-9\-]/g.test(this.value)) {
                    this.value = this.value.replace(/[^0-9\-#]/g, '');
                }
            });

            //CheckBox 帳號密碼同E-Mail  change
            $(document).on("change", "#aprEmail", function () {
                if ($("#aprEmail").is(":checked")) {
                    $(".aptr").hide();
                }
                else {
                    $(".aptr").show();
                }
            });

            //新增 button
            $(document).on("click", "#addbtn", function () {
                $("#actionStr").html("新增");
                $("#idtmp").val("");
                $(".str").val("");
                $("#M_Manager_Name").html("");
                $("#aprEmail").prop("checked", false);
                $(".aptr").show();
                $("#subbtn").show();
                $("#savebtn").hide();
                $(".SpnAPR").show();
                $(".mar").hide();
                $("#ddlcompTd").attr("colspan", "3");
                openEditBlock();
            });

            //身份 change
            $(document).on("change", "#ddlCompetence", function () {
                if (this.value == "01") {
                    $(".mar").show();
                    $("#ddlcompTd").attr("colspan", "");
                }
                else {
                    $(".mar").hide();
                    $("#ddlcompTd").attr("colspan", "3");
                }
            });

            //新增&修改
            $(document).on("click", "#subbtn,#savebtn", function () {
                var msg = "";
                if ($("#M_Name").val() == "")
                    msg += "請輸入【姓名】\n";
                if ($("#M_JobTitle").val() == "")
                    msg += "請輸入【職稱】\n";
                if ($("#M_Email").val() == "")
                    msg += "請輸入【E-Mail】\n";
                if ($("#M_Email").val().trim() != "") {
                    var pattern = new RegExp(/^[+a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/i);
                    if (!pattern.test($("#M_Email").val()))
                        msg += "請輸入正確Email\n";
                }
                if (!$("#aprEmail").is(":checked")) {
                    if ($("#M_Account").val() == "")
                        msg += "請輸入【帳號】\n";
                    if ($("#M_Pwd").val() == "")
                        msg += "請輸入【密碼】\n";
                }
                if ($("#M_Tel").val() == "")
                    msg += "請輸入【電話】\n";
                if ($("#ddlCity").val() == "")
                    msg += "請選擇【執行機關】\n";
                if ($("#M_Office").val() == "")
                    msg += "請輸入【承辦局處】\n";
                if ($("#ddlCompetence").val() == "")
                    msg += "請選擇【權限/身份】";
                if ($("#ddlCompetence").val() == "01" && $("#M_Manager_ID").val()=="")
                    msg += "請選擇【承辦主管】";

                if (msg != "") {
                    alert(msg);
                    return false;
                }

                var status = (this.id == "subbtn") ? "New" : "Mod";
                $.ajax({
                    type: "POST",
                    async: false, //在沒有返回值之前,不會執行下一步動作
                    url: "../handler/addMember.ashx",
                    data: {
                        mode: status,
                        id: $("#idtmp").val(),
                        M_Name: $("#M_Name").val(),
                        M_JobTitle: $("#M_JobTitle").val(),
                        M_Tel: $("#M_Tel").val(),
                        M_Ext: $("#M_Ext").val(),
                        M_Fax: $("#M_Fax").val(),
                        M_Phone: $("#M_Phone").val(),
                        M_Email: $("#M_Email").val(),
                        aprEmail: $("#aprEmail").is(":checked"),
                        M_Account: $("#M_Account").val(),
                        M_Pwd: $("#M_Pwd").val(),
                        Old_Pwd: $("#Old_Pwd").val(),
                        M_Addr: $("#M_Addr").val(),
                        M_City: $("#ddlCity").val(),
                        M_Office: $("#M_Office").val(),
                        M_Competence: $("#ddlCompetence").val(),
                        M_Manager_ID: $("#M_Manager_ID").val()
                    },
                    error: function (xhr) {
                        alert("Error " + xhr.status);
                        console.log(xhr.responseText);
                    },
                    success: function (data) {
                        if (data.indexOf("Error") > -1)
                            alert(data);
                        else {
                            if (data == "MailRepeat") {
                                alert("此E-Mail已存在，請重新輸入");
                                return false;
                            }

                            if (data == "reLogin") {
                                alert("請重新登入");
                                window.location = "../WebPage/Login.aspx";
                                return;
                            }

                            if (data == "succeed") {
                                if (status == "Mod") alert("儲存成功");
                                $.fancybox.close();
                                getData(0);
                            }
                        }
                    }
                });
            });


            //修改 button
            $(document).on("click", "input[name='modbtn']", function () {
                $("#actionStr").html("修改");
                $(".str").val("");
                $("#M_Manager_Name").html("");
                $("#aprEmail").prop("checked", false);
                $(".aptr").show();
                $("#subbtn").hide();
                $("#savebtn").show();
                $(".SpnAPR").hide();
                $("#ddlCompetence").prop("disabled", true);
                var id = $(this).attr("aid");
                $.ajax({
                    type: "POST",
                    async: false, //在沒有返回值之前,不會執行下一步動作
                    url: "../handler/editMember.ashx",
                    data: {
                        mode: "Edit",
                        id: id
                    },
                    error: function (xhr) {
                        alert("Error " + xhr.status);
                        console.log(xhr.responseText);
                    },
                    success: function (data) {
                        if (data.indexOf("Error") > -1)
                            alert(data);
                        else {
                            if (data != null) {
                                data = $.parseXML(data);
                                if ($(data).find("data_item").length > 0) {
                                    $(data).find("data_item").each(function (i) {
                                        if ($(this).children("M_Competence").text().trim() == "01") {
                                            $(".mar").show();
                                            $("#ddlcompTd").attr("colspan", "");
                                        }
                                        else {
                                            $(".mar").hide();
                                            $("#ddlcompTd").attr("colspan", "3");
                                        }

                                        $("#idtmp").val($(this).children("M_ID").text().trim());
                                        $("#M_Name").val($(this).children("M_Name").text().trim());
                                        $("#M_JobTitle").val($(this).children("M_JobTitle").text().trim());
                                        $("#M_Tel").val($(this).children("M_Tel").text().trim());
                                        $("#M_Ext").val($(this).children("M_Ext").text().trim());
                                        $("#M_Fax").val($(this).children("M_Fax").text().trim());
                                        $("#M_Phone").val($(this).children("M_Phone").text().trim());
                                        $("#M_Email").val($(this).children("M_Email").text().trim());
                                        $("#M_Account").val($(this).children("M_Account").text().trim());
                                        $("#M_Pwd").val($(this).children("M_Pwd").text().trim());
                                        $("#Old_Pwd").val($(this).children("M_Pwd").text().trim());
                                        $("#M_Addr").val($(this).children("M_Addr").text().trim());
                                        $("#ddlCity").val($(this).children("M_City").text().trim());
                                        $("#M_Office").val($(this).children("M_Office").text().trim());
                                        $("#ddlCompetence").val($(this).children("M_Competence").text().trim());
                                        $("#M_Manager_Name").html($(this).children("Manager").text().trim());
                                        $("#M_Manager_ID").val($(this).children("M_Manager_ID").text().trim());
                                    });
                                }
                                openEditBlock();
                            }
                        }
                    }
                });
            });

            //刪除 button
            $(document).on("click", "input[name='delbtn']", function () {
                var id = $(this).attr("aid");
                if (confirm("確定刪除？")) {
                    $.ajax({
                        type: "POST",
                        async: false, //在沒有返回值之前,不會執行下一步動作
                        url: "../handler/editMember.ashx",
                        data: {
                            mode: "Delete",
                            id: id
                        },
                        error: function (xhr) {
                            alert("Error " + xhr.status);
                            console.log(xhr.responseText);
                        },
                        success: function (data) {
                            if (data.indexOf("Error") > -1)
                                alert(data);
                            else
                                getData(0);
                        }
                    });
                }
            });

            //取消 fancybox button
            $(document).on("click", "#cancelbtn", function () {
                if (confirm('資料尚未儲存，確定取消？')) 
                    $.fancybox.close();
            });
        });

        function getData(p) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getMemberList.ashx",
                data: {
                    CurrentPage: p,
                    SearchStr: $("#SearchStr").val(),
                    sortMethod: $("#sortMethod").val(),
                    sortName: $("#sortName").val()
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if (data.indexOf("Error") > -1)
                        alert(data);
                    else {
                        if (data != null) {
                            data = $.parseXML(data);
                            $("#tablist tbody").empty();
                            var tabstr = '';
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function (i) {
                                    tabstr += '<tr aid="' + $(this).children("GroupId").text().trim() + '">';
                                    tabstr += '<td align="center" nowrap="nowrap">' + (i + 1) + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("M_Name").text().trim() + '</td>';
                                    tabstr += '<td nowrap="nowrap">' + $(this).children("M_Email").text().trim() + '</td>';
                                    if ($(this).children("M_Competence").text().trim() == "SA")
                                        tabstr += '<td align="center" nowrap="nowrap">系統管理員</td>';
                                    else
                                        tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("City").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("M_Office").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("Comp").text().trim() + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $.datepicker.formatDate('yy/mm/dd', new Date($(this).children("M_CreateDate").text().trim())) + '</td>';
                                    tabstr += '<td align="center"><input type="button" class="genbtnS" name="modbtn" aid="' + $(this).children("M_ID").text().trim() + '" value="修改" />';
                                    tabstr += '<input type="button" class="genbtnS" name= "delbtn" aid= "' + $(this).children("M_ID").text().trim() + '" value="刪除" /></td > ';
                                    tabstr += '</tr>';
                                });
                            }
                            else
                                tabstr += "<tr><td colspan='8'>查詢無資料</td></tr>";
                            $("#tablist tbody").append(tabstr);
                            $(".stripeMe tr").mouseover(function () { $(this).addClass("spe"); }).mouseout(function () { $(this).removeClass("spe"); });
                            $(".stripeMe table:not(td > table) > tbody > tr:not('.spe'):even").addClass("alt");
                            PageFun(p, $("total", data).text());
                        }
                    }
                }
            });
        }

        function openEditBlock() {
            $.fancybox({
                href: "#editblock",
                title: "",
                closeBtn: false,
                minWidth: "700",
                minHeight: "380",
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
                        locked: false,   //開始fancybox時,背景是否回top
                        closeClick: false //點背景關閉 fancybox
                    }
                },
                afterClose: function () {
                    $("#ddlCompetence").prop("disabled", false);
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
                        var ddlstr = '<option value="">--請選擇--</option>';
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

        function openManagerList() {
            if ($("#ddlCity").val() == "") {
                alert("請選擇執行機關");
                return;
            }

            $.colorbox({
                iframe: true,
                href: "ManagerList.aspx?city=" + $("#ddlCity").val(),
                width: 800,
                height: 630,
                overlayClose: false, //點背景關閉 colorbox
                onComplete: function () {
                    $('#cboxTitle').hide();
                    $('#cboxClose').remove();
                }
            });
        }

        function returnManagerValue(jv) {
            var jsonval = $.parseJSON(jv)
            $("#M_Manager_Name").html(jsonval.name);
            $("#M_Manager_ID").val(jsonval.guid);
        }
        
        var listNum = 10; //每頁顯示個數,計算用
        var PagesLen; //總頁數 
        var PageNum = 4; //下方顯示分頁數(PageNum+1個)
        function PageFun(PageNow, TotalItem) {
            //Math.ceil -> 無條件進位
            PagesLen = Math.ceil(TotalItem / listNum);
            if (PagesLen <= 1)
                $("#changpage").hide();
            else {
                $("#changpage").show();
                upPage(PageNow, TotalItem);
            }
        }
    </script>
    <style>
        a.asc:after {
            content: attr(data-content) '▲';
        }
        a.desc:after {
            content: attr(data-content) '▼';
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <input type="hidden" id="idtmp" />
        <input type="hidden" id="sortMethod" name="sortMethod" value="desc" />
        <input type="hidden" id="sortName" name="sortName" value="M_CreateDate" />
    <div class="container">
        <div class="twocol filetitlewrapper">
	        <div class="left"><span class="filetitle font-size5">成員管理</span></div><!-- left -->
            <div class="right"></div><!-- right -->
        </div><!-- twocol -->

        <div class="twocol" style="margin-top:20px;">
	        <div class="left">關鍵字：<input type="text" id="SearchStr" class="inputex" />&nbsp;<input type="button" value="查詢" class="genbtn" onclick="getData(0);" /></div><!-- left -->
            <div class="right"><input type="button" id="addbtn" value="新增" class="genbtn" /></div><!-- right -->
        </div><!-- twocol -->
        <br />
        <div class="stripeMe margin5T font-normal">
            <table id="tablist" width="100%" border="0" cellspacing="0" cellpadding="0">
                <thead>
                    <tr>
                        <th nowrap="nowrap" style="width:40px;">項次</th>
                        <th nowrap="nowrap"><a href="javascript:void(0);" name="sortbtn" atp="M_Name">姓名</a></th>
                        <th nowrap="nowrap"><a href="javascript:void(0);" name="sortbtn" atp="M_Email">E-Mail</a></th>
                        <th nowrap="nowrap"><a href="javascript:void(0);" name="sortbtn" atp="M_City">執行機關</a></th>
                        <th nowrap="nowrap"><a href="javascript:void(0);" name="sortbtn" atp="M_Office">承辦局處</a></th>
                        <th nowrap="nowrap"><a href="javascript:void(0);" name="sortbtn" atp="M_Competence">身份</a></th>
                        <th nowrap="nowrap"><a href="javascript:void(0);" name="sortbtn" atp="M_CreateDate">建立日期</th>
                        <th nowrap="nowrap" style="width:150px;">動作</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
            <div id="changpage" class="margin20T textcenter"></div>
        </div>
    </div>

    <div id="editblock" style="display:none;">
        <div class="font-bold" style="margin-bottom:10px;">成員管理 > <span id="actionStr"></span></div>
        <div class="stripeMe">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <th><span style="color:red;">*</span>姓名</th>
                    <td><input type="text" id="M_Name" class="inputex width100 str" maxlength="50" /></td>
                    <th><span style="color:red;">*</span>職稱</th>
                    <td><input type="text" id="M_JobTitle" class="inputex width100 str" maxlength="50" /></td>
                </tr>
                <tr>
                    <th><span style="color:red;">*</span>E-Mail</th>
                    <td colspan="3">
                        <input type="text" id="M_Email" class="inputex width100 str" maxlength="200" /><br />
                        <span class="SpnAPR"><input type="checkbox" id="aprEmail" name="aprEmail" style="margin-top:5px;" />&nbsp;帳號密碼同E-Mail</span>
                    </td>
                </tr>
                <tr class="aptr">
                    <th><span style="color:red;">*</span>帳號</th>
                    <td><input type="text" id="M_Account" class="inputex width100 str" maxlength="50" /></td>
                    <th><span style="color:red;">*</span>密碼</th>
                    <td><input type="password" id="M_Pwd" class="inputex width100 str" maxlength="50" /><input type="hidden" id="Old_Pwd" class="str" /></td>
                </tr>
                <tr>
                    <th><span style="color:red;">*</span>電話/分機</th>
                    <td colspan="3">
                        <input type="text" id="M_Tel" class="inputex width40 str num" maxlength="20" />&nbsp;-&nbsp;<input type="text" id="M_Ext" class="inputex width15 str num" maxlength="10" /><br />
                        電話格式 Ex：03-7654321
                    </td>
                </tr>
                <tr>
                    <th>傳真</th>
                    <td><input type="text" id="M_Fax" class="inputex width100 str num" maxlength="20" /></td>
                    <th>手機</th>
                    <td><input type="text" id="M_Phone" class="inputex width100 str num" maxlength="20" /></td>
                </tr>
                <tr>
                    <th>地址</th>
                    <td colspan="3"><input type="text" id="M_Addr" class="inputex width100 str" maxlength="200" /></td>
                </tr>
                <tr>
                    <th><span style="color:red;">*</span>執行機關</th>
                    <td><select id="ddlCity" class="str"></select></td>
                    <th><span style="color:red;">*</span>承辦局處</th>
                    <td><input type="text" id="M_Office" class="inputex width100 str" /></td>
                </tr>
                <tr>
                    <th><span style="color:red;">*</span>權限/身份</th>
                    <td id="ddlcompTd" colspan="3"><select id="ddlCompetence" class="str"></select></td>
                    <th class="mar" style="display:none;"><span style="color:red;">*</span>承辦主管</th>
                    <td class="mar" style="display:none;">
                        <a href="javascript:openManagerList()"><img src="../App_Themes/images/btn-search.gif" /></a>
                        <span id="M_Manager_Name" style="color:blue;"></span>
                        <input type="hidden" id="M_Manager_ID" class="str" />
                    </td>
                </tr>
            </table>
        </div>
        <div style="text-align:right; margin-top:10px;">
            <input type="button" id="subbtn" class="genbtn" value="送出" />
            <input type="button" id="savebtn" class="genbtn" value="儲存" style="display:none;" />
            <input type="button" id="cancelbtn" class="genbtn" value="取消" />
        </div>
    </div>
</asp:Content>

