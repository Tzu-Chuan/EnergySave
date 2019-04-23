<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="PersonalFile.aspx.cs" Inherits="WebPage_PersonalFile" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            getddl("02", "#ddlCity");
            getData();

            //儲存 button
            $(document).on("click", "#savebtn", function () {
                var msg = "";
                if ($("#M_Name").val() == "")
                    msg += "請輸入【姓名】\n";
                if ($("#M_JobTitle").val() == "")
                    msg += "請輸入【職稱】\n";
                if ($("#M_Email").val() == "")
                    msg += "請輸入【E-Mail】\n";
                if ($("#M_Email").val().trim() != "") {
                    var pattern = new RegExp(/^[+a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/);
                    if (!pattern.test($("#M_Email").val()))
                        msg += "請輸入正確Email\n";
                }
                if (!$("#aprEmail").is(":checked")) {
                    if ($("#M_Account").val() == "")
                        msg += "請輸入【帳號】\n";
                    if ($("#M_Pwd").val() == "")
                        msg += "請輸入【密碼】\n";
                }
                if ($("#ddlCity").val() == "")
                    msg += "請選擇【執行機關】\n";
                if ($("#M_Office").val() == "")
                    msg += "請輸入【承辦局處】\n";

                if (msg != "") {
                    alert(msg);
                    return false;
                }
                
                $.ajax({
                    type: "POST",
                    async: false, //在沒有返回值之前,不會執行下一步動作
                    url: "../handler/addMember.ashx",
                    data: {
                        id: $.getParamValue('v'),
                        mode: "pFile",
                        M_Name: $("#M_Name").val(),
                        M_JobTitle: $("#M_JobTitle").val(),
                        M_Tel: $("#M_Tel").val(),
                        M_Ext: $("#M_Ext").val(),
                        M_Fax: $("#M_Fax").val(),
                        M_Phone: $("#M_Phone").val(),
                        M_Email: $("#M_Email").val(),
                        M_Account: $("#M_Account").val(),
                        M_Pwd: $("#M_Pwd").val(),
                        Old_Pwd: $("#Old_Pwd").val(),
                        M_Addr: $("#M_Addr").val(),
                        M_City: $("#ddlCity").val(),
                        M_Office: $("#M_Office").val()
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
                                return;
                            }

                            if (data == "reLogin") {
                                alert("請重新登入");
                                window.location = "../WebPage/Login.aspx";
                                return;
                            }

                            if (data == "succeed") {
                                alert("儲存完成");
                                getData();
                            }
                        }
                    }
                });
            });
        });

        function getData() {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/editMember.ashx",
                data: {
                    mode: "pFile"
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
                                        $("#compTd").attr("colspan", "");
                                        $("#M_Comp").html("承辦人");
                                    }
                                    else {
                                        $(".mar").hide();
                                        $("#compTd").attr("colspan", "3");
                                        $("#M_Comp").html("承辦主管");
                                    }
                                    
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
                                    $("#M_Manager_Name").html($(this).children("Manager").text().trim());
                                });
                            }
                        }
                    }
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
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div class="container stripeMe">
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
                <th>權限/身份</th>
                <td id="compTd" colspan="3"><span id="M_Comp"></span></td>
                <th class="mar" style="display:none;">承辦主管</th>
                <td class="mar" style="display:none;"><span id="M_Manager_Name"></span></td>
            </tr>
        </table>
        <div style="text-align:right; margin-top:10px;">
            <input type="button" id="savebtn" class="genbtn" value="儲存" />
        </div>
    </div>
</asp:Content>

