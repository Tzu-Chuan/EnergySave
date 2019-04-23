<%@ Page Title="" Language="C#" MasterPageFile="~/Manage/Admin.master" AutoEventWireup="true" CodeFile="ChangeContractor.aspx.cs" Inherits="Manage_ChangeContractor" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            getddl("02", "#ddlcity");

            $(document).on("change", "#ddlcity", function () {
                $("#PerName").html("");
                $("#PersonGuid").val("");
                $("#nPerName").html("");
                $("#nPersonGuid").val("");
                $("#msgblock").html("");
            });

            $(document).on("click", "#subbtn", function () {
                var msg = "";
                if ($("#ddlcity").val() == "")
                    msg += "請選擇【承辦人執行機關】\n";
                if ($("#PersonGuid").val() == "")
                    msg += "請選擇【原承辦人】\n";
                if ($("#nPersonGuid").val() == "")
                    msg += "請選擇【新承辦人】\n";

                if (msg != "") {
                    alert(msg);
                    return;
                }


                $.ajax({
                    type: "POST",
                    async: false, //在沒有返回值之前,不會執行下一步動作
                    url: "../handler/updateConntractor.ashx",
                    data: {
                        orgid: $("#PersonGuid").val(),
                        newid: $("#nPersonGuid").val()
                    },
                    error: function (xhr) {
                        alert("Error " + xhr.status);
                        console.log(xhr.responseText);
                    },
                    success: function (data) {
                        if (data.indexOf("Error") > -1)
                            alert(data);
                        else if (data = "succeed") {
                            var str = "承辦人已更新<br>";
                            str += "原承辦人的計畫基本資料、月報、季報...等資料權限將轉移至新承辦人<br>";
                            str += "原承辦人無法再修改原有資料";
                            $("#msgblock").html(str);
                            $(".keyin").val("");
                            $(".keyspn").html("");
                        }
                    }
                });
            });
        });

        function openManagerList(type) {
            if ($("#ddlcity").val() == "") {
                alert("請選擇執行機關");
                return;
            }
            
            $.colorbox({
                iframe: true,
                href: "ContractorList.aspx?city=" + $("#ddlcity").val() + "&tp=" + type,
                width: 800,
                height: 630,
                overlayClose: false, //點背景關閉 colorbox
                onComplete: function () {
                    $('#cboxTitle').hide();
                    $('#cboxClose').remove();
                }
            });
        }

        function returnPersonValue(jv) {
            var jsonval = $.parseJSON(jv)
            $("#PerName").html(jsonval.name);
            $("#PersonGuid").val(jsonval.guid);
        }

        function returnNewPersonValue(jv) {
            var jsonval = $.parseJSON(jv)
            $("#nPerName").html(jsonval.name);
            $("#nPersonGuid").val(jsonval.guid);
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
    <div class="container">
        <div class="twocol filetitlewrapper">
	        <div class="left"><span class="filetitle font-size5">更換承辦人</span></div><!-- left -->
            <div class="right"></div><!-- right -->
        </div><!-- twocol -->

        <div class="margin20T">
            <div class="margin10T">承辦人執行機關：<select class="keyin" id="ddlcity"></select></div>
            <div id="orgpanel" class="margin10T">
                原承辦人：<a href="javascript:openManagerList('org')"><img src="../App_Themes/images/btn-search.gif" /></a>
                <span class="keyspn" id="PerName" style="color:blue;"></span><input class="keyin" type="hidden" id="PersonGuid" name="PersonGuid" />
            </div>
            <div id="newpanel" class="margin10T">
                新承辦人：<a href="javascript:openManagerList('new')"><img src="../App_Themes/images/btn-search.gif" /></a>
                <span class="keyspn" id="nPerName" style="color:blue;"></span><input class="keyin" type="hidden" id="nPersonGuid" name="nPersonGuid" />
            </div>
        </div>
        <div id="msgblock" class="margin20T" style="color:red;"></div>
        <div class="margin20T"><input type="button" id="subbtn" class="genbtn" value="送出" /></div>
    </div>
</asp:Content>

