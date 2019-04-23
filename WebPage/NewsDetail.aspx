<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="NewsDetail.aspx.cs" Inherits="WebPage_NewsDetail" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            getDataByID();
            //下載檔案
            $(document).on("click", "a[name='downloadbtn']", function () {
                var id = $(this).attr("fid");
                location.href = "../DOWNLOAD.aspx?v=" + id;
            });
        });// end js

        //撈修改那筆公告
        function getDataByID() {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getNews_ById.aspx",
                data: {
                    N_ID: $.getParamValue('v')
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if ($(data).find("Error").length > 0)
                        alert(data);
                    else {
                        if (data != null) {
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function (i) {
                                    $("#strTitle").html($(this).children("N_Title").text().trim());
                                    $("#strDate").html($(this).children("N_Date").text().trim());
                                    $("#strContent").html($(this).children("N_Content").text().trim());
                                    getFileList("06", $(this).children("N_Guid").text().trim());
                                });
                            }
                        }
                    }
                }
            });
        }

        function getFileList(type,nGuid) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/GetFileList.aspx",
                data: {
                    pGuid: nGuid,
                    type: type
                },
                error: function (xhr) {
                   alert(xhr.responseText);
                },
                success: function (data) {
                     if ($(data).find("Error").length > 0) {
                         alert($(data).find("Error").attr("Message"));
                     }
                     else {
                        if (data != null) {
                            $("#divFile").empty();
                            var divstr = '附件下載<br />';
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function (i) {
                                    divstr += '<div style="margin-top:5px;">'
                                    divstr += '<a href="javascript:void(0);" name="downloadbtn" fid="' + $(this).children("download_id").text().trim() + '">'
                                    divstr += $(this).children("file_orgname").text().trim() + $(this).children("file_exten").text().trim() + '</a><br />';
                                    divstr += '</div>'
                                });
                                divstr += '</tbody>';
                                $("#divFile").append(divstr);
                            }
                        }
                    }
                }
            });
        }

        //分頁設定
        //ListNum: 每頁顯示資料筆數
        //PageNum: 分頁頁籤顯示數
        PageOption.Selector = "#changpage";
        PageOption.ListNum = 10;
        PageOption.PageNum = 10;
        PageOption.PrevStep = false;
        PageOption.NextStep = false;
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left" style="word-break:break-all;"><b><span id="strTitle" class="font-size3"></span></b></div>
            <div class="right"><span id="strDate"></span></div><!-- right -->
        </div>

        <!-- twocol -->
        <br />
        <div id="strContent"></div>
        <div id="divFile" style="margin-top:20px;"></div>
    </div>

</asp:Content>

