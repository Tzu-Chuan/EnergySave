<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="News.aspx.cs" Inherits="WebPage_News" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            getData(0);

            //datepicker
            $("#SearchDate").datepicker({
                changeMonth: true,
                changeYear: true,
                dateFormat: 'yy/mm/dd',
                dayNamesMin: ["日", "一", "二", "三", "四", "五", "六"],
                monthNamesShort: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"],
                yearRange: '-100:+100'
            });
            
            //下載檔案
            $(document).on("click", "a[name='downloadbtn']", function () {
                var id = $(this).attr("fid");
                location.href = "../DOWNLOAD.aspx?v=" + id;
            });
        });// end js

        //撈公告列表
        function getData(p) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getNews.aspx",
                data: {
                    CurrentPage: p,
                    SearchDate: $("#SearchDate").val(),
                    SearchStr: $("#SearchStr").val()
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
                            $("#div_list").empty();
                            var tabstr = '';
                            tabstr += '<div id="datalist" style="background-color:#ecfae0;border-radius:20px;padding:5px 1px 5px 1px;">';
                            //tabstr += '<div style="font-size:1.61em;font-color:#649016;margin:10px 15px;">公告</div>';
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function (i) {
                                    tabstr += '<div>';
                                    tabstr += '<div class="font-size3" style="margin:15px;padding:15px;background-color:#FFFFFF;border-radius:10px;word-break:break-all;">';
                                    tabstr += '<span>' + $(this).children("N_Date").text().trim() + '</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;';
                                    tabstr += '<span style="text-indent:4em;"><a href = "NewsDetail.aspx?v=' + $(this).children("N_ID").text().trim() + '" target="_blank">' + $(this).children("N_Title").text().trim() + '</a></span>';
                                    tabstr += '</div></div>';
                                });
                            }
                            else
                                tabstr += '<div style="background-color:#ecfae0;border-radius:20px;padding:2px;"><div style="margin:10px">查詢無資料</div></div>';
                            tabstr += '</div>';
                            $("#div_list").append(tabstr);
                            //$(".stripeMe tr").mouseover(function () { $(this).addClass("spe"); }).mouseout(function () { $(this).removeClass("spe"); });
                            //$(".stripeMe table:not(td > table) > tbody > tr:not('.spe'):even").addClass("alt");
                            CreatePage(p, $("total", data).text());
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
            <div class="left"><span class="filetitle font-size5">最新消息</span></div>
            <div class="right">最新消息</div><!-- right -->
        </div>

        <div class="twocol" style="margin-top: 20px;">
            <div class="left">
                發布日期：<input type="text" id="SearchDate" class="inputex" />&nbsp;
                關鍵字：<input type="text" id="SearchStr" class="inputex" />&nbsp;
                <input type="button" value="查詢" class="genbtn" onclick="getData(0);" />
            </div>
            <!-- left -->
        </div>
        <!-- twocol -->
        <br />
        <div class="stripeMe margin5T font-normal tabcss">
            <div id="div_list"></div>
            <div id="changpage" class="margin20T textcenter"></div>
        </div>
    </div>

</asp:Content>

