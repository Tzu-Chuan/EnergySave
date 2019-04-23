<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ManagerList.aspx.cs" Inherits="Manage_ManagerList" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1" />
<meta name="keywords" content="關鍵字內容" />
<meta name="description" content="描述" /><!--告訴搜尋引擎這篇網頁的內容或摘要。--> 
<meta name="generator" content="Notepad" /><!--告訴搜尋引擎這篇網頁是用什麼軟體製作的。--> 
<meta name="author" content="工研院 資科中心" /><!--告訴搜尋引擎這篇網頁是由誰製作的。--> 
<meta name="copyright" content="本網頁著作權所有" /><!--告訴搜尋引擎這篇網頁是...... --> 
<meta name="revisit-after" content="3 days" /><!--告訴搜尋引擎3天之後再來一次這篇網頁，也許要重新登錄。-->
<title>承辦主管清單</title>
<link href="<%= ResolveUrl("~/App_Themes/css/OchiLayout.css") %>" rel="stylesheet" type="text/css" /><!-- ochsion layout base -->
<link href="<%= ResolveUrl("~/App_Themes/css/OchiColor.css") %>" rel="stylesheet" type="text/css" /><!-- ochsion layout color -->
<link href="<%= ResolveUrl("~/App_Themes/css/jquery-ui.css") %>" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="<%= ResolveUrl("~/js/jquery-1.11.2.min.js") %>"></script>
<script type="text/javascript" src="<%= ResolveUrl("~/js/jquery-ui-1.10.2.custom.min.js") %>"></script>
<script type="text/javascript" src="<%= ResolveUrl("~/js/downfile.js") %>"></script>
<script type="text/javascript" src="<%= ResolveUrl("~/js/Pager.js") %>"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            if ($.getParamValue('str') != "")
                $("#SearchStr").val($.getParamValue('str'));

            getData(0);

            //click return value
            $(document).on("click", "#tablist tbody tr td", function (e) {
                var jsonval = new Object();
                jsonval.id = $(this).closest('tr').attr("aid");
                jsonval.guid = $(this).closest('tr').attr("aguid");
                jsonval.name = $(this).closest('tr').attr("aname");
                jsonval = JSON.stringify(jsonval);
                parent.returnManagerValue(jsonval);
                parent.$.colorbox.close();
            });
        });

        function getData(p) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getManagerList.ashx",
                data: {
                    CurrentPage: p,
                    SearchStr: $("#SearchStr").val(),
                    City: $.getParamValue('city')
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if (data != null) {
                        data = $.parseXML(data);
                        $("#tablist tbody").empty();
                        var tabstr = '';
                        if ($(data).find("data_item").length > 0) {
                            $(data).find("data_item").each(function (i) {
                                tabstr += '<tr aid="' + $(this).children("M_ID").text().trim() + '" aguid="' + $(this).children("M_Guid").text().trim() +'" aname="' + $(this).children("M_Name").text().trim() + '">';
                                tabstr += '<td align="center" nowrap="nowrap" style="cursor: pointer;">' + $(this).children("M_Name").text().trim() + '</td>';
                                tabstr += '<td align="center" nowrap="nowrap" style="cursor: pointer;">' + $(this).children("M_JobTitle").text().trim() + '</td>';
                                tabstr += '<td align="center" nowrap="nowrap" style="cursor: pointer;">' + $(this).children("M_Email").text().trim() + '</td>';
                                tabstr += '<td align="center" nowrap="nowrap" style="cursor: pointer;">' + $(this).children("City").text().trim() + '</td>';
                                tabstr += '<td align="center" nowrap="nowrap" style="cursor: pointer;">' + $(this).children("M_Office").text().trim() + '</td>';
                                tabstr += '</tr>';
                            });
                        }
                        else
                            tabstr += "<tr><td colspan='5'>查詢無資料</td></tr>";
                        $("#tablist tbody").append(tabstr);
                        $(".stripeMe tr").mouseover(function () { $(this).addClass("spe"); }).mouseout(function () { $(this).removeClass("spe"); });
                        $(".stripeMe table:not(td > table) > tbody > tr:not('.spe'):even").addClass("alt");
                        PageFun(p, $("total", data).text());
                    }
                }
            });
        }

        var nowPage = 0; //當前頁
        var listNum = 10; //每頁顯示個數
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
    .indpblock{border:1px solid #CCC; padding:5px; border-radius:5px;}
    a.asc:after {content: attr(data-content) '▲';}
    a.desc:after {content: attr(data-content) '▼';}
</style>
</head>
<body>
    <form id="form1" runat="server">
    <div class="indpblock margin10T">
        <div class="gentable">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <td align="right" class="titlebg" width="100">
                        <div class="font-title titlebackicon">關鍵字</div>
                    </td>
                    <td>
                        <input id="SearchStr" type="text" class="inputex width50" />&nbsp;(格式：標題欄位)</td>
                    <td align="right">
                        <a id="SearchBtn" href="javascript:getData(0);" class="genbtnS">查詢</a>
                        <a href="javascript:parent.$.colorbox.close();" class="genbtnS">離開</a>
                    </td>
                </tr>
            </table>
        </div><!-- gentable -->
    </div><!-- indpblock -->
    <div class="stripeMe margin5T font-normal" style="margin-top:20px;">
        <table id="tablist" width="100%" border="0" cellspacing="0" cellpadding="0">
            <thead>
                <tr>
                    <th nowrap="nowrap">姓名</th>
                    <th nowrap="nowrap">職稱</th>
                    <th nowrap="nowrap">E-mail</th>
                    <th nowrap="nowrap">執行機關</th>
                    <th nowrap="nowrap">承辦局處</th>
                </tr>
            </thead>
            <tbody></tbody>
        </table>
    </div>
    <div id="changpage" class="margin20T textcenter"></div>
    </form>
</body>
</html>