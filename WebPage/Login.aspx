<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Login.aspx.cs" Inherits="WebPage_Login" %>

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
<link href="<%= ResolveUrl("~/App_Themes/css/bootstrap.css") %>" rel="stylesheet" /><!-- normalize & bootstrap's grid system -->
<link href="<%= ResolveUrl("~/App_Themes/css/font-awesome.min.css") %>" rel="stylesheet" type="text/css" /><!-- CSS icon -->
<link href="<%= ResolveUrl("~/App_Themes/css/superfish.css") %>" rel="stylesheet" type="text/css" /><!-- 下拉選單 -->
<link href="<%= ResolveUrl("~/App_Themes/css/jquery.mmenu.css") %>" rel="stylesheet" type="text/css" /><!-- mmenu css:行動裝置選單 -->
<link href="<%= ResolveUrl("~/App_Themes/css/colorbox.css") %>" rel="stylesheet" type="text/css" />
<link href="<%= ResolveUrl("~/App_Themes/css/jquery.powertip.css") %>" rel="stylesheet" type="text/css" />
<link href="<%= ResolveUrl("~/App_Themes/css/jquery.datetimepicker.css") %>" rel="stylesheet" type="text/css" /><!-- datepicker -->
<link href="<%= ResolveUrl("~/App_Themes/css/OchiLayout.css") %>" rel="stylesheet" type="text/css" /><!-- ochsion layout base -->
<link href="<%= ResolveUrl("~/App_Themes/css/OchiColor.css") %>" rel="stylesheet" type="text/css" /><!-- ochsion layout color -->
<link href="<%= ResolveUrl("~/App_Themes/css/OchiRWD.css") %>" rel="stylesheet" type="text/css" /><!-- ochsion layout RWD -->
<link href="<%= ResolveUrl("~/App_Themes/css/style.css") %>" rel="stylesheet" type="text/css" />
<link href="<%= ResolveUrl("~/App_Themes/css/jquery-ui.css") %>" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="<%= ResolveUrl("~/js/jquery-1.11.2.min.js") %>"></script>
<script type="text/javascript" src="<%= ResolveUrl("~/js/jquery-ui-1.10.2.custom.min.js") %>"></script>
<script type="text/javascript" src="<%= ResolveUrl("~/js/jquery.breakpoint-min.js") %>"></script><!-- 斷點設定 -->
<script type="text/javascript" src="<%= ResolveUrl("~/js/superfish.min.js") %>"></script><!-- 桌機版下拉選單 -->
<script type="text/javascript" src="<%= ResolveUrl("~/js/supposition.js") %>"></script><!-- 下拉選單:修正最後項在視窗大小不夠時的BUG -->
<script type="text/javascript" src="<%= ResolveUrl("~/js/jquery.mmenu.min.js") %>"></script><!-- mmenu js:行動裝置選單 -->
<script type="text/javascript" src="<%= ResolveUrl("~/js/jquery.touchSwipe.min.js") %>"></script><!-- 增加JS觸控操作 for mmenu -->
<script type="text/javascript" src="<%= ResolveUrl("~/js/jquery.colorbox.js") %>"></script>
<script type="text/javascript" src="<%= ResolveUrl("~/js/jquery.easytabs.min.js") %>"></script><!-- easytabs tab -->
<script type="text/javascript" src="<%= ResolveUrl("~/js/jquery.powertip.min.js") %>"></script><!-- powertip:tooltips -->
<script type="text/javascript" src="<%= ResolveUrl("~/js/jquery.datetimepicker.js") %>"></script><!-- datepicker -->
<script type="text/javascript" src="<%= ResolveUrl("~/js/restables.min.js") %>"></script><!-- RWD table -->
<script type="text/javascript" src="<%= ResolveUrl("~/js/downfile.js") %>"></script>
<script type="text/javascript" src="<%= ResolveUrl("~/js/Pager.js") %>"></script>
<script type="text/javascript" src="<%= ResolveUrl("~/js/jquery.blockUI.js") %>"></script>
<title>縣市共推住商節電行動計畫填報平台</title>
<script type="text/javascript">
    $(document).ready(function () {
        $(document).on("keyup", "body", function (e) {
            if (e.keyCode == 13)
                $("#lgbtn").click();
        });

        $(document).on("click", "#lgbtn", function () {
            $.ajax({
                type: "POST",
                async: true, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/Login.ashx",
                data: {
                    acc: $("#accStr").val(),
                    pw: $("#pwStr").val(),
                    vCode: $("#codeStr").val()
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if (data.indexOf("Error") > -1)
                        alert(data);
                    else {
                        if (data == "Failed")
                            alert("帳號密碼有誤");
                        else if (data == "CodeFailed")
                            alert("驗証碼錯誤");
                        else
                            location.href = "News.aspx";
                    }
                }
            });
        });
    });

    function changeCode() {
        $("#imgCode").attr("src", "../handler/ValidateNumber.ashx?" + Math.random());
    }
</script>
</head>
<body>
    <form id="form1" runat="server">
    <div class="WrapperBody">
        <div class="WrapperHeader container">
            <div class="logo"><a href="#"><img src="<%= ResolveUrl("~/App_Themes/images/logo.png") %>" /></a></div>
        </div>

        <div class="container">
            <div style="text-align:center; font-size:16pt; margin-top:50px;">登入系統</div>
            <div class="stripeMe" style="margin-top:20px;">
                <table align="center">
                    <tr>
                        <th align="right">帳號：</th>
                        <td colspan="2"><input type="text" id="accStr" style="width:99%" /></td>
                    </tr>
                    <tr>
                        <th align="right">密碼：</th>
                        <td colspan="2"><input type="password" id="pwStr" style="width:99%" /></td>
                    </tr>
                    <tr>
                        <th align="right">驗証碼：</th>
                        <td><input type="text" id="codeStr" maxlength="4" />&nbsp;<a href="javascript:void(0);" onclick="changeCode()">變更驗証碼</a></td>
                        <td><img src="../handler/ValidateNumber.ashx" alt="驗證碼" id="imgCode" height="30" /></td>
                    </tr>
                </table>
            </div>
            <div style="text-align:center; margin-top:10px;">
                <input type="button" id="lgbtn" value="登入" class="genbtn" />
            </div>
        </div>
    </div>

    <div class="WrapperFooter">
        <div class="footerblock container font-size2 font-normal lineheight02">
            版權所有©2018 工業技術研究院｜ 建議瀏覽解析度1024x768以上<br />
            業務窗口：陳奕宏(電話:03-5917242)
        </div><!--{* footerblock *}-->
    </div>
    </form>
<script type="text/javascript" src="<%= ResolveUrl("~/js/GenCommon.js") %>"></script><!-- UIcolor JS -->
<script type="text/javascript" src="<%= ResolveUrl("~/js/PageCommon.js") %>"></script><!-- 系統共用 JS -->
<script type="text/javascript" src="<%= ResolveUrl("~/js/autoHeight.js") %>"></script><!-- 高度不足頁面的絕對置底footer -->
</body>
</html>
