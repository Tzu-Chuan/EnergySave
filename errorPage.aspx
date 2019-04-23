<%@ Page Language="C#" AutoEventWireup="true" CodeFile="errorPage.aspx.cs" Inherits="errorPage" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
     <div>
     <fieldset>
            <legend class="font_black font_size15">錯誤訊息</legend>
            <table>
                <tr>
                    <td class="font_black font_size15">
                        <br />
                        <asp:Label ID="Label1" runat="server"></asp:Label>                   
                    </td>
                </tr>
                <tr>
                    <td class="font_size13">
                    <a href="<%=ResolveUrl("~/WebPage/Login.aspx")%>" target="_self">返回首頁</a>
                    </td>
                </tr>
             </table>
        </fieldset>
    </div>
    </form>
</body>
</html>
