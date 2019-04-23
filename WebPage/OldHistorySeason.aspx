<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="OldHistorySeason.aspx.cs" Inherits="WebPage_OldHistorySeason" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">舊版季報歷史資料</span></div>
            <div class="right">歷史資料 / 舊版季報</div>
        </div>
        <div>
            <ul>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/基隆市_107年第2季季報.pdf") %>">基隆市_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/台北市_107年第2季季報.pdf") %>">台北市_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/新北市_107年第2季季報.pdf") %>">新北市_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/桃園市_107年第2季季報.pdf") %>">桃園市_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/新竹市_107年第2季季報.pdf") %>">新竹市_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/新竹縣_107年第2季季報.pdf") %>">新竹縣_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/苗栗縣_107年第2季季報.pdf") %>">苗栗縣_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/臺中市_107年第2季季報.pdf") %>">臺中市_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/彰化縣_107年第2季季報.pdf") %>">彰化縣_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/南投縣_107年第2季季報.pdf") %>">南投縣_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/雲林縣_107年第2季季報.pdf") %>">雲林縣_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/嘉義市_107年第2季季報.pdf") %>">嘉義市_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/嘉義縣_107年第2季季報.pdf") %>">嘉義縣_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/臺南市_107年第2季季報.pdf") %>">臺南市_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/高雄市_107年第2季季報.pdf") %>">高雄市_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/屏東縣_107年第2季季報.pdf") %>">屏東縣_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/臺東縣_107年第2季季報.pdf") %>">臺東縣_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/花蓮縣_107年第2季季報.pdf") %>">花蓮縣_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/宜蘭縣_107年第2季季報.pdf") %>">宜蘭縣_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/澎湖縣_107年第2季季報.pdf") %>">澎湖縣_107年第2季季報</a></li>
                <li><a href="<%= ResolveUrl("~/OldSeasonPDF/金門縣_107年第2季季報.pdf") %>">金門縣_107年第2季季報</a></li>
            </ul>                                                                      
        </div>
    </div>
</asp:Content>

