<%@ Page Title="" Language="C#" MasterPageFile="~/Manage/Admin.master" AutoEventWireup="true" CodeFile="NewsManage.aspx.cs" Inherits="Manage_NewsManage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            getData(0);

            //datepicker
            $("#SearchDate,#N_Date").datepicker({
                changeMonth: true,
                changeYear: true,
                dateFormat: 'yy/mm/dd',
                dayNamesMin: ["日", "一", "二", "三", "四", "五", "六"],
                monthNamesShort: ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"],
                yearRange: '-100:+100'
            });

            var mod = "";
            //新增/修改 送出 button
            $(document).on("click", "#addbtn,input[name='modbtn']", function () {
                $("#tmpGuid").val("");
                $(".mfile").hide();
                $("#N_Date").val("");
                $("#N_Title").val("");
                $("#N_Content").val("");
                $("#modid").val("");
                mod = "add";

                if (this.id == "addbtn") {
                    $("#tmpGuid").val(newGuid().replace(/-/g, ""));
                    openEditBlock("新增");
                }
                else {
                    mod = "mod";
                    getDataByID($(this).attr("aid"));
                    getFileList("06", $(this).attr("agid"));
                    $("#tmpGuid").val($(this).attr("agid"));
                    openEditBlock("編輯");
                }
            });

            //取消 button
            $(document).on("click", "#cancelbtn", function () {
                if (confirm('資料尚未儲存，確定取消？'))
                    $("#editblock").dialog("close");
            });

            //刪除按鈕
            $(document).on("click", "input[name='delbtn']", function () {
                if (confirm('是否確定刪除該筆公告？'))
                    delData($(this).attr("aid"));
            });

            //新增&修改
            $(document).on("click", "#subbtn", function () {
                var msg = "";
                if ($("#N_Date").val() == "")
                    msg += "請輸入【發布日期】\n";
                if ($("#N_Title").val() == "")
                    msg += "請輸入【標題】\n";
                if ($("#N_Content").val() == "")
                    msg += "請輸入【內容】\n";


                if (msg != "") {
                    alert(msg);
                    return false;
                }

                var form = document.createElement('form');
                var iframe = document.createElement('iframe');
                var N_Date = document.createElement('input');
                var N_Title = document.createElement('input');
                var N_Content = document.createElement('input');
                var N_ID = document.createElement('input');
                var N_Mod = document.createElement('input');
                var N_Guid = $('<input type="hidden" id="N_Guid" name="N_Guid" value="' + $("#tmpGuid").val() + '" />');

                iframe.setAttribute("name", "postiframe");
                iframe.setAttribute("id", "postiframe");
                iframe.setAttribute("style", "display: none");

                N_Date.setAttribute("id", "N_Date");
                N_Date.setAttribute("name", "N_Date");
                N_Date.setAttribute("value", $("#N_Date").val());

                N_Title.setAttribute("id", "N_Title");
                N_Title.setAttribute("name", "N_Title");
                N_Title.setAttribute("value", $("#N_Title").val());

                N_Content.setAttribute("id", "N_Content");
                N_Content.setAttribute("name", "N_Content");
                N_Content.setAttribute("value", encodeURIComponent($("#N_Content").val()));

                N_ID.setAttribute("id", "N_ID");
                N_ID.setAttribute("name", "N_ID");
                N_ID.setAttribute("value", $("#modid").val());

                N_Mod.setAttribute("id", "N_Mod");
                N_Mod.setAttribute("name", "N_Mod");
                N_Mod.setAttribute("value", mod);

                
                $("input[name='N_Guid']").remove();

                form.appendChild(iframe);
                //form.appendChild(file.get(0));
                form.appendChild(N_Date);
                form.appendChild(N_Title);
                form.appendChild(N_Content);
                form.appendChild(N_ID);
                form.appendChild(N_Mod);
                form.appendChild(N_Guid[0]);
                $("body").append(form);

                form.setAttribute("action", "../handler/modNews.ashx");
                form.setAttribute("method", "post");
                form.setAttribute("enctype", "multipart/form-data");
                form.setAttribute("encoding", "multipart/form-data");
                form.setAttribute("target", "postiframe");
                form.setAttribute("style", "display: none");
                form.submit();
            });

            //上傳 button
            $(document).on("click", "#FileBtn", function () {
                $.fancybox({
                    href: "../WebPage/File_Upload.aspx?v=" + $("#tmpGuid").val() + "&tp=06",
                    title: "",
                    closeBtn: false,
                    type: "iframe",
                    minWidth: "800",
                    minHeight: $(window).height() - 200,
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
                    }
                });
            });

            //刪除檔案
            $(document).on("click", "a[name='delfilebtn']", function () {
                if (confirm("資料刪除後將無法復原，確定刪除?")) {
                    $.ajax({
                        type: "POST",
                        async: false, //在沒有返回值之前,不會執行下一步動作
                        url: "../handler/deleteFile.ashx",
                        data: {
                            id: this.id
                        },
                        error: function (xhr) {
                            alert("Error " + xhr.status);
                            console.log(xhr.responseText);
                        },
                        success: function (data) {
                            if (data.indexOf("Error") > -1)
                                alert(data);
                            else {
                                if (data == "succeed") {
                                    getFileList("06", $("#tmpGuid").val());
                                }
                            }
                        }
                    });
                }
            });

            //下載檔案
            $(document).on("click", "a[name='downloadbtn']", function () {
                var id = $(this).attr("fid");
                location.href = "../DOWNLOAD.aspx?v=" + id;
            });
        });// end js

        //fancybox open
        function openEditBlock(titStr) {
            $("#editblock").dialog({
                title: titStr,
                autoOpen: false,
                width: 800,
                height: $(window).height() - 200,
                closeOnEscape: false,
                position:{ my: "center", at: "center", of: window },
                modal: true,
                resizable: false
            });

            $("#editblock").dialog("open");
        }

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
                    if ($(data).find("Error").length > 0)
                        alert(data);
                    else {
                        if (data != null) {
                            $("#tablist tbody").empty();
                            var tabstr = '';
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function (i) {
                                    tabstr += '<tr>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + (i + 1) + '</td>';
                                    tabstr += '<td align="center" nowrap="nowrap">' + $(this).children("N_Date").text().trim() + '</td>';
                                    tabstr += '<td style="word-break:break-all;">' + $(this).children("N_Title").text().trim() + '</td>';
                                    tabstr += '<td align="center"><input type="button" class="genbtn" name="modbtn" aid="' + $(this).children("N_ID").text().trim() + '" agid="' + $(this).children("N_Guid").text().trim() + '" value="修改" />';
                                    tabstr += '<input type="button" class="genbtn" name= "delbtn" aid= "' + $(this).children("N_ID").text().trim() + '" value="刪除" /></td > ';
                                    tabstr += '</tr>';
                                });
                            }
                            else
                                tabstr += "<tr><td colspan='6'>查詢無資料</td></tr>";
                            $("#tablist tbody").append(tabstr);
                            $(".stripeMe tr").mouseover(function () { $(this).addClass("spe"); }).mouseout(function () { $(this).removeClass("spe"); });
                            $(".stripeMe table:not(td > table) > tbody > tr:not('.spe'):even").addClass("alt");
                            CreatePage(p, $("total", data).text());
                        }
                    }
                }
            });
        }

        //撈修改那筆公告
        function getDataByID(id) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getNews_ById.aspx",
                data: {
                    N_ID: id
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
                                    //N_Date N_Title N_Content modid
                                    $("#N_Date").val($(this).children("N_Date").text().trim());
                                    $("#N_Title").val($(this).children("N_Title").text().trim());
                                    $("#N_Content").val($(this).children("N_Content").text().trim());
                                    $("#modid").val($(this).children("N_ID").text().trim());
                                });
                            }
                        }
                    }
                }
            });
        }

        //刪除公告
        function delData(id) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/deleteNews.aspx",
                data: {
                    N_ID: id
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if ($(data).find("Error").length > 0)
                        alert(data);
                    else {
                        alert("刪除成功");
                        getData(0);
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
                            $("#FileList").find("tbody").empty();
                            var tabstr = "<tbody>";
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function (i) {
                                    tabstr += '<tr>';
                                    tabstr += '<td><a href="javascript:void(0);" name="downloadbtn" fid="' + $(this).children("download_id").text().trim() + '">'
                                    tabstr += $(this).children("file_orgname").text().trim() + $(this).children("file_exten").text().trim() + '</a></td>';
                                    tabstr += '<td align="center" width="50px"><a href="javascript:void(0);" id=' + $(this).children("file_id").text().trim() + ' name="delfilebtn"><img src="../App_Themes/images/icon-delete-new.png" /></a></td>';
                                    tabstr += '</tr>';
                                });
                                tabstr += '</tbody>';
                                $("#FileList").append(tabstr);
                            }
                            else
                                $("#FileList").append('<tr><td colspan="2">查詢無資料</td></tr>');
                            $(".mfile").show();
                        }
                    }
                }
            });
        }        
        
        //新增/修改 送出後回傳值
        function feedback(msg) {
            if ($(msg).find("Error").length > 0) {
                alert(msg);
            } else {
                alert(msg);
                getData(0);
                $("#editblock").dialog("close");
            }
        }

        function upFile_feedback(tmpGuid) {
            getFileList("06", tmpGuid);
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
    <style>
        /* Table Mouseover */
        .tabcss tbody tr td {background-color: transparent;}
        .tabcss tbody tr:nth-child(even) {background-color: #fbfdf8;}
        .tabcss tbody tr:hover {background-color: #FEFBC2;}
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <input type="hidden" id="tmpGuid" />
    <input type="hidden" id="sortMethod" name="sortMethod" value="desc" />
    <input type="hidden" id="sortName" name="sortName" value="M_CreateDate" />
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">最新消息</span></div>
            <!-- left -->
            <div class="right"></div>
            <!-- right -->
        </div>
        <!-- twocol -->

        <div class="twocol" style="margin-top: 20px;">
            <div class="left">
                發布日期：<input type="text" id="SearchDate" class="inputex" />&nbsp;
                關鍵字：<input type="text" id="SearchStr" class="inputex" />&nbsp;
                <input type="button" value="查詢" class="genbtn" onclick="getData(0);" />

            </div>
            <!-- left -->
            <div class="right">
                <input type="button" id="addbtn" value="新增" class="genbtn" />
            </div>
            <!-- right -->
        </div>
        <!-- twocol -->
        <br />
        <div class="stripeMe margin5T font-normal tabcss">
            <table id="tablist" width="100%" border="0" cellspacing="0" cellpadding="0">
                <thead>
                    <tr>
                        <th nowrap="nowrap" style="width: 40px;">項次</th>
                        <th nowrap="nowrap" style="width: 150px;">發布日期</th>
                        <th nowrap="nowrap">標題</th>
                        <th nowrap="nowrap" style="width: 150px;">動作</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
            <div id="changpage" class="margin20T textcenter"></div>
        </div>
    </div>

    <div id="editblock" style="display: none;">
        <div class="stripeMe">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                    <th><span style="color: red;">*</span>標題</th>
                    <td colspan="3"><input type="text" id="N_Title" class="inputex width100 str" maxlength="200" /></td>
                </tr>
                <tr>
                    <th><span style="color: red;">*</span>發布日期</th>
                    <td><input type="text" id="N_Date" class="inputex str" maxlength="50" /></td>
                </tr>
                <tr class="aptr">
                    <th><span style="color: red;">*</span>說明</th>
                    <td><textarea id="N_Content" class="inputex width100 str" style="height:150px;"></textarea></td>
                </tr>
                <tr class="aptr">
                    <th>附件上傳</th>
                    <td><input type="button" class="genbtnS" id="FileBtn" value="選擇檔案" /></td>
                </tr>
            </table>
            <div class="mfile" style="margin-top: 10px; display: none;">附件檔</div>
            <div class="stripeMe mfile tabcss" style="margin-bottom: 10px; display: none;">
                <table id="FileList" width="100%" border="0" cellspacing="0" cellpadding="0">
                    <thead>
                        <tr>
                            <th>檔案名稱</th>
                            <th>刪除</th>
                        </tr>
                    </thead>
                    <tr>
                        <td colspan="2">查詢無資料</td>
                    </tr>
                </table>
            </div>
        </div>
        <div style="text-align: right; margin-top: 10px;">
            <input type="button" id="subbtn" class="genbtn" value="送出" />
            <input type="button" id="cancelbtn" class="genbtn" value="取消" />
            <input type="hidden" id="modid" value="" />
        </div>
    </div>
</asp:Content>

