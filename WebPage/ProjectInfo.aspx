<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="ProjectInfo.aspx.cs" Inherits="WebPage_ProjectInfo" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <script type="text/javascript">
        $(document).ready(function () {
            getFileList("01", "#moneyFileList");
            getFileList("02", "#plandescFileList");

            //上傳 button
            $(document).on("click", "#MoneyUpBtn,#PlanDescUpBtn", function () {
                var type = (this.id == "MoneyUpBtn") ? "01" : "02";
                $.fancybox({
                    href: "File_Upload.aspx?v=" + $("#<%= tmpguid.ClientID %>").val() + "&tp=" + type,
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
                    var str_atp = $(this).attr("atp");
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
                                    if (str_atp == "01") {
                                        getFileList("01", "#moneyFileList");
                                    }
                                    else {
                                        getFileList("02", "#plandescFileList");
                                    }
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
        });//js end

        function getFileList(type,tabName) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/GetFileList.aspx",
                data: {
                    pGuid: $("#<%= tmpguid.ClientID %>").val(),
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
                            $(tabName).find("tbody").empty();
                            var tabstr = "<tbody>";
                            if ($(data).find("data_item").length > 0) {
                                $(data).find("data_item").each(function (i) {
                                    tabstr += '<tr>';
                                    tabstr += '<td><a href="javascript:void(0);" name="downloadbtn" fid="' + $(this).children("download_id").text().trim() + '">'
                                    tabstr += $(this).children("file_orgname").text().trim() + $(this).children("file_exten").text().trim() + '</a></td>';
                                    tabstr += '<td align="center" width="50px"><a href="javascript:void(0);" id=' + $(this).children("file_id").text().trim() + ' name="delfilebtn" atp="' + $(this).children("file_type").text().trim() +
                                        '"><img src="../App_Themes/images/icon-delete-new.png" /></a></td>';
                                    tabstr += '</tr>';
                                });
                                tabstr += '</tbody>';
                                $(tabName).append(tabstr);
                                $(tabName).find("tr").mouseover(function () { $(this).addClass("spe"); }).mouseout(function () { $(this).removeClass("spe"); });
                                $(tabName+" table:not(td > table) > tbody > tr:not('.spe'):even").addClass("alt");
                            }
                            else
                                $(tabName).append('<tr><td colspan="2">查詢無資料</td></tr>');
                        }
                    }
                }
            });
        }

        function upFile_feedback(type) {
            if (type == "01")
                getFileList(type, "#moneyFileList");
            else
                getFileList(type, "#plandescFileList");
        }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">計畫書基本資料</span></div>
            <!-- left -->
            <div class="right">基本資料 / 計畫書基本資料</div>
            <!-- right -->
        </div>
        <!-- twocol -->
        <div class="OchiTrasTable width100 TitleLength08">
            <!-- 雙欄 -->
            <div class="OchiRow">
                <!--管理者挑選承辦人-->
                <%--<div class="OchiThird " style="display:none;" id="div_sel_people">
                    <div class="OchiCell OchiTitle TitleSetWidth">承辦人</div>
                    <div class="OchiCell width100" id="div_img_sel_people">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <div class="OchiCellInner width100">
                                <img src="../App_Themes/images/btn-search.gif" style="cursor:pointer" id="Pbox" onclick="openfancybox(this)" />
                                <span id="sel_people_name" style="color:blue;"></span>
                                <input type="hidden" id="sel_people" class="str" />
                            </div>
                        </div>
                        <!-- OchiTableInner -->
                        <!-- cell內容end -->
                    </div>
                </div>--%>
                <div class="OchiHalf " id="div_outer_city">
                    <div class="OchiCell OchiTitle TitleSetWidth">執行機關</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <div class="OchiCellInner width100" id="div_city">
                            </div>
                        </div>
                        <!-- OchiTableInner -->
                        <!-- cell內容end -->
                    </div>
                </div>
                <!-- OchiHalf -->
                <div class="OchiHalf " id="div_outer_office">
                    <div class="OchiCell OchiTitle TitleSetWidth">承辦局處</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <div class="OchiCellInner width100" id="div_office">
                            </div>
                        </div>
                        <!-- OchiTableInner -->
                        <!-- cell內容end -->
                    </div>
                </div>
                <!-- OchiHalf -->
            </div>
            <!-- OchiRow -->

            <!--執行期程 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">執行期程</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100">
                        <div class="OchiCellInner nowrap textcenter">開始:</div>
                        <div class="OchiCellInner width20">
                            <input type="text" class="width100 inputex" disabled="disabled"  id="txt_alldates" /></div>
                        <div class="OchiCellInner nowrap textcenter">&nbsp;~&nbsp;</div>
                        <div class="OchiCellInner nowrap textcenter">結束:</div>
                        <div class="OchiCellInner width20">
                            <input type="text" class="width100 inputex" disabled="disabled" id="txt_alldatee" /></div>
                        <div class="OchiCellInner nowrap textcenter">合計:</div>
                        <div class="OchiCellInner width10">
                            <input type="text" class="width100 inputex" disabled="disabled" id="txt_allmonth" /></div>
                        <div class="OchiCellInner nowrap width30">&nbsp;月</div>
                    </div>
                    <!-- OchiTableInner -->
                    <!-- cell內容end -->
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 第一期起訖 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">第1期</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100">
                        <div class="OchiCellInner nowrap textcenter">開始:</div>
                        <div class="OchiCellInner width20">
                            <input type="text" class="width100 inputex Jdatepicker" id="txt_1_sdate" value="" autocomplete="off" /></div>
                        <div class="OchiCellInner nowrap textcenter">&nbsp;~&nbsp;</div>
                        <div class="OchiCellInner nowrap textcenter">結束:</div>
                        <div class="OchiCellInner width20">
                            <input type="text" class="width100 inputex Jdatepicker" id="txt_1_edate" autocomplete="off" /></div>
                        <div class="OchiCellInner nowrap textcenter" style="display:none;">合計:</div>
                        <div class="OchiCellInner width45">
                            <input type="text" class="width100 inputex" id="txt_1_month_sum" disabled="disabled" style="display:none;" />
                        </div>
                        <div class="OchiCellInner nowrap width30" style="display:none;">&nbsp;月</div>
                    </div>
                    <!-- OchiTableInner -->
                    <!-- cell內容end -->
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 第二期起訖 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">第2期</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100">
                        <div class="OchiCellInner nowrap textcenter">開始:</div>
                        <div class="OchiCellInner width20">
                            <input type="text" class="width100 inputex Jdatepicker" id="txt_2_sdate" autocomplete="off" /></div>
                        <div class="OchiCellInner nowrap textcenter">&nbsp;~&nbsp;</div>
                        <div class="OchiCellInner nowrap textcenter">結束:</div>
                        <div class="OchiCellInner width20">
                            <input type="text" class="width100 inputex Jdatepicker" id="txt_2_edate" autocomplete="off" /></div>
                        <div class="OchiCellInner nowrap textcenter" style="display:none;">合計:</div>
                        <div class="OchiCellInner width45">
                            <input type="text" class="width100 inputex" id="txt_2_month_sum" disabled="disabled" style="display:none;" />
                        </div>
                        <div class="OchiCellInner nowrap width30" style="display:none;">&nbsp;月</div>
                    </div>
                    <!-- OchiTableInner -->
                    <!-- cell內容end -->
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 第三期起訖 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">第3期</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100">
                        <div class="OchiCellInner nowrap textcenter">開始:</div>
                        <div class="OchiCellInner width20">
                            <input type="text" class="width100 inputex Jdatepicker" id="txt_3_sdate" autocomplete="off" /></div>
                        <div class="OchiCellInner nowrap textcenter">&nbsp;~&nbsp;</div>
                        <div class="OchiCellInner nowrap textcenter">結束:</div>
                        <div class="OchiCellInner width20">
                            <input type="text" class="width100 inputex Jdatepicker" id="txt_3_edate" value="" autocomplete="off" /></div>
                        <div class="OchiCellInner nowrap textcenter" style="display:none;">合計:</div>
                        <div class="OchiCellInner width45">
                            <input type="text" class="width100 inputex" id="txt_3_month_sum" disabled="disabled" style="display:none;" />
                        </div>
                        <div class="OchiCellInner nowrap width30" style="display:none;">&nbsp;月</div>
                    </div><!-- OchiTableInner -->
                    <!-- cell內容end -->
                </div><!-- OchiCell -->
            </div><!-- OchiRow -->

            <!-- 核定之經費額度 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">核定之經費額度</div>
                <div class="OchiCell width100">
                    <div class="stripeMe margin5T font-normal">
                        <!--設備汰換-->
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <thead>
                                <tr>
                                    <th nowrap="nowrap">期別</th>
                                    <th nowrap="nowrap">第1期</th>
                                    <th nowrap="nowrap">第2期</th>
                                    <th nowrap="nowrap">第3期</th>
                                    <th nowrap="nowrap">全程</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td>節電基礎工作</td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_money" id="txt_money_item1_1" t="strfloat3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_money" id="txt_money_item1_2" t="strfloat3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_money" id="txt_money_item1_3" t="strfloat3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_money_item1_sum" disabled="disabled" t="strfloat3" /></td>
                                </tr>
                                <tr>
                                    <td nowrap="nowrap">設備汰換及智慧用電</td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_money" id="txt_money_item3_1" t="strfloat3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_money" id="txt_money_item3_2" t="strfloat3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_money" id="txt_money_item3_3" t="strfloat3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_money_item3_sum" disabled="disabled" t="strfloat3" /></td>
                                </tr>
                                <tr>
                                    <td>因地制宜</td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_money" id="txt_money_item2_1" t="strfloat3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_money" id="txt_money_item2_2" t="strfloat3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_money" id="txt_money_item2_3" t="strfloat3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_money_item2_sum" disabled="disabled" t="strfloat3" /></td>
                                </tr>
                                <tr>
                                    <td>擴大補助</td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_money" id="txt_money_item4_1" t="strfloat3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_money" id="txt_money_item4_2" t="strfloat3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_money" id="txt_money_item4_3" t="strfloat3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_money_item4_sum" disabled="disabled" t="strfloat3" /></td>
                                </tr>
                                <tr class="spe">
                                    <td>總計</td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_money_sum1" disabled="disabled" t="strfloat3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_money_sum2" disabled="disabled" t="strfloat3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_money_sum3" disabled="disabled" t="strfloat3" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_money_sumall" disabled="disabled" t="strfloat3" /></td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                    <div class="textright">單位:仟元</div>
                    <div class="twocol margin5T">
                        <div class="left"><a href="javascript:void(0);" id="MoneyUpBtn" class="genbtnS" style="display:none;">核定函上傳</a></div> 
                    </div>
                    <div class="mfile" style="margin-top:10px;display:none;">附件檔</div>
                    <div class="stripeMe mfile" style="margin-bottom:10px;display:none;">
                        <table id="moneyFileList" width="100%" border="0" cellspacing="0" cellpadding="0">
                            <thead>
                                <tr>
                                    <th>檔案名稱</th>
                                    <th>刪除</th>
                                </tr>
                            </thead>
                        </table>
                    </div>
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 其他經費來源 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">其他經費來源</div>
                <div class="OchiCell width100">
                    <div>
                        <input type="checkbox" id="chk_Other_Oneself" />&nbsp;
                        自籌款，金額:<input type="text" class="inputex" size="10" id="txt_Other_Oneself_Money" t="strfloat" />&nbsp;仟元
                    </div>
                    <div class="margin10T">
                        <input type="checkbox" id="chk_Other_Other" />&nbsp;
                        其他機關補助，機關名稱:<input type="text" class="inputex" id="txt_Other_Other_name" />，
                        金額:<input type="text" class="inputex" size="10" id="txt_Other_Other_Money" t="strfloat" />&nbsp;仟元
                    </div>
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 承諾節電目標 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">承諾節電目標</div>
                <div class="OchiCell width100">
                    <textarea rows="4" class="inputex width100" maxlength="1000" id="txt_Target"></textarea>
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 本期計畫推動摘要 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">本期計畫推動摘要</div>
                <div class="OchiCell width100">
                    <textarea rows="7" class="inputex width100 itemhint margin5T" title="1,000字以內" maxlength="1000" id="txt_summary"></textarea>
                    <div class="twocol margin5T">
                        <div class="left"><a href="javascript:void(0);" id="PlanDescUpBtn" class="genbtnS" style="display:none;">計畫書上傳</a></div>
                    </div>
                    <div class="pdfile" style="margin-top:5px;display:none;">附件檔</div>
                    <div class="stripeMe pdfile" style="margin-bottom:10px;display:none;">
                        <table id="plandescFileList" width="100%" border="0" cellspacing="0" cellpadding="0">
                            <thead>
                                <tr>
                                    <th>檔案名稱</th>
                                    <th>刪除</th>
                                </tr>
                            </thead>
                        </table>
                    </div>
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 設備汰換與智慧用電預計完成數 單欄 -->
            <div class="OchiRow">
                <div class="OchiCell OchiTitle TitleSetWidth">設備汰換與智慧用電預計完成數</div>
                <div class="OchiCell width100">
                    <div class="stripeMe margin5T font-normal">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <thead>
                                <tr>
                                    <th nowrap="nowrap">項目</th>
                                    <th nowrap="nowrap">第1期</th>
                                    <th nowrap="nowrap">第2期</th>
                                    <th nowrap="nowrap">第3期</th>
                                    <th nowrap="nowrap">全程</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td>無風管冷氣(kW)<br />
                                        註：每台冷氣約4kW</td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item1_1" t="strfloat" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item1_2" t="strfloat" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item1_3" t="strfloat" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item1_all" disabled="disabled" t="strfloat" /></td>
                                </tr>
                                <tr>
                                    <td>老舊辦公室照明(具)</td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item2_1" t="strint" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item2_2" t="strint" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item2_3" t="strint" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_Finish_item2_all" disabled="disabled" t="strint" /></td>
                                </tr>
                                <tr>
                                    <td nowrap="nowrap">室內停車場智慧照明(盞)</td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item3_1" t="strint" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item3_2" t="strint" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item3_3" t="strint" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_Finish_item3_all" disabled="disabled" t="strint" /></td>
                                </tr>
                                <tr>
                                    <td>中型能管系統(套)</td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item4_1" t="strint" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item4_2" t="strint" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item4_3" t="strint" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_Finish_item4_all" disabled="disabled" t="strint" /></td>
                                </tr>
                                <tr>
                                    <td>大型能管系統(套)</td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item5_1" t="strint" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item5_2" t="strint" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" name="name_FinishItem" id="txt_Finish_item5_3" t="strint" /></td>
                                    <td>
                                        <input type="text" class="inputex width100" id="txt_Finish_item5_all" disabled="disabled" t="strint" /></td>
                                </tr>

                            </tbody>
                        </table>
                    </div>
                    <!--<div class="textright">單位:仟元</div>-->

                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 簽核資料 單欄 -->
            <div class="OchiRow" name="div_chk" style="display:none;">
                <div class="OchiCell OchiTitle TitleSetWidth">承辦人</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100">
                        <div class="OchiCellInner nowrap textcenter">姓名:</div>
                        <div class="OchiCellInner width28">
                            <input type="text" class="width100 inputex" id="txt_Cname" disabled="disabled" />
                        </div>
                        <div class="OchiCellInner nowrap textcenter">職稱:</div>
                        <div class="OchiCellInner width28">
                            <input type="text" class="width100 inputex" id="txt_JobTitle" disabled="disabled" /></div>
                        <div class="OchiCellInner nowrap textcenter">電話:</div>
                        <div class="OchiCellInner width28">
                            <input type="text" class="width100 inputex" id="txt_Tel" disabled="disabled" /></div>
                    </div>
                    <!-- OchiTableInner -->
                    <div class="OchiTableInner width100 margin10T">
                        <div class="OchiCellInner nowrap textcenter">手機:</div>
                        <div class="OchiCellInner width28">
                            <input type="text" class="width100 inputex" id="txt_Phone" disabled="disabled" /></div>
                        <div class="OchiCellInner nowrap textcenter">傳真:</div>
                        <div class="OchiCellInner width28">
                            <input type="text" class="width100 inputex" id="txt_Fax" disabled="disabled" /></div>
                        <div class="OchiCellInner width33"></div>
                    </div>
                    <!-- OchiTableInner -->
                    <div class="OchiTableInner width100 margin10T">
                        <div class="OchiCellInner nowrap textcenter">Email:</div>
                        <div class="OchiCellInner width100">
                            <input type="text" class="width100 inputex" id="txt_Email" disabled="disabled" /></div>
                    </div>
                    <!-- OchiTableInner -->
                    <div class="OchiTableInner width100 margin10T">
                        <div class="OchiCellInner nowrap textcenter">地址:</div>
                        <div class="OchiCellInner width100">
                            <input type="text" class="width100 inputex" id="txt_Addr" disabled="disabled" /></div>
                    </div>
                    <!-- OchiTableInner -->
                    <!-- cell內容end -->
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->

            <!-- 簽核資料 單欄 -->
            <div class="OchiRow" name="div_chk" style="display:none;">
                <div class="OchiCell OchiTitle TitleSetWidth">承辦主管</div>
                <div class="OchiCell width100">
                    <!-- cell內容start -->
                    <div class="OchiTableInner width100">
                        <div class="OchiCellInner nowrap textcenter">姓名:</div>
                        <div class="OchiCellInner width28">
                            <input type="text" class="width100 inputex" id="txt_ManagerCname" disabled="disabled" /></div>
                        <div class="OchiCellInner nowrap textcenter">職稱:</div>
                        <div class="OchiCellInner width28">
                            <input type="text" class="width100 inputex" id="txt_ManagerJobTitle" disabled="disabled" /></div>
                        <div class="OchiCellInner nowrap textcenter">電話:</div>
                        <div class="OchiCellInner width28">
                            <input type="text" class="width100 inputex" id="txt_ManagerTel" disabled="disabled" /></div>
                    </div>
                    <!-- OchiTableInner -->
                    <div class="OchiTableInner width100 margin10T">
                        <div class="OchiCellInner nowrap textcenter">手機:</div>
                        <div class="OchiCellInner width28">
                            <input type="text" class="width100 inputex" id="txt_ManagerPhone" disabled="disabled" /></div>
                        <div class="OchiCellInner nowrap textcenter">傳真:</div>
                        <div class="OchiCellInner width28">
                            <input type="text" class="width100 inputex" id="txt_ManagerFax" disabled="disabled" /></div>
                        <div class="OchiCellInner width33"></div>
                    </div>
                    <!-- OchiTableInner -->
                    <div class="OchiTableInner width100 margin10T">
                        <div class="OchiCellInner nowrap textcenter">Email:</div>
                        <div class="OchiCellInner width100">
                            <input type="text" class="width100 inputex" id="txt_ManagerEmail" disabled="disabled" /></div>
                    </div>
                    <!-- OchiTableInner -->
                    <div class="OchiTableInner width100 margin10T">
                        <div class="OchiCellInner nowrap textcenter">地址:</div>
                        <div class="OchiCellInner width100">
                            <input type="text" class="width100 inputex" id="txt_ManagerAddr" disabled="disabled" /></div>
                    </div>
                    <!-- OchiTableInner -->
                    <!-- cell內容end -->
                </div>
                <!-- OchiCell -->
            </div>
            <!-- OchiRow -->
        </div>
        <!-- OchiTrasTable -->
        <div class="twocol margin15T margin5B">
            <div class="right">
                <a href="javascript:void(0)" class="genbtn" id="btn_save">儲存</a>
                <a href="javascript:void(0)" class="genbtn" id="btn_check" style="display:none;">定稿</a>
                <a href="javascript:void(0)" class="genbtn" id="btn_next">下一步</a>
                <!--<a href="javascript:void(0)" class="genbtn" id="btn_cancel">取消</a>-->
            </div>
        </div>
        <!-- twocol -->
    </div>
    <input type="text" id="hidden_I_ID" style="display:none;" />
    <input type="text" id="hidden_I_Guid" style="display:none;" />
    <input type="text" id="hidden_M_ID" style="display:none;" />
    <input type="text" id="hidden_M_competence" style="display:none;" />
    <input type="text" id="hidden_I_People" style="display:none;" />
    <input type="text" id="hidden_user" style="display:none;" />
    <input type="hidden" id="tmpguid" runat="server" />
    <script>
        setInterval(function () {
            var breakStatus = false;
            if ($("#txt_alldates").val() == "" || $("#txt_alldatee").val() == "") {
                breakStatus = true;
                return;
            }
            var che_oneself = $("#chk_Other_Oneself").is(":checked") ? "Y" : "";
            var che_other = $("#chk_Other_Other").is(":checked") ? "Y" : "";
            if (che_oneself != "Y" && $("#txt_Other_Oneself_Money").val() != "" && $("#txt_Other_Oneself_Money").val() != "0.0") {
                breakStatus = true;
                return;
            }
            if (che_other != "Y" && (($("#txt_Other_Other_Money").val() != "" && $("#txt_Other_Other_Money").val() != "0.0") || $("#txt_Other_Other_name").val() != "")) {
                breakStatus = true;
                return;
            }
            //如果這個輸入欄位disabled掉 表示有人是key URL進來 會把填寫欄位都disabled 這時候不需要自動存檔
            if ($("#txt_1_sdate").prop("disabled")==true) {
                breakStatus = true;
                return;
            }
            if (breakStatus == true)
                return;
            mod_project("auto");
        }, 1200000);//20 minutes 1200000
        $(function () {
            //撈計畫資訊
            load_projectinfo();
            //填執行期程日期自動算出幾個月
            $(".Jdatepicker").blur(function () {
                dateDiff($(this).attr("id"));
            });

            //數字欄位判斷整數或小數點一位
            //數字的輸入欄位 keyup
            $(document).on("keyup", "input", function () {
                $("input").css("border-color", "");
                var thist = $(this).attr("t");
                var strVal = $(this).val();
                var re = /^(\+|-)?\d+$/;
                if ($(this).val() != "" && thist == "strint")//正整數
                {
                    if (re.test(strVal) && strVal > 0) {
                        return true;
                    } else {
                        alert("請輸入數字且為正整數");
                        $(this).css("border-color", "red");
                        $(this).val("");
                        return false;
                    }
                }
                if ($(this).val() != "" && thist == "strfloat")//小數
                {
                    if (strVal.indexOf(".") > 0) {//有小數點
                        var splitVal = strVal.split('.');
                        if (splitVal[1].length > 1) {
                            if (re.test(strVal) && strVal > 0) {
                                return true;
                            } else {
                                alert("請輸入數字且只能到小數第一位");
                                $(this).css("border-color", "red");
                                $(this).val("");
                                return false;
                            }
                        } else {
                            if (splitVal[1] == "" || (splitVal[1] != "" && re.test(splitVal[1]) && splitVal[1] >= 0)) {
                                return true;
                            } else {
                                alert("請輸入數字");
                                $(this).css("border-color", "red");
                                $(this).val("");
                                return false;
                            }
                        }
                    }
                }
                if ($(this).val() != "" && thist == "strfloat3")//小數第三位
                {
                    if (strVal.indexOf(".") > 0) {//有小數點
                        var splitVal = strVal.split('.');
                        if (splitVal[1].length > 3)
                        {
                             alert("請輸入數字且只能到小數第三位");
                            $(this).css("border-color", "red");
                            $(this).val("");
                            return false;
                        } else {
                            if (splitVal[1] == "" || (splitVal[1] != "" && re.test(splitVal[1]) && splitVal[1] >= 0)) {
                                return true;
                            } else {
                                alert("請輸入數字");
                                $(this).css("border-color", "red");
                                $(this).val("");
                                return false;
                            }
                        }
                    }
                }
                return true;

            });

            //填核定之經費額度 輸入檢查&自動加總 (設備汰換)
            $("input[name='name_money']").blur(function () {
                if ($(this).val() == "") {
                    $(this).val("0.0");
                }
                sumMoney();
            });

            //填核定之經費額度 輸入檢查&自動加總 (擴大補助)
            $("input[name='name_Exmoney']").blur(function () {
                if ($(this).val() == "") {
                    $(this).val("0.0");
                }
                sumExMoney();
            });

            //設備汰換與智慧用電預計完成數 輸入檢查&自動加總
            $("input[name='name_FinishItem']").blur(function () {
                if ($(this).val() == "") {
                    if ($(this).attr("t") == "strint") {
                        $(this).val("0");
                    } else {
                        $(this).val("0.0");
                    }
                    
                }
                sumFinishItem();
            });
            //儲存
            $("#btn_save").click(function () {
                mod_project($(this).attr("id"));
            });
            //取消
            $("#btn_cancel").click(function () {
                if(confirm("確認是否已儲存資料，即將放棄編輯返回計畫列表?")){
                    location.href = "ProjectList.aspx";
                }
                
            });
            //下一步
            $("#btn_next").click(function () {
                //下一步跟儲存合併 點下一步的同時就會先儲存在導頁
                //不跳出"確認是否已儲存資料"的訊息，但該擋的欄位還是要擋
                if ($("#btn_next").html() == "下一頁") {
                    location.href = "CheckPoint.aspx?v=" + $("#hidden_M_ID").val() + "";
                } else {
                    //mod_project($(this).attr("id"));
                    if (confirm("請確認資料已儲存，是否進入下一步?")) {
                        location.href = "CheckPoint.aspx?v=" + $("#hidden_M_ID").val() + "";
                    }
                }
            });
        });
        //撈計畫資料
        function load_projectinfo() {
            $.ajax({
                type: "POST",
                async: true, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/ProjectInfo.ashx",
                data: {
                    func: "load_projectinfo",
                    mid: $.getParamValue('v'),//人員ID
                    str_tmpguid:$("#<%= tmpguid.ClientID %>").val()
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if (data=="timeout") {
                        alert("請重新登入");
                        location.href = "Login.aspx";
                    } else if (data == "nodata") {
                        alert("查無填寫人員資料");
                    } else if (data == "error") {
                        alert("error");
                    } else if (data=="nodate") {
                        alert("執行期程尚未設定,請聯絡系統管理員");
                        location.href = "ProjectList.aspx";
                    } else if (data == "DiffCity") {
                        alert("您沒有權限進入此頁面");
                        window.location = "ProjectList.aspx";
                    }
                    //管理者代填的功能拿掉20171201
                    //else if (data == "SA") {//管理者進來新增資料
                    //    //顯示上傳按鈕及上傳附件
                    //    $(".mfile").show();
                    //    $(".pdfile").show();
                    //    $("#MoneyUpBtn,#PlanDescUpBtn").show();
                    //    $("#div_sel_people").show();
                    //    $("#div_outer_city,#div_outer_office").removeClass("OchiHalf");
                    //    $("#div_outer_city,#div_outer_office").addClass("OchiThird");
                    //    $("#hidden_M_competence").val(data);
                    //}
                    else {
                        if (data.length != 0 && data[0].I_ID != null) {
                            $("#hidden_I_ID").val(data[0].I_ID);
                            $("#hidden_I_Guid").val(data[0].I_GUID);
                            if (data[0].I_1_Sdate == "") {
                                $("#txt_1_sdate").val(data[0].all_dates);
                            } else {
                                $("#txt_1_sdate").val(data[0].I_1_Sdate);
                            }
                            $("#txt_1_edate").val(data[0].I_1_Edate);
                            $("#txt_2_sdate").val(data[0].I_2_Sdate);
                            $("#txt_2_edate").val(data[0].I_2_Edate);
                            $("#txt_3_sdate").val(data[0].I_3_Sdate);
                            if (data[0].I_1_Sdate == "") {
                                $("#txt_3_edate").val(data[0].all_datee);
                            } else {
                                $("#txt_3_edate").val(data[0].I_3_Edate);
                            }
                            $("#txt_money_item1_1").val(data[0].I_Money_item1_1);
                            $("#txt_money_item1_2").val(data[0].I_Money_item1_2);
                            $("#txt_money_item1_3").val(data[0].I_Money_item1_3);
                            $("#txt_money_item1_sum").val(data[0].I_Money_item1_all);
                            $("#txt_money_item2_1").val(data[0].I_Money_item2_1);
                            $("#txt_money_item2_2").val(data[0].I_Money_item2_2);
                            $("#txt_money_item2_3").val(data[0].I_Money_item2_3);
                            $("#txt_money_item2_sum").val(data[0].I_Money_item2_all);
                            $("#txt_money_item3_1").val(data[0].I_Money_item3_1);
                            $("#txt_money_item3_2").val(data[0].I_Money_item3_2);
                            $("#txt_money_item3_3").val(data[0].I_Money_item3_3);
                            $("#txt_money_item3_sum").val(data[0].I_Money_item3_all);
                             //擴大補助
                            $("#txt_money_item4_1").val(data[0].I_Money_item4_1);
                            $("#txt_money_item4_2").val(data[0].I_Money_item4_2);
                            $("#txt_money_item4_3").val(data[0].I_Money_item4_3);
                            $("#txt_money_item4_sum").val(data[0].I_Money_item4_all);
                            if (data[0].I_Other_Oneself == "Y") {
                                $("#chk_Other_Oneself").attr("checked", true);
                            }
                            $("#txt_Other_Oneself_Money").val(data[0].I_Other_Oneself_Money);
                            if (data[0].I_Other_Other == "Y") {
                                $("#chk_Other_Other").attr("checked", true);
                            }
                            $("#txt_Other_Other_name").val(data[0].I_Other_Other_name);
                            $("#txt_Other_Other_Money").val(data[0].I_Other_Other_Money);
                            $("#txt_Target").val(data[0].I_Target);
                            $("#txt_summary").val(data[0].I_Summary);
                            $("#txt_Finish_item1_1").val(data[0].I_Finish_item1_1);
                            $("#txt_Finish_item1_2").val(data[0].I_Finish_item1_2);
                            $("#txt_Finish_item1_3").val(data[0].I_Finish_item1_3);
                            $("#txt_Finish_item1_all").val(data[0].I_Finish_item1_all);
                            $("#txt_Finish_item2_1").val(data[0].I_Finish_item2_1);
                            $("#txt_Finish_item2_2").val(data[0].I_Finish_item2_2);
                            $("#txt_Finish_item2_3").val(data[0].I_Finish_item2_3);
                            $("#txt_Finish_item2_all").val(data[0].I_Finish_item2_all);
                            $("#txt_Finish_item3_1").val(data[0].I_Finish_item3_1);
                            $("#txt_Finish_item3_2").val(data[0].I_Finish_item3_2);
                            $("#txt_Finish_item3_3").val(data[0].I_Finish_item3_3);
                            $("#txt_Finish_item3_all").val(data[0].I_Finish_item3_all);
                            $("#txt_Finish_item4_1").val(data[0].I_Finish_item4_1);
                            $("#txt_Finish_item4_2").val(data[0].I_Finish_item4_2);
                            $("#txt_Finish_item4_3").val(data[0].I_Finish_item4_3);
                            $("#txt_Finish_item4_all").val(data[0].I_Finish_item4_all);
                            $("#txt_Finish_item5_1").val(data[0].I_Finish_item5_1);
                            $("#txt_Finish_item5_2").val(data[0].I_Finish_item5_2);
                            $("#txt_Finish_item5_3").val(data[0].I_Finish_item5_3);
                            $("#txt_Finish_item5_all").val(data[0].I_Finish_item5_all);
                            $("#txt_Cname").val(data[0].Cname);
                            $("#txt_JobTitle").val(data[0].JobTitle);
                            $("#txt_Tel").val(data[0].Tel);
                            $("#txt_Phone").val(data[0].Phone);
                            $("#txt_Fax").val(data[0].Fax);
                            $("#txt_CityName").val(data[0].CityName);
                            $("#txt_Email").val(data[0].Email);
                            $("#txt_Addr").val(data[0].Addr);
                            $("#txt_ManagerCname").val(data[0].ManagerCname);
                            $("#txt_ManagerJobTitle").val(data[0].ManagerJobTitle);
                            $("#txt_ManagerTel").val(data[0].ManagerTel);
                            $("#txt_ManagerPhone").val(data[0].ManagerPhone);
                            $("#txt_ManagerFax").val(data[0].ManagerFax);
                            $("#txt_ManagerCityName").val(data[0].ManagerCityName);
                            $("#txt_ManagerEmail").val(data[0].ManagerEmail);
                            $("#txt_ManagerAddr").val(data[0].ManagerAddr);
                            $("#div_city").empty().append(data[0].CityName);
                            $("#div_office").empty().append(data[0].I_Office);
                            $("#txt_alldates").val(data[0].all_dates);
                            $("#txt_alldatee").val(data[0].all_datee);
                            $("#txt_allmonth").val(monthDiff(new Date(data[0].all_dates), new Date(data[0].all_datee)));
                            $("#hidden_M_ID").val(data[0].M_ID);//
                            $("#hidden_M_competence").val(data[0].M_competence);
                            $("#hidden_I_People").val(data[0].I_People);
                            $("#hidden_user").val(data[0].user);
                            sumMoney();
                            showorhide("old", data[0].I_GUID, data[0].I_People, data[0].user, data[0].M_competence, data[0].I_Flag, data[0].chk_flag, data[0].M_ID, data[0].M_Name);
                        } else {
                            $("#txt_1_sdate").val(data[0].PD_StartDate);
                            $("#txt_3_edate").val(data[0].PD_EndDate);
                            $("#txt_alldates").val(data[0].PD_StartDate);
                            $("#txt_alldatee").val(data[0].PD_EndDate);
                            $("#txt_alldatee").val(data[0].PD_EndDate);
                            $("#div_city").empty().append(data[0].CityName);
                            $("#div_office").empty().append(data[0].I_Office);
                            $("#hidden_I_People").val(data[0].I_People);
                            $("#hidden_user").val(data[0].user);
                            $("#hidden_M_ID").val(data[0].M_ID);//
                            $("#txt_allmonth").val(monthDiff(new Date(data[0].PD_StartDate), new Date(data[0].PD_EndDate)));
                            showorhide("new", data[0].I_Guid, data[0].I_People, data[0].user, data[0].M_competence, "", data[0].chk_flag, "", "");
                        }
                        $("#btn_next").removeAttr("disabled");
                        $("#btn_next").show();
                        changeZero();
                    }
                }
            });//ajax end
        }
        //計算某一期相差幾個月
        function dateDiff(str_date_id) {
            //20190417 拿掉判斷下一期的開始日期是否在上一期的結束日之前
            //因為現在會有期程重疊 所以不擋這些
            var dtS;
            var dtE;
            var total_months;
            switch (str_date_id) {
                case "txt_1_sdate":
                case "txt_1_edate":
                    if ($("#txt_1_sdate").val() != "" && $("#txt_1_edate").val() != "") {
                        dtS = new Date($("#txt_1_sdate").val());
                        dtE = new Date($("#txt_1_edate").val());
                        if (dtS >= dtE) {
                            alert("請選擇第一期正確的結束日");
                            $("#txt_1_edate").val("");
                            //$("#txt_1_month_sum").val("");
                        } else {
                            total_months = (dtE.getFullYear() - dtS.getFullYear()) * 12 + (dtE.getMonth() - dtS.getMonth());
                            //$("#txt_1_month_sum").val(total_months+1);
                        }
                    }
                    if ($("#txt_1_sdate").val() != "") {
                        dtS = new Date($("#txt_alldates").val());
                        dtE = new Date($("#txt_1_sdate").val());
                        if (dtS > dtE) {
                            alert("請選擇第一期正確的開始日");
                            $("#txt_1_sdate").val("");
                            //$("#txt_1_month_sum").val("");
                        }
                    }
                    //if ($("#txt_2_sdate").val() != "" && $("#txt_1_edate").val() != "") {
                    //    dtS = new Date($("#txt_1_edate").val());
                    //    dtE = new Date($("#txt_2_sdate").val());
                    //    if (dtS > dtE) {
                    //        alert("請選擇第二期正確的開始日");
                    //        $("#txt_2_sdate").val("");
                    //        //$("#txt_1_month_sum").val("");
                    //    }
                    //}
                    break;
                case "txt_2_sdate":
                case "txt_2_edate":
                    if ($("#txt_2_sdate").val() != "" && $("#txt_2_edate").val() != "") {
                        dtS = new Date($("#txt_2_sdate").val());
                        dtE = new Date($("#txt_2_edate").val());
                        if (dtS >= dtE) {
                            alert("請選擇第二期正確的開始與結束日");
                            if (str_date_id == "txt_2_sdate") {
                                $("#txt_2_sdate").val("");
                            } else {
                                $("#txt_2_edate").val("");
                            }
                            //$("#txt_2_month_sum").val("");
                        } else {
                            total_months = (dtE.getFullYear() - dtS.getFullYear()) * 12 + (dtE.getMonth() - dtS.getMonth());
                            //$("#txt_2_month_sum").val(total_months+1);
                        }
                    }
                    //if ($("#txt_2_sdate").val() != "" && $("#txt_1_edate").val() != "") {
                    //    var dtS2 = new Date($("#txt_1_edate").val());
                    //    var dtE2 = new Date($("#txt_2_sdate").val());
                    //    if (dtS2 >= dtE2) {
                    //        alert("請選擇第二期正確的開始日");
                    //        $("#txt_2_sdate").val("");
                    //        //$("#txt_2_month_sum").val("");
                    //    } else {
                    //        //total_months = (dtE.getFullYear() - dtS.getFullYear()) * 12 + (dtE.getMonth() - dtS.getMonth());
                    //        //$("#txt_2_month_sum").val(total_months + 1);
                    //    }
                    //}
                    //if ($("#txt_3_sdate").val() != "" && $("#txt_2_edate").val() != "") {
                    //    var dtS2 = new Date($("#txt_2_edate").val());
                    //    var dtE2 = new Date($("#txt_3_sdate").val());
                    //    if (dtS2 >= dtE2) {
                    //        alert("請選擇第三期正確的開始日");
                    //        $("#txt_3_sdate").val("");
                    //        //$("#txt_2_month_sum").val("");
                    //    } else {
                    //        //total_months = (dtE.getFullYear() - dtS.getFullYear()) * 12 + (dtE.getMonth() - dtS.getMonth());
                    //        //$("#txt_2_month_sum").val(total_months + 1);
                    //    }
                    //}
                    break;
                case "txt_3_sdate":
                case "txt_3_edate":
                    if ($("#txt_3_sdate").val() != "" && $("#txt_3_edate").val() != "") {
                        dtS = new Date($("#txt_3_sdate").val());
                        dtE = new Date($("#txt_3_edate").val());
                        if (dtS >= dtE) {
                            alert("請選擇第三期正確的開始日");
                            $("#txt_3_sdate").val("");
                            //$("#txt_3_month_sum").val("");
                        } else {
                            //total_months = (dtE.getFullYear() - dtS.getFullYear()) * 12 + (dtE.getMonth() - dtS.getMonth());
                            //$("#txt_3_month_sum").val(monthDiff(dtS, dtE));
                        }
                    }
                    //if ($("#txt_3_sdate").val() != "" && $("#txt_2_edate").val() != "") {
                    //    var dtS2 = new Date($("#txt_2_edate").val());
                    //    var dtE2 = new Date($("#txt_3_sdate").val());
                    //    if (dtS2 >= dtE2) {
                    //        alert("請選擇第三期正確的開始日");
                    //        $("#txt_3_sdate").val("");
                    //        //$("#txt_3_month_sum").val("");
                    //    } else {
                    //        //total_months = (dtE.getFullYear() - dtS.getFullYear()) * 12 + (dtE.getMonth() - dtS.getMonth());
                    //        //$("#txt_3_month_sum").val(total_months + 1);
                    //    }
                    //}
                    //if ($("#txt_3_edate").val() != "") {
                    //    dtS = new Date($("#txt_3_edate").val());
                    //    dtE = new Date($("#txt_alldatee").val());
                    //    if (dtS > dtE) {
                    //        alert("請選擇第三期正確的結束日");
                    //        $("#txt_1_sdate").val("");
                    //        //$("#txt_1_month_sum").val("");
                    //    }
                    //}
                    break;
            }
        }
        //計算幾個月
        function monthDiff(d1, d2) {
            var months;
            months = (d2.getFullYear() - d1.getFullYear()) * 12;
            //months -= d1.getMonth() + 1;
            months -= d1.getMonth()-1;
            months += d2.getMonth();
            //2017/01/01~2017/01/23 這樣要算1個月
            //2017/01/01~2017/02/12 這樣要算2個月
            return months <= 0 ? 1 : months;
        }

        //核定之經費額度 加總
        function sumMoney() {
            //20180530調整為到小數點第三位(原本只到小數點第一位)
            var i11 = ($("#txt_money_item1_1").val() == "") ? 0 : parseFloat($("#txt_money_item1_1").val());
            var i12 = ($("#txt_money_item1_2").val() == "") ? 0 : parseFloat($("#txt_money_item1_2").val());
            var i13 = ($("#txt_money_item1_3").val() == "") ? 0 : parseFloat($("#txt_money_item1_3").val());
            var i21 = ($("#txt_money_item2_1").val() == "") ? 0 : parseFloat($("#txt_money_item2_1").val());
            var i22 = ($("#txt_money_item2_2").val() == "") ? 0 : parseFloat($("#txt_money_item2_2").val());
            var i23 = ($("#txt_money_item2_3").val() == "") ? 0 : parseFloat($("#txt_money_item2_3").val());
            var i31 = ($("#txt_money_item3_1").val() == "") ? 0 : parseFloat($("#txt_money_item3_1").val());
            var i32 = ($("#txt_money_item3_2").val() == "") ? 0 : parseFloat($("#txt_money_item3_2").val());
            var i33 = ($("#txt_money_item3_3").val() == "") ? 0 : parseFloat($("#txt_money_item3_3").val());
            var i41 = ($("#txt_money_item4_1").val() == "") ? 0 : parseFloat($("#txt_money_item4_1").val());
            var i42 = ($("#txt_money_item4_2").val() == "") ? 0 : parseFloat($("#txt_money_item4_2").val());
            var i43 = ($("#txt_money_item4_3").val() == "") ? 0 : parseFloat($("#txt_money_item4_3").val());
            $("#txt_money_item1_sum").val(parseFloat(i11 + i12 + i13).toFixed(3));
            $("#txt_money_item2_sum").val(parseFloat(i21 + i22 + i23).toFixed(3));
            $("#txt_money_item3_sum").val(parseFloat(i31 + i32 + i33).toFixed(3));
            $("#txt_money_item4_sum").val(parseFloat(i41 + i42 + i43).toFixed(3));
            $("#txt_money_sum1").val(parseFloat(i11 + i21 + i31 + i41).toFixed(3));
            $("#txt_money_sum2").val(parseFloat(i12 + i22 + i32 + i42).toFixed(3));
            $("#txt_money_sum3").val(parseFloat(i13 + i23 + i33 + i43).toFixed(3));
            $("#txt_money_sumall").val(parseFloat(i11 + i12 + i13 + i21 + i22 + i23 + i31 + i32 + i33 + i41 + i42 + i43).toFixed(3));
        }

        //設備汰換與智慧用電預計完成數 自動加總
        function sumFinishItem() {
            //txt_Finish_item1_1 txt_Finish_item1_all
            var f11 = ($("#txt_Finish_item1_1").val() == "") ? 0 : parseFloat($("#txt_Finish_item1_1").val());
            var f12 = ($("#txt_Finish_item1_2").val() == "") ? 0 : parseFloat($("#txt_Finish_item1_2").val());
            var f13 = ($("#txt_Finish_item1_3").val() == "") ? 0 : parseFloat($("#txt_Finish_item1_3").val());
            var f21 = ($("#txt_Finish_item2_1").val() == "") ? 0 : parseFloat($("#txt_Finish_item2_1").val());
            var f22 = ($("#txt_Finish_item2_2").val() == "") ? 0 : parseFloat($("#txt_Finish_item2_2").val());
            var f23 = ($("#txt_Finish_item2_3").val() == "") ? 0 : parseFloat($("#txt_Finish_item2_3").val());
            var f31 = ($("#txt_Finish_item3_1").val() == "") ? 0 : parseFloat($("#txt_Finish_item3_1").val());
            var f32 = ($("#txt_Finish_item3_2").val() == "") ? 0 : parseFloat($("#txt_Finish_item3_2").val());
            var f33 = ($("#txt_Finish_item3_3").val() == "") ? 0 : parseFloat($("#txt_Finish_item3_3").val());
            var f41 = ($("#txt_Finish_item4_1").val() == "") ? 0 : parseFloat($("#txt_Finish_item4_1").val());
            var f42 = ($("#txt_Finish_item4_2").val() == "") ? 0 : parseFloat($("#txt_Finish_item4_2").val());
            var f43 = ($("#txt_Finish_item4_3").val() == "") ? 0 : parseFloat($("#txt_Finish_item4_3").val());
            var f51 = ($("#txt_Finish_item5_1").val() == "") ? 0 : parseFloat($("#txt_Finish_item5_1").val());
            var f52 = ($("#txt_Finish_item5_2").val() == "") ? 0 : parseFloat($("#txt_Finish_item5_2").val());
            var f53 = ($("#txt_Finish_item5_3").val() == "") ? 0 : parseFloat($("#txt_Finish_item5_3").val());
            $("#txt_Finish_item1_all").val(parseFloat(f11 + f12 + f13).toFixed(1));
            $("#txt_Finish_item2_all").val(parseFloat(f21 + f22 + f23).toFixed(0));
            $("#txt_Finish_item3_all").val(parseFloat(f31 + f32 + f33).toFixed(0));
            $("#txt_Finish_item4_all").val(parseFloat(f41 + f42 + f43).toFixed(0));
            $("#txt_Finish_item5_all").val(parseFloat(f51 + f52 + f53).toFixed(0));
            //txt_Finish_item1_1 txt_Finish_item1_all
        }
        //儲存
        function mod_project(thisid) {
            if ($("#txt_alldates").val() == "" || $("#txt_alldatee").val() == "") {
                alert("執行期程尚未設定,請聯絡系統管理員");
                return;
            }
            var che_oneself = $("#chk_Other_Oneself").is(":checked") ? "Y" : "";
            var che_other = $("#chk_Other_Other").is(":checked") ? "Y" : "";
            if (che_oneself != "Y" && $("#txt_Other_Oneself_Money").val() != "" && $("#txt_Other_Oneself_Money").val() != "0.0") {
                alert("請確認其他經費來源-自籌款");
                return;
            }
            if (che_other != "Y" && (($("#txt_Other_Other_Money").val() != "" && $("#txt_Other_Other_Money").val() != "0.0") || $("#txt_Other_Other_name").val() != "")) {
                alert("請確認其他經費來源-其他機關補助");
                return;
            }
            
            $.ajax({
                type: "POST",
                async: true, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/ProjectInfo.ashx",
                data: {
                    func: "mod_project",
                    str_tmpguid:$("#<%= tmpguid.ClientID %>").val(),
                    str_I_ID: $("#hidden_I_ID").val(),
                    str_M_ID: $("#hidden_M_ID").val(),
                    str_I_Guid: $("#hidden_I_Guid").val(),
                    str_1_sdate: $("#txt_1_sdate").val(),
                    str_1_edate: $("#txt_1_edate").val(),
                    str_2_sdate: $("#txt_2_sdate").val(),
                    str_2_edate: $("#txt_2_edate").val(),
                    str_3_sdate: $("#txt_3_sdate").val(),
                    str_3_edate: $("#txt_3_edate").val(),
                    str_money_item1_1: $("#txt_money_item1_1").val(),
                    str_money_item1_2: $("#txt_money_item1_2").val(),
                    str_money_item1_3: $("#txt_money_item1_3").val(),
                    str_money_item1_sum: $("#txt_money_item1_sum").val(),
                    str_money_item2_1: $("#txt_money_item2_1").val(),
                    str_money_item2_2: $("#txt_money_item2_2").val(),
                    str_money_item2_3: $("#txt_money_item2_3").val(),
                    str_money_item2_sum: $("#txt_money_item2_sum").val(),
                    str_money_item3_1: $("#txt_money_item3_1").val(),
                    str_money_item3_2: $("#txt_money_item3_2").val(),
                    str_money_item3_3: $("#txt_money_item3_3").val(),
                    str_money_item3_sum: $("#txt_money_item3_sum").val(),
                    str_money_item4_1: $("#txt_money_item4_1").val(),
                    str_money_item4_2: $("#txt_money_item4_2").val(),
                    str_money_item4_3: $("#txt_money_item4_3").val(),
                    str_money_item4_sum: $("#txt_money_item4_sum").val(),
                    str_Other_Oneself: che_oneself,
                    str_Other_Oneself_Money: $("#txt_Other_Oneself_Money").val(),
                    str_Other_Other: che_other,
                    str_Other_Other_name: $("#txt_Other_Other_name").val(),
                    str_Other_Other_Money: $("#txt_Other_Other_Money").val(),
                    str_Target: $("#txt_Target").val(),
                    str_summary: $("#txt_summary").val(),
                    str_Finish_item1_1: $("#txt_Finish_item1_1").val(),
                    str_Finish_item1_2: $("#txt_Finish_item1_2").val(),
                    str_Finish_item1_3: $("#txt_Finish_item1_3").val(),
                    str_Finish_item1_all: $("#txt_Finish_item1_all").val(),
                    str_Finish_item2_1: $("#txt_Finish_item2_1").val(),
                    str_Finish_item2_2: $("#txt_Finish_item2_2").val(),
                    str_Finish_item2_3: $("#txt_Finish_item2_3").val(),
                    str_Finish_item2_all: $("#txt_Finish_item2_all").val(),
                    str_Finish_item3_1: $("#txt_Finish_item3_1").val(),
                    str_Finish_item3_2: $("#txt_Finish_item3_2").val(),
                    str_Finish_item3_3: $("#txt_Finish_item3_3").val(),
                    str_Finish_item3_all: $("#txt_Finish_item3_all").val(),
                    str_Finish_item4_1: $("#txt_Finish_item4_1").val(),
                    str_Finish_item4_2: $("#txt_Finish_item4_2").val(),
                    str_Finish_item4_3: $("#txt_Finish_item4_3").val(),
                    str_Finish_item4_all: $("#txt_Finish_item4_all").val(),
                    str_Finish_item5_1: $("#txt_Finish_item5_1").val(),
                    str_Finish_item5_2: $("#txt_Finish_item5_2").val(),
                    str_Finish_item5_3: $("#txt_Finish_item5_3").val(),
                    str_Finish_item5_all: $("#txt_Finish_item5_all").val(),
                    str_savetype: thisid
                    //str_selPeople: $("#sel_people").val()
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                }, beforeSend: function () {
                    //$.blockUI({ message: '處理中，請稍待...' });
                },
                complete: function () {
                    //$.unblockUI();
                },
                success: function (data) {
                    if (data == "timeout") {
                        alert("請重新登入");
                        location.href = "Login.aspx";
                    } else if (data == "nodata") {
                        alert("查無填寫人員資料");
                    } else if (data.indexOf("Error") >-1) {
                        alert(data);
                    }
                    else {
                        if (thisid == "btn_save") {//儲存按鈕
                            alert("儲存成功");
                        } else if (thisid == "auto") {//自動存檔
                            
                        } else {//下一步
                            location.href = "CheckPoint.aspx?v=" + $("#hidden_M_ID").val() + "";
                        }
                    }
                }
            });//ajax end
        }
        ////開窗 挑選承辦人//管理者代填的功能拿掉20171201
        //function openfancybox(item) {
        //    switch ($(item).attr("id")) {
        //        case "Pbox"://人事異動 人員挑選
        //            link = "getContractor.aspx";
        //            fbox(link);
        //            break;
        //    }

        //}
        //function fbox(link) {
        //    $.fancybox({
        //        href: link,
        //        type: "iframe",
        //        minHeight: "400",
        //        closeClick: false,
        //        openEffect: 'elastic',
        //        closeEffect: 'elastic',
        //        closeBtn: false,//不顯示X鈕
        //        helpers: {
        //            overlay: { closeClick: false } // 點旁邊遮罩也不能關閉fancybox
        //        }
        //    });
        //}

        ////開窗 挑選承辦人 回傳的值//管理者代填的功能拿掉20171201
        //function returnValue(jv) {
        //    var jsonvalue = $.parseJSON(jv);
        //    var p_id = jsonvalue.id;//承辦人ID
        //    var p_guid = jsonvalue.guid;//承辦人GUID
        //    var p_name = jsonvalue.name;//承辦人姓名
        //    var p_city = jsonvalue.city;//承辦人機關代碼
        //    var p_cityname = jsonvalue.cityname;//承辦人機關名稱
        //    var p_office = jsonvalue.office;//承辦人局處
        //    var p_aIGuid = jsonvalue.aIGuid;//承辦人的計畫資料GUID
        //    if (p_aIGuid != "") {
        //        alert("該承辦員已經有填寫過計畫基本資料，請從計畫列表挑選進入修改");
        //        $("#sel_people").val("");
        //        $("#sel_people_name").empty();
        //        $("#div_city").empty();
        //        $("#div_office").empty();
        //        $("#hidden_M_ID").val("");
        //        $(".Jdatepicker").val("");
        //        $("#txt_alldates,#txt_alldatee,#txt_allmonth").val("");
        //    } else {
        //        $("#sel_people").val(p_id);
        //        $("#sel_people_name").empty().append(p_name);
        //        $("#div_city").empty().append(p_cityname);
        //        $("#div_office").empty().append(p_office);
        //        $("#hidden_M_ID").val(p_id);
        //        loaddate(p_city);
        //    }
            
        //}
        //隱藏欄位
        function disabledinput(){
            $("input[type='text']").attr("disabled", "disabled");
            $("input[type='checkbox']").attr("disabled", "disabled");
            $(".genbtnS,.genbtn").attr("disabled", "disabled");
            $(".genbtnS,.genbtn").css("display", "none");
            $("textarea").attr("disabled", "disabled");
        }
        //撈期程資料
        function loaddate(cityno) {
            $.ajax({
                type: "POST",
                async: true, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/ProjectInfo.ashx",
                data: {
                    func: "load_date",
                    p_city: cityno//人員ID
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if (data == "timeout") {
                        alert("請重新登入");
                        location.href = "Login.aspx";
                    } else if (data == "nodata") {
                        alert("查無人員相關資料");
                    } else if (data == "error") {
                        alert("error");
                    } else if (data == "DiffCity") {
                        alert("您沒有權限進入此頁面");
                        window.location = "ProjectList.aspx";
                    } else if (data == "nodate") {
                        alert("執行期程尚未設定,請聯絡系統管理員");
                        $("#txt_1_sdate,#txt_3_edate,#txt_alldates,#txt_alldatee,#txt_allmonth").val("");
                        $("#sel_people,#hidden_M_ID").val("");
                        $("#sel_people_name").empty();
                        $("#div_city").empty();
                        $("#div_office").empty();
                    }
                    else {
                        $("#txt_1_sdate").val(data[0].PD_StartDate);
                        $("#txt_3_edate").val(data[0].PD_EndDate);
                        $("#txt_alldates").val(data[0].PD_StartDate);
                        $("#txt_alldatee").val(data[0].PD_EndDate);
                        $("#txt_allmonth").val(monthDiff(new Date(data[0].PD_StartDate), new Date(data[0].PD_EndDate)));
                    }
                }
            });//ajax end
        }

        //判斷要顯示什麼不顯示什麼
        function showorhide(strType, Iguid, People, user, competence, flag, chk_flag, M_ID, M_Name) {
            if (strType == "old")
            {//修改
                if (chk_flag != "" && chk_flag != null && competence=="SA") {
                    //定稿 只有管理者可以動
                    $(".mfile").show();
                    $(".pdfile").show();
                    $("#MoneyUpBtn,#PlanDescUpBtn").show();
                    dateDiff("txt_1_edate");
                    dateDiff("txt_2_edate");
                    dateDiff("txt_3_sdate");
                    $("#btn_next").html("下一步");
                }
                if (chk_flag != "" && chk_flag != null && competence != "SA") {
                    //定稿 承辦人&承辦主管都不能動
                    $("#btn_next").html("下一頁");
                    disabledinput();
                }
                if (flag != "Y" && competence == "SA") {
                    //管理者不能動未定稿的資料
                    $("#btn_next").html("下一頁");
                    disabledinput();
                }
                if (flag != "Y" && competence == "01") {
                    $(".mfile").show();
                    $(".pdfile").show();
                    $("#MoneyUpBtn,#PlanDescUpBtn").show();
                }
            } else {//新增
                if (chk_flag != "" && chk_flag!=null) {
                    //該機關底下已經有定稿
                    //什麼都不能做 只能下一頁
                    $("#btn_next").html("下一頁");
                    disabledinput();
                } else {
                    if (competence == "01") {
                        //顯示上傳按鈕及上傳附件
                        $(".mfile").show();
                        $(".pdfile").show();
                        $("#MoneyUpBtn,#PlanDescUpBtn").show();
                        dateDiff("txt_1_edate");
                        dateDiff("txt_2_edate");
                        dateDiff("txt_3_sdate");
                    } else {
                        alert("您沒有權限新增");
                        location.href = "ProjectList.aspx";
                    }
                    
                }
            }
        }

        //數字欄位 如果為空就給0 或 0.0
        function changeZero() {
            $("input").each(function () {
                if ($(this).attr("t") == "strint" && ($(this).val() == "" || $(this).val() == "0" || $(this).val() == "0.0")) {
                    $(this).val("0");
                }
                if ($(this).attr("t") == "strfloat" && ($(this).val() == "" || $(this).val() == "0" || $(this).val() == "0.0")) {
                    $(this).val("0.0");
                }
                if ($(this).attr("t") == "strfloat3" && ($(this).val() == "" || $(this).val() == "0" || $(this).val() == "0.000")) {
                    $(this).val("0.000");
                }
            });

        }
    </script>
</asp:Content>

