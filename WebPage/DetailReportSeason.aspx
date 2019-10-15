<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="DetailReportSeason.aspx.cs" Inherits="WebPage_DetailReportSeason" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">季報</span></div>
            <div class="right">進度報表 / 季報</div>
        </div>
        <!-- twocol -->
        <div style="text-align:right; margin-top:10px;">
            <a href="javascript:void(0);" class="genbtn" id="exbtn" style="display: none;">匯出</a>
        </div>

        <div class="OchiTrasTable width100 TitleLength08">
            <!-- 提報年分  提報季別-->
            <div class="OchiRow">
                <div class="OchiThird">
                    <div class="OchiCell OchiTitle TitleSetWidth">提報年份</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <span id="rYear"></span>
                        </div>
                        <!-- cell內容end -->
                    </div>
                    <!-- OchiCell -->
                </div>
                <div class="OchiThird">
                    <div class="OchiCell OchiTitle TitleSetWidth">提報季別</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <span id="rSeason"></span>
                        </div>
                        <!-- cell內容end -->
                    </div>
                    <!-- OchiCell -->
                </div>
                <div class="OchiThird">
                    <div class="OchiCell OchiTitle TitleSetWidth">期程</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <span id="rStage"></span>
                        </div>
                        <!-- cell內容end -->
                    </div>
                    <!-- OchiCell -->
                </div>
            </div>
            <!-- 雙欄 執行機關 承辦局處 -->
            <div class="OchiRow">
                <div class="OchiThird">
                    <div class="OchiCell OchiTitle TitleSetWidth">執行機關</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <div class="OchiCellInner width100">
                                <span id="rsCity"></span>
                            </div>
                        </div>
                        <!-- OchiTableInner -->
                        <!-- cell內容end -->
                    </div>
                </div>

                <div class="OchiThird">
                    <div class="OchiCell OchiTitle TitleSetWidth">承辦局處</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <div class="OchiCellInner width100">
                                <span id="rsOffice"></span>
                            </div>
                        </div>
                        <!-- OchiTableInner -->
                        <!-- cell內容end -->
                    </div>
                </div>
            </div>
            <!-- OchiRow -->

            <!-- 單欄 執行期程 -->
            <div class="OchiRow">
                <div class="OchiHalf">
                    <div class="OchiCell OchiTitle TitleSetWidth">執行期程</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <div class="OchiCellInner nowrap textcenter">開始:</div>
                            <div class="OchiCellInner width20"><span id="rsStartDate"></span></div>
                            <div class="OchiCellInner nowrap textcenter">&nbsp;~&nbsp;</div>
                            <div class="OchiCellInner nowrap textcenter">結束:</div>
                            <div class="OchiCellInner width20"><span id="rsEndDate"></span></div>
                        </div>
                        <!-- OchiTableInner -->
                        <!-- cell內容end -->
                    </div>
                    <!-- OchiCell -->
                </div>
                <div class="OchiHalf">
                    <div class="OchiCell OchiTitle TitleSetWidth">合計：</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <div class="OchiCellInner width100"><span id="rsTotalMonth"></span>月</div>
                        </div>
                        <!-- OchiTableInner -->
                        <!-- cell內容end -->
                    </div>
                </div>
            </div>
            <!-- OchiRow -->

            <div class="twocol margin5T margin5B" id="div_nodata"></div>
            <!--季報資料-->
            <!-- 單欄 預算狀態-->
            <div class="OchiRow">
                <!--<div class="font-size3 margin10T" style="font-size: 16px;">預算狀態</div>-->
                <div class="font-bold margin5B" style="font-size: 16pt;">預算狀態</div>
                <div class="margin5B" id="RS_CostDesc"></div>
                備註:如議會已納入預算
            </div>
            <!-- OchiRow -->
            <!-- 單欄 經費實支數-->
            <div class="OchiRow">
                <!--<div class="font-size3 margin10T" style="font-size: 16px;">經費實支數</div>-->
                <div class="font-bold margin5T" style="font-size: 16pt;">經費實支數</div>
                <div class="stripeMe margin5T font-normal">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <thead>
                            <tr>
                                <th nowrap="nowrap" style="width: 300px;">推動工作</th>
                                <th nowrap="nowrap" style="width: 200px;">經費(千元) A</th>
                                <th nowrap="nowrap" style="width: 200px;">實支數(千元) B</th>
                                <th nowrap="nowrap">實支率(%) C = B / A</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td style="text-align: center;">節電基礎工作</td>
                                <td style="text-align: right;"><span id="rsMoney1"></span></td>
                                <td style="text-align: right;"><span id="RS_Type01Real"></span></td>
                                <td style="text-align: right;"><span id="rsMoneyRatio1"></span></td>
                            </tr>
                            <tr>
                                <td style="text-align: center;">因地制宜</td>
                                <td style="text-align: right;"><span id="rsMoney2"></span></td>
                                <td style="text-align: right;"><span id="RS_Type02Real"></span></td>
                                <td style="text-align: right;"><span id="rsMoneyRatio2"></span></td>
                            </tr>
                            <tr>
                                <td style="text-align: center;">設備汰換及智慧用電</td>
                                <td style="text-align: right;"><span id="rsMoney3"></span></td>
                                <td style="text-align: right;"><span id="RS_Type03Real"></span></td>
                                <td style="text-align: right;"><span id="rsMoneyRatio3"></span></td>
                            </tr>
                            <tr>
                                <td style="text-align: center;">擴大補助</td>
                                <td style="text-align: right;"><span id="rsMoney4"></span></td>
                                <td style="text-align: right;"><span id="RS_Type04Real"></span></td>
                                <td style="text-align: right;"><span id="rsMoneyRatio4"></span></td>
                            </tr>
                            <tr>
                                <td style="text-align: center;">整體工作</td>
                                <td style="text-align: right;"><span id="rsAllMoney"></span></td>
                                <td style="text-align: right;"><span id="rsAllRealMoney"></span></td>
                                <td style="text-align: right;"><span id="rsAllMoneyRatio"></span></td>
                            </tr>
                        </tbody>
                    </table>
                </div>
                <div class="textright margin5B">單位:仟元</div>
            </div>
            <!-- OchiRow -->
            <!-- 單欄 整體進度-->
            <div class="OchiRow">
                <div class="font-bold margin5T" style="font-size: 16pt;">整體進度</div>
                <!-- cell內容start -->
                <div class="OchiTableInner width100 margin10TB" style="margin-left: 50px;">
                    <div class="OchiCellInner nowrap">預定執行進度：</div>
                    <div class="OchiCellInner width50"><span id="allProcess"></span>%</div>
                    <div class="OchiCellInner nowrap">實際執行進度：</div>
                    <div class="OchiCellInner width50"><span id="allRealProcess"></span>%</div>
                </div>
                <!-- OchiTableInner -->
                <!-- cell內容end -->
            </div>

            <!--節電基礎工作-->
            <div class="OchiRow">
                <div class="margin10T" style="font-size: 16pt; font-weight: bolder;">
                    壹、節電基礎工作&nbsp;<span style="color: blue;">(本季預定：<span id="Process_01"></span>%；本季實際 <span id="RealProcess_01"></span>%)</span>
                </div>
                <!--查核點動態資訊-->
                <div id="div_type01"></div>
                <!--檔案上傳
                <div class="twocol margin10TB">
                    <div class="right"><a href="javascript:void(0);" class="genbtnS" id="btn_up03">參考資料上傳</a></div>
                </div>-->
                <div class="mfile" style="margin-top: 10px; display: none;">附件檔</div>
                <div class="stripeMe mfile" style="margin-bottom: 10px; display: none;">
                    <table id="bwFileList" width="100%" border="0" cellspacing="0" cellpadding="0">
                        <thead>
                            <tr>
                                <th>檔案名稱</th>
                            </tr>
                        </thead>
                        <tr>
                            <td>查詢無資料</td>
                        </tr>
                    </table>
                </div>
            </div>

            <!--因地制宜-->
            <div class="OchiRow">
                <div class="margin10T" style="font-size: 16pt; font-weight: bolder;">
                    貳、因地制宜&nbsp;<span style="color: blue;">(本季預定：<span id="Process_02"></span>%；本季實際 <span id="RealProcess_02"></span>%)</span>
                </div>
                <!--查核點動態資訊-->
                <div id="div_type02"></div>
                <!--檔案上傳
                <div class="twocol margin10TB">
                    <div class="right"><a href="javascript:void(0);" class="genbtnS" id="btn_up04">參考資料上傳</a></div>
                </div>-->
                <div class="mfile" style="margin-top: 10px; display: none;">附件檔</div>
                <div class="stripeMe mfile" style="margin-bottom: 10px; display: none;">
                    <table id="pFileList" width="100%" border="0" cellspacing="0" cellpadding="0">
                        <thead>
                            <tr>
                                <th>檔案名稱</th>
                            </tr>
                        </thead>
                        <tr>
                            <td>查詢無資料</td>
                        </tr>
                    </table>
                </div>
            </div>

            <!--設備汰換與智慧用電-->
            <div class="OchiRow">
                <div class="margin10T" style="font-size: 16pt; font-weight: bolder;">
                    參、設備汰換與智慧用電&nbsp;<span style="color: blue;">(本季預定：<span id="Process_03"></span>%；本季實際 <span id="RealProcess_03"></span>%)</span>
                </div>
                <!--查核點動態資訊-->
                <div id="div_type03"></div>
                <!--檔案上傳
                <div class="twocol margin10TB">
                    <div class="right"><a href="javascript:void(0);" class="genbtnS" id="btn_up05">參考資料上傳</a></div>
                </div>-->
                <div class="mfile" style="margin-top: 10px; display: none;">附件檔</div>
                <div class="stripeMe mfile" style="margin-bottom: 10px; display: none;">
                    <table id="sFileList" width="100%" border="0" cellspacing="0" cellpadding="0">
                        <thead>
                            <tr>
                                <th>檔案名稱</th>
                            </tr>
                        </thead>
                        <tr>
                            <td>查詢無資料</td>
                        </tr>
                    </table>
                </div>

                <div class="stripeMe margin10TB font-normal">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <thead>
                            <tr>
                                <th nowrap="nowrap">項目</th>
                                <th nowrap="nowrap" style="width: 200px;">本期預計完成數</th>
                                <th nowrap="nowrap" style="width: 200px;">本期累計完成數</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td style="text-align: center;">無風管冷氣(kW)註：每台冷氣約4kW</td>
                                <td style="text-align: right;"><span id="DR_Finish1"></span></td>
                                <td style="text-align: right;"><span id="RS_03Type01C"></span></td>
                            </tr>
                            <tr>
                                <td style="text-align: center;">老舊辦公室照明(具)</td>
                                <td style="text-align: right;"><span id="DR_Finish2"></span></td>
                                <td style="text-align: right;"><span id="RS_03Type02C"></span></td>
                            </tr>
                            <tr>
                                <td style="text-align: center;">室內停車場智慧照明(盞)</td>
                                <td style="text-align: right;"><span id="DR_Finish3"></span></td>
                                <td style="text-align: right;"><span id="RS_03Type03C"></span></td>
                            </tr>
                            <tr>
                                <td style="text-align: center;">中型能管系統(套)</td>
                                <td style="text-align: right;"><span id="DR_Finish4"></span></td>
                                <td style="text-align: right;"><span id="RS_03Type04C"></span></td>
                            </tr>
                            <tr>
                                <td style="text-align: center;">大型能管系統(套)</td>
                                <td style="text-align: right;"><span id="DR_Finish5"></span></td>
                                <td style="text-align: right;"><span id="RS_03Type05C"></span></td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>

            
            <!--擴大補助-->
            <div class="OchiRow">
                <div class="margin10T" style="font-size: 16pt; font-weight: bolder;">
                    肆、擴大補助&nbsp;<span style="color: blue;">(本季預定：<span id="Process_04"></span>%；本季實際 <span id="RealProcess_04"></span>%)</span>
                </div>
                <!--查核點動態資訊-->
                <div id="div_type04"></div>
                <!--檔案上傳
                <div class="twocol margin10TB">
                    <div class="right"><a href="javascript:void(0);" class="genbtnS" id="btn_up05">參考資料上傳</a></div>
                </div>-->
                <div class="mfile" style="margin-top: 10px; display: none;">附件檔</div>
                <div class="stripeMe mfile" style="margin-bottom: 10px; display: none;">
                    <table id="exFileList" width="100%" border="0" cellspacing="0" cellpadding="0">
                        <thead>
                            <tr>
                                <th>檔案名稱</th>
                            </tr>
                        </thead>
                        <tr>
                            <td>查詢無資料</td>
                        </tr>
                    </table>
                </div>
                
                <div class="stripeMe margin10TB font-normal">
                    <table id="exFinishTab" width="100%" border="0" cellspacing="0" cellpadding="0">
                        <thead>
                            <tr>
                                <th nowrap="nowrap">項目</th>
                                <th nowrap="nowrap" style="width:200px;">本期預計完成數</th>
                                <th nowrap="nowrap" style="width:200px;">本期累計完成數</th>
                            </tr>
                        </thead>
                        <tbody></tbody>
                    </table>
                </div>
            </div>

            <!-- 頁尾 -->
            <!-- 承辦人資料 -->
            <div class="OchiRow">
                <div class="OchiThird">
                    <div class="OchiCell OchiTitle TitleSetWidth">填表人</div>
                    <div class="OchiCell width100"><span id="UserName"></span></div>
                </div>
                <!-- OchiThird -->
                <div class="OchiThird">
                    <div class="OchiCell OchiTitle TitleSetWidth">電話</div>
                    <div class="OchiCell width100"><span id="UserTel"></span></div>
                </div>
                <!-- OchiThird -->
                <div class="OchiThird">
                    <div class="OchiCell OchiTitle TitleSetWidth">填表日期</div>
                    <div class="OchiCell width100"><span id="WriteDate"></span></div>
                </div>
                <!-- OchiThird -->
            </div>
            <!-- 承辦主管資料 -->
            <div class="OchiRow">
                <div class="OchiHalf">
                    <div class="OchiCell OchiTitle TitleSetWidth">主管</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <div class="OchiCellInner width100"><span id="ManagerName"></span></div>
                        </div>
                        <!-- cell內容end -->
                    </div>
                </div>
                <!-- OchiHalf -->
                <div class="OchiHalf">
                    <div class="OchiCell OchiTitle TitleSetWidth">主管簽核日期</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <div class="OchiCellInner width100"><span id="ManagerCheckDate"></span></div>
                        </div>
                        <!-- cell內容end -->
                    </div>
                </div>
                <!-- OchiHalf -->
            </div>
        </div>
        <!--div_data end-->

        <div class="twocol margin15T margin5B">
            <div style="text-align: center;">
                <a href="javascript:void(0);" class="genbtn pbtn" id="btn_ok" style="display: none;">通過</a>&nbsp;&nbsp;&nbsp;
                <a href="javascript:void(0);" class="genbtn pbtn" id="btn_notok" style="display: none;">不通過</a>
            </div>
        </div>

        <!--審核不通過 跳出fancybox輸入意見-->
        <div id="opblock" style="display: none; text-align: center;">
            <div style="margin-bottom: 10px;">請輸入審核不通過意見</div>
            <div style="margin-bottom: 10px;">
                <textarea id="textOp" rows="7" cols="50"></textarea>
            </div>
            <input type="button" id="submitbtn" value="確定" class="genbtn" />
        </div>
        <!--匯出-->
        <div id="exblock" style="display: none; text-align: center;">
            <div style="margin-bottom: 10px;">請選擇匯出檔案類型</div>
            <div style="margin-bottom: 10px;">
                <input type="radio" name="extype" value="word" checked="checked" />&nbsp;Word&nbsp;&nbsp;
                <input type="radio" name="extype" value="pdf" />&nbsp;PDF
            </div>
            <input type="button" id="expbtn" value="確定" class="genbtn" />
        </div>

        <input type="hidden" id="tmp_rsid" />
        <input type="hidden" id="hidden_reportGuid" />
        <input type="hidden" id="hidden_year" />
        <input type="hidden" id="hidden_season" />
        <input type="hidden" id="hidden_typeval" />
        <input type="hidden" id="hidden_rsguid" />
        <!-- twocol -->
    </div>
    <script>
        $(function () {
            $(".pbtn").hide();
            $("#rYear").html($.getParamValue('year') + " 年");
            $("#rStage").html("第 " + $.getParamValue('stage') + " 期");
            getInfo();

            //附件檔
            getFileList("03", "#bwFileList");
            getFileList("04", "#pFileList");
            getFileList("05", "#sFileList");
            getFileList("07", "#exFileList");

            //下載檔案
            $(document).on("click", "a[name='downloadbtn']", function () {
                var id = $(this).attr("fid");
                location.href = "../DOWNLOAD.aspx?v=" + id;
            });
            //通過 不通過按鈕
            $("#btn_ok,#btn_notok").click(function () {
                var str = "";
                var this_id = $(this).attr("id");
                if (this_id == "btn_ok") {
                    str = "通過";
                }
                if (this_id == "btn_notok") {
                    str = "不通過";
                }
                if (confirm("該季報是否[" + str + "]?")) {
                    goCheck(this_id);
                }

            });

            //匯出 button
            $(document).on("click", "#exbtn", function () {
                openExport();
            });

            //確認匯出
            $(document).on("click", "#expbtn", function () {
                var type = $("input[name='extype']:checked").val();
                window.open("../handler/ExportReportSeason.aspx?v=" + $("#tmp_rsid").val() + "&tp=" + type);
                $.fancybox.close();
            });

            //審核不通過確定按鈕(發送MAIL)
            $("#submitbtn").click(function () {
                checkUpdate($("#hidden_typeval").val());
            });

        });

        function getInfo() {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/getHistoryDetail_S.aspx",
                data: {
                    RS_ID: $.getParamValue('v')
                    //RS_ID: $.getParamValue('year')
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if ($(data).find("Error").length > 0) {
                        alert($(data).find("Error").attr("Message"));
                    } else if ($(data).find("timeout").length > 0) {
                        alert("請重新登入");
                        location.href = "Login.aspx";
                    }
                    else {
                        var chkType = '', chkStatus = '';
                        //季報資料
                        if ($(data).find("data_item").length > 0) {
                            $(data).find("data_item").each(function (i) {
                                $("#tmp_rsid").val($(this).children("RS_ID").text().trim());
                                $("#hidden_rsguid").val($(this).children("RS_Guid").text().trim());
                                $("#RS_CostDesc").html(replaceLine($(this).children("RS_CostDesc").text().trim()));
                                $("#RS_Type01Real").html($(this).children("RS_Type01Real").text().trim());
                                $("#RS_Type02Real").html($(this).children("RS_Type02Real").text().trim());
                                $("#RS_Type03Real").html($(this).children("RS_Type03Real").text().trim());
                                $("#RS_Type04Real").html($(this).children("RS_Type04Real").text().trim());
                                $("#RS_03Type01C").html($(this).children("RS_03Type01C").text().trim());
                                $("#RS_03Type02C").html($(this).children("RS_03Type02C").text().trim());
                                $("#RS_03Type03C").html($(this).children("RS_03Type03C").text().trim());
                                $("#RS_03Type04C").html($(this).children("RS_03Type04C").text().trim());
                                $("#RS_03Type05C").html($(this).children("RS_03Type05C").text().trim());
                                $("#hidden_reportGuid").val($(this).children("RS_Guid").text().trim());
                                chkType = $(this).children("RC_CheckType").text().trim();
                                chkStatus = $(this).children("RC_Status").text().trim();
                                //基本資料 rsTotalMonth
                                $("#rYear").html($(this).children("RS_Year").text().trim() + "年");
                                $("#rSeason").html("第" + $(this).children("RS_Season").text().trim() + "季");
                                $("#rStage").html("第" + $(this).children("RS_Stage").text().trim() + "期");
                                $("#rsCity").html($(this).children("CityName").text().trim());
                                $("#rsOffice").html($(this).children("M_Office").text().trim());
                                var StartDate = $(this).children("RS_StartDay").text().trim();
                                var EndDate = $(this).children("RS_EndDay").text().trim();
                                $("#rsStartDate").html($(this).children("RS_StartDay").text().trim());
                                $("#rsEndDate").html($(this).children("RS_EndDay").text().trim());
                                $("#rsTotalMonth").html(countMonth(StartDate, EndDate));
                                $("#rsMoney1").append($(this).children("RS_Type01Money").text().trim());
                                $("#rsMoney2").html($(this).children("RS_Type02Money").text().trim());
                                $("#rsMoney3").html($(this).children("RS_Type03Money").text().trim());
                                $("#rsMoney4").html($(this).children("RS_Type04Money").text().trim());
                                $("#RS_Type01Real").html($(this).children("RS_Type01Real").text().trim());
                                $("#RS_Type02Real").html($(this).children("RS_Type02Real").text().trim());
                                $("#RS_Type03Real").html($(this).children("RS_Type03Real").text().trim());
                                $("#RS_Type04Real").html($(this).children("RS_Type04Real").text().trim());
                                $("#rsMoneyRatio1").html($(this).children("RS_Type01RealRate").text().trim());
                                $("#rsMoneyRatio2").html($(this).children("RS_Type02RealRate").text().trim());
                                $("#rsMoneyRatio3").html($(this).children("RS_Type03RealRate").text().trim());
                                $("#rsMoneyRatio4").html($(this).children("RS_Type04RealRate").text().trim());
                                var rs4 = ($("#rsMoney4")[0].innerText == "") ? 0 : parseFloat($("#rsMoney4")[0].innerText);
                                var allMoney =(parseFloat($("#rsMoney1")[0].innerText) + parseFloat($("#rsMoney2")[0].innerText) + parseFloat($("#rsMoney3")[0].innerText) + rs4).toFixed(3);
                                $("#rsAllMoney").html(allMoney);
                                countMoney();

                                //整體預定&實際進度
                                $("#allProcess").html(Number($(this).children("RS_AllSchedule").text().trim()).toFixed(2));
                                $("#allRealProcess").html(Number($(this).children("RS_AllRealSchedule").text().trim()).toFixed(2));
                                //三大項分別的預定&實際進度
                                //Process_01 RealProcess_01
                                $("#Process_01").html(Number($(this).children("RS_01Schedule").text().trim()).toFixed(2));
                                $("#RealProcess_01").html(Number($(this).children("RS_01RealSchedule").text().trim()).toFixed(2));
                                $("#Process_02").html(Number($(this).children("RS_02Schedule").text().trim()).toFixed(2));
                                $("#RealProcess_02").html(Number($(this).children("RS_02RealSchedule").text().trim()).toFixed(2));
                                $("#Process_03").html(Number($(this).children("RS_03Schedule").text().trim()).toFixed(2));
                                $("#RealProcess_03").html(Number($(this).children("RS_03RealSchedule").text().trim()).toFixed(2));
                                $("#Process_04").html(Number($(this).children("RS_04Schedule").text().trim()).toFixed(2));
                                $("#RealProcess_04").html(Number($(this).children("RS_04RealSchedule").text().trim()).toFixed(2));

                                //設備汰換與智慧用電 五大項預計&時計完成數
                                $("#DR_Finish1").html($(this).children("I_Finish_item1_" + $.getParamValue('stage')).text().trim());
                                $("#DR_Finish2").html($(this).children("I_Finish_item2_" + $.getParamValue('stage')).text().trim());
                                $("#DR_Finish3").html($(this).children("I_Finish_item3_" + $.getParamValue('stage')).text().trim());
                                $("#DR_Finish4").html($(this).children("I_Finish_item4_" + $.getParamValue('stage')).text().trim());
                                $("#DR_Finish5").html($(this).children("I_Finish_item5_" + $.getParamValue('stage')).text().trim());

                                //頁尾人員基本資料
                                $("#UserName").html($(this).children("M_Name").text().trim());
                                $("#UserTel").html($(this).children("M_Tel").text().trim());
                                $("#WriteDate").html(formatDate($(this).children("RC_CreateDate").text().trim()));
                                $("#ManagerName").html($(this).children("BossName").text().trim());
                                if ($(this).children("RC_CheckDate").text().trim() != "") {
                                    $("#ManagerCheckDate").html($.datepicker.formatDate('yy/mm/dd', new Date($(this).children("RC_CheckDate").text().trim())));
                                }

                                //三大項目動態資訊
                                var cpXml = $.parseXML($(this).children("RS_CheckPointData").text().trim());
                                var pdXml = $.parseXML($(this).children("RS_PushItemDesc").text().trim());
                                if (pdXml == null) {
                                    $("#div_type01").append(getCheckPoint_Old(cpXml, "PushItem[P_Type='01']"));
                                    $("#div_type02").append(getCheckPoint_Old(cpXml, "PushItem[P_Type='02']"));
                                    $("#div_type03").append(getCheckPoint_Old(cpXml, "PushItem[P_Type='03']"));
                                    $("#div_type04").append(getCheckPoint_Old(cpXml, "PushItem[P_Type='04']"));
                                }
                                else {
                                    $("#div_type01").append(getCheckPoint(cpXml, pdXml, "PushItem[P_Type='01']"));
                                    $("#div_type02").append(getCheckPoint(cpXml, pdXml, "PushItem[P_Type='02']"));
                                    $("#div_type03").append(getCheckPoint(cpXml, pdXml, "PushItem[P_Type='03']"));
                                    $("#div_type04").append(getCheckPoint(cpXml, pdXml, "PushItem[P_Type='04']"));
                                }

                                //擴大補助預計完成數
                                CreateExFinishTable(data,cpXml);

                                //設備汰換 預期及實際完成數量 DR_Finish1 RS_03Type01C
                                $("#DR_Finish1").html($(this).children("RS_03Type01S").text().trim());
                                $("#RS_03Type01C").html($(this).children("RS_03Type01C").text().trim());
                                $("#DR_Finish2").html($(this).children("RS_03Type02S").text().trim());
                                $("#RS_03Type02C").html($(this).children("RS_03Type02C").text().trim());
                                $("#DR_Finish3").html($(this).children("RS_03Type03S").text().trim());
                                $("#RS_03Type03C").html($(this).children("RS_03Type03C").text().trim());
                                $("#DR_Finish4").html($(this).children("RS_03Type04S").text().trim());
                                $("#RS_03Type04C").html($(this).children("RS_03Type04C").text().trim());
                                $("#DR_Finish5").html($(this).children("RS_03Type05S").text().trim());
                                $("#RS_03Type05C").html($(this).children("RS_03Type05C").text().trim());
                            });
                        }

                        //權限判斷
                        var comp = "";
                        if ($(data).find("person").length > 0) {
                            $(data).find("person").each(function (i) {
                                comp = $(this).children("comp").text().trim();

                            });

                        }
                        switch (comp) {
                            case "01":
                            case "SA":
                                //exbtn exbtnPDF btn_ok btn_notok
                                $("#btn_ok").hide();
                                $("#btn_notok").hide();
                                $("#exbtn").show();
                                $("#exbtnPDF").show();
                                break;
                            case "02":
                                if (chkType=="" && chkStatus=="A") {
                                    $("#btn_ok").show();
                                    $("#btn_notok").show();
                                }
                                $("#exbtn").show();
                                $("#exbtnPDF").show();
                                break;
                        }

                    }
                }
            });
        }

        //推動項目與查核點 Version 1
        function getCheckPoint_Old(xml, tag) {
            if ($(xml).find(tag).length > 0) {
                var divstr = '';
                $(xml).find(tag).each(function (i) {
                    var marginTop = (i == 0) ? "margin10T" : "margin30T";
                    divstr += '<div class="font-bold ' + marginTop + '" style="font-size:14pt;">' + (i + 1) + '、' + $(this).attr("P_ItemName") + '</div>';
                    divstr += '<div class="font-size3 margin5TB">(1)執行進度</div>';
                    divstr += '<div class="stripecomplex font-normal">';
                    var tab1 = '<table width="100%" border="0" cellspacing="0" cellpadding="0">';
                    tab1 += '<thead><tr>';
                    tab1 += '<th nowrap="nowrap" rowspan="2" style="width:150px;">工作比重(%)</th>';
                    tab1 += '<th nowrap="nowrap" rowspan="2" style="width:150px;">年月<br />進度(%)</th>';
                    var year_str = ''; //年
                    var month_str = '<tr>'; //月
                    var cpstr = ''; // 查核點
                    var pstr = ''; // 預定進度
                    var realstr = ''; // 實際進度
                    var cpdesc = ''; // 查核點進度說明
                    var tmpCount = 0; // 年 colspan
                    $(this).children().each(function (i) {
                        //執行進度-Head
                        //年
                        if ($(this).parent().children().length == 1) //當查核點只有一個月份時
                            year_str += '<th  style="text-align:center;">' + $(this).children("CP_ReserveYear").text().trim() + '年</th>';
                        else {
                            if (i == 0)
                                tmpCount += 1;
                            else if ($(this).prev().children("CP_ReserveYear").text().trim() != $(this).children("CP_ReserveYear").text().trim()) { //跨年時
                                year_str += '<th colspan="' + tmpCount + '" style="text-align:center;">' + $(this).prev().children("CP_ReserveYear").text().trim() + '年</th>';
                                tmpCount = 1; //跨年需重置
                            }
                            else if ($(this).parent().children().length == (i + 1)) { //最後一筆資料
                                tmpCount += 1; //最後一筆也要算
                                year_str += '<th colspan="' + tmpCount + '" style="text-align:center;">' + $(this).children("CP_ReserveYear").text().trim() + '年</th></tr>';
                            }
                            else
                                tmpCount += 1;
                        }
                        //月
                        month_str += '<th style="text-align:center;">' + $(this).children("CP_ReserveMonth").text().trim() + '月</th>';

                        //執行進度-Body
                        if (i == 0) {
                            cpstr += '<tr>';
                            cpstr += '<td rowspan="3" style="text-align:center;">' + $(this).parent().attr("P_WorkRatio") + '%</td>';
                            cpstr += '<td style="text-align:center;">查核點</td>';
                            cpstr += '<td style="text-align:center;">' + $(this).children("CP_Point").text().trim() + '</td>';
                            pstr += '<tr>';
                            pstr += '<td style="text-align:center;">累計預定進度(%)</td>';
                            if ($(this).children("CP_Process").text().trim() != "") {
                                pstr += '<td style="text-align:center;">' + $(this).children("CP_Process").text().trim() + '%</td>';
                            } else {
                                pstr += '<td></td>';
                            }
                            realstr += '<tr>';
                            realstr += '<td style="text-align:center;">累計實際進度(%)</td>';
                            if ($(this).children("CP_RealProcess").text().trim() != "") {
                                realstr += '<td style="text-align:center;">' + $(this).children("CP_RealProcess").text().trim() + '</td>';
                            } else {
                                realstr += '<td></td>';
                            }
                        }
                        else if ($(this).parent().children().length == (i + 1)) { //最後一筆資料
                            cpstr += '<td style="text-align:center;">' + $(this).children("CP_Point").text().trim() + '</td></tr>';
                            pstr += '<td style="text-align:center;">' + $(this).children("CP_Process").text().trim() + '%</td></tr>';
                            if ($(this).children("CP_RealProcess").text().trim() != "") {
                                realstr += '<td style="text-align:center;">' + $(this).children("CP_RealProcess").text().trim() + '%</td>';
                            } else {
                                realstr += '<td></td>';
                            }
                        }
                        else {
                            cpstr += '<td style="text-align:center;">' + $(this).children("CP_Point").text().trim() + '</td>';
                            pstr += '<td style="text-align:center;">' + $(this).children("CP_Process").text().trim() + '%</td>';
                            if ($(this).children("CP_RealProcess").text().trim() != "") {
                                realstr += '<td style="text-align:center;">' + $(this).children("CP_RealProcess").text().trim() + '%</td>';
                            } else {
                                realstr += '<td></td>';
                            }
                        }

                        //查核點進度說明-Body
                        cpdesc += '<tr>';
                        cpdesc += '<td>' + $(this).children("CP_Point").text().trim() + '  ' + $(this).children("CP_Desc").text().trim() + '</td>';
                        cpdesc += '<td>' + replaceLine($(this).children("CP_Summary").text().trim()) + '</td>';
                        cpdesc += '<td>' + replaceLine($(this).children("CP_BackwardDesc").text().trim()) + '</td>';
                        cpdesc += '</tr>';
                    });
                    tab1 += year_str + month_str;
                    tab1 += '</tr></thead>';
                    tab1 += '<tbody>' + cpstr + pstr + realstr + '</tbody>';
                    tab1 += '</table>';
                    divstr += tab1 + '</div>';
                    divstr += '<div class="font-size3 margin10T">(2)查核點進度說明</div>';
                    divstr += '<div class="stripecomplex margin5T font-normal">';
                    var tab2 = '<table width="100%" border="0" cellspacing="0" cellpadding="0">';
                    tab2 += '<thead><tr>';
                    tab2 += '<th>查核點</th>';
                    tab2 += '<th style="width:35%;">辦理情形</th>';
                    tab2 += '<th style="width:35%;">進度差異說明</th>';
                    tab2 += '</tr></thead>';
                    tab2 += '<tbody>' + cpdesc + '</tbody>';
                    tab2 += '</table>';
                    divstr += tab2 + '</div>';
                });
                return divstr;
            }
        }

        //推動項目與查核點 Version 2
        function getCheckPoint(xml, pdXml, tag) {
            if ($(xml).find(tag).length > 0) {
                var divstr = '';
                $(xml).find(tag).each(function (i) {
                    var marginTop = (i == 0) ? "margin10T" : "margin30T";
                    divstr += '<div class="font-bold ' + marginTop + '" style="font-size:14pt;">' + (i + 1) + '、' + $(this).attr("P_ItemName") + '</div>';
                    divstr += '<div class="font-size3 margin5TB">(1)執行進度</div>';
                    divstr += '<div class="stripecomplex font-normal">';
                    var tab1 = '<table width="100%" border="0" cellspacing="0" cellpadding="0">';
                    tab1 += '<thead><tr>';
                    tab1 += '<th nowrap="nowrap" rowspan="2" style="width:150px;">工作比重(%)</th>';
                    tab1 += '<th nowrap="nowrap" rowspan="2" style="width:150px;">年月<br />進度(%)</th>';
                    var year_str = ''; //年
                    var month_str = '<tr>'; //月
                    var cpstr = ''; // 查核點
                    var pstr = ''; // 預定進度
                    var realstr = ''; // 實際進度
                    var cpdesc = ''; // 查核點進度說明
                    var tmpCount = 0; // 年 colspan
                    $(this).children().each(function (i) {
                        //執行進度-Head
                        //年
                        if ($(this).parent().children().length == 1) //當查核點只有一個月份時
                            year_str += '<th  style="text-align:center;">' + $(this).children("CP_ReserveYear").text().trim() + '年</th>';
                        else {
                            if (i == 0)
                                tmpCount += 1;
                            else if ($(this).prev().children("CP_ReserveYear").text().trim() != $(this).children("CP_ReserveYear").text().trim()) { //跨年時
                                year_str += '<th colspan="' + tmpCount + '" style="text-align:center;">' + $(this).prev().children("CP_ReserveYear").text().trim() + '年</th>';
                                tmpCount = 1; //跨年需重置
                                if ($(this).parent().children().length == (i + 1)) //若為最後一筆資料需補上
                                  year_str += '<th colspan="' + tmpCount + '" style="text-align:center;">' + $(this).children("CP_ReserveYear").text().trim() + '年</th></tr>';
                            }
                            else if ($(this).parent().children().length == (i + 1)) { //最後一筆資料
                                tmpCount += 1; //最後一筆也要算
                                year_str += '<th colspan="' + tmpCount + '" style="text-align:center;">' + $(this).children("CP_ReserveYear").text().trim() + '年</th></tr>';
                            }
                            else
                                tmpCount += 1;
                        }
                        //月
                        month_str += '<th style="text-align:center;">' + $(this).children("CP_ReserveMonth").text().trim() + '月</th>';

                        //執行進度-Body
                        if (i == 0) {
                            cpstr += '<tr>';
                            cpstr += '<td rowspan="3" style="text-align:center;">' + $(this).parent().attr("P_WorkRatio") + '%</td>';
                            cpstr += '<td style="text-align:center;">查核點</td>';
                            cpstr += '<td style="text-align:center;">' + $(this).children("CP_Point").text().trim() + '</td>';
                            pstr += '<tr>';
                            pstr += '<td style="text-align:center;">累計預定進度(%)</td>';
                            if ($(this).children("CP_Process").text().trim() != "") {
                                pstr += '<td style="text-align:center;">' + $(this).children("CP_Process").text().trim() + '%</td>';
                            } else {
                                pstr += '<td></td>';
                            }
                            realstr += '<tr>';
                            realstr += '<td style="text-align:center;">累計實際進度(%)</td>';
                            if ($(this).children("CP_RealProcess").text().trim() != "") {
                                realstr += '<td style="text-align:center;">' + $(this).children("CP_RealProcess").text().trim() + '</td>';
                            } else {
                                realstr += '<td></td>';
                            }
                        }
                        else if ($(this).parent().children().length == (i + 1)) { //最後一筆資料
                            cpstr += '<td style="text-align:center;">' + $(this).children("CP_Point").text().trim() + '</td></tr>';
                            pstr += '<td style="text-align:center;">' + $(this).children("CP_Process").text().trim() + '%</td></tr>';
                            if ($(this).children("CP_RealProcess").text().trim() != "") {
                                realstr += '<td style="text-align:center;">' + $(this).children("CP_RealProcess").text().trim() + '%</td>';
                            } else {
                                realstr += '<td></td>';
                            }
                        }
                        else {
                            cpstr += '<td style="text-align:center;">' + $(this).children("CP_Point").text().trim() + '</td>';
                            pstr += '<td style="text-align:center;">' + $(this).children("CP_Process").text().trim() + '%</td>';
                            if ($(this).children("CP_RealProcess").text().trim() != "") {
                                realstr += '<td style="text-align:center;">' + $(this).children("CP_RealProcess").text().trim() + '%</td>';
                            } else {
                                realstr += '<td></td>';
                            }
                        }

                        //查核點進度說明-Body
                        //cpdesc += '<tr>';
                        //cpdesc += '<td>' + $(this).children("CP_Point").text().trim() + '  ' + $(this).children("CP_Desc").text().trim() + '</td>';
                        //cpdesc += '<td>' + replaceLine($(this).children("CP_Summary").text().trim()) + '</td>';
                        //cpdesc += '<td>' + replaceLine($(this).children("CP_BackwardDesc").text().trim()) + '</td>';
                        //cpdesc += '</tr>';

                        //查核點項目 (Version 3)
                        cpdesc += $(this).children("CP_Point").text().trim() + '  ' + $(this).children("CP_Desc").text().trim() + '<br>';
                    });
                    tab1 += year_str + month_str;
                    tab1 += '</tr></thead>';
                    tab1 += '<tbody>' + cpstr + pstr + realstr + '</tbody>';
                    tab1 += '</table>';
                    divstr += tab1 + '</div>';
                    divstr += '<div class="font-size3 margin10T">(2)查核點進度說明</div>';
                    divstr += '<div class="stripecomplex margin5T font-normal">';
                    var tab2 = '<table width="100%" border="0" cellspacing="0" cellpadding="0">';
                    tab2 += '<thead><tr>';
                    tab2 += '<th>查核點</th>';
                    tab2 += '<th>年 季</th>';
                    tab2 += '<th style="width:35%;">辦理情形</th>';
                    tab2 += '<th style="width:35%;">進度差異說明</th>';
                    tab2 += '</tr></thead>';
                    /// 查核點進度說明
                    var pdstr = '';
                    if ($(pdXml).find('pd_item[PD_PushitemGuid="' + $(this).attr("P_Guid") + '"]').length > 0) {
                        // Rowspan
                        var rspan_tmp = $(pdXml).find('pd_item[PD_PushitemGuid="' + $(this).attr("P_Guid") + '"]').length;
                        // 處理資料庫中現有資料
                        $(pdXml).find('pd_item[PD_PushitemGuid="' + $(this).attr("P_Guid") + '"]').each(function (i) {
                            pdstr += '<tr>';
                            if (i == 0)
                                pdstr += '<td rowspan="' + rspan_tmp + '">' + cpdesc + '</td>';
                            pdstr += '<td nowrap="nowrap" style="text-align:center;">' + $(this).attr("PD_Year") + '年<br>第' + $(this).attr("PD_Season") + '季</td>';
                            pdstr += '<td>' + $(this).attr("PD_Summary") + '</td>';
                            pdstr += '<td>' + $(this).attr("PD_BackwardDesc") + '</td>';
                            pdstr += '</tr>';
                        });
                    }
                    tab2 += '<tbody>' + pdstr + '</tbody>';
                    tab2 += '</table>';
                    divstr += tab2 + '</div>';
                });
                return divstr;
            }
        }
        
       //擴大補助預計(累計)完成數
        function CreateExFinishTable(data, CPxml) {
            if ($(CPxml).find("PushItem[P_Type='04']").length > 0) {
                var tabstr = '';
                $("#exFinishTab tbody").empty();
                $(CPxml).find("PushItem[P_Type='04']").each(function (i) {
                    if ($(this).attr("P_ItemName") != "其他") {
                        tabstr += '<tr>';
                        tabstr += '<td style="text-align:center;">' + $(this).attr("P_ItemName") + '</td>';
                        tabstr += '<td style="text-align:right;">' + $(this).attr("P_ExFinish") + '</td>';
                        var PGuid = $(this).attr("P_Guid");
                        $(data).find("ex_item").each(function () {
                            if ($(this).children("EF_PushitemId").text().trim() == PGuid) {
                                tabstr += '<td style="text-align:right;">' + $(this).children("EF_Finish").text().trim() + '</td>';
                            }
                        });
                        tabstr += '</tr>';
                    }
                });
                $("#exFinishTab tbody").append(tabstr);
            }
            else {
                $("#exFinishTab tbody").empty();
                $("#exFinishTab tbody").append('<tr><td colspan="3">查詢無資料</td></tr>');
            }
        }

        function getFileList(type, tabName) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/GetFileList.aspx",
                data: {
                    //pGuid: $("#tmpguid").val(),
                    pGuid: $("#hidden_rsguid").val(),
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
                                    //tabstr += '<td align="center" width="50px"><a href="javascript:void(0);" id=' + $(this).children("file_id").text().trim() + ' name="delfilebtn" atp="' + $(this).children("file_type").text().trim() +
                                    //   '"><img src="../App_Themes/images/icon-delete-new.png" /></a></td>';
                                    tabstr += '</tr>';
                                });
                                tabstr += '</tbody>';
                                $(tabName).append(tabstr);
                                $(".mfile").show();
                                $(tabName).find("tr").mouseover(function () { $(this).addClass("spe"); }).mouseout(function () { $(this).removeClass("spe"); });
                                $(tabName + " table:not(td > table) > tbody > tr:not('.spe'):even").addClass("alt");
                            }
                            else
                                $(tabName).append('<tr><td>查詢無資料</td></tr>');
                        }
                    }
                }
            });
        }

        //計算實支率
        function countMoney() {
            var R01 = ($("#RS_Type01Real")[0].innerText == "") ? 0 : parseFloat($("#RS_Type01Real")[0].innerText);
            var R02 = ($("#RS_Type02Real")[0].innerText == "") ? 0 : parseFloat($("#RS_Type02Real")[0].innerText);
            var R03 = ($("#RS_Type03Real")[0].innerText == "") ? 0 : parseFloat($("#RS_Type03Real")[0].innerText);
            var R04 = ($("#RS_Type04Real")[0].innerText == "") ? 0 : parseFloat($("#RS_Type04Real")[0].innerText);
            
            var tmpCount = (R01 / parseFloat($("#rsMoney1")[0].innerText)) * 100;
            $("#rsMoneyRatio1").html(tmpCount.toFixed(0) + "%");
            var tmpCount2 = (R02 / parseFloat($("#rsMoney2")[0].innerText)) * 100;
            $("#rsMoneyRatio2").html(tmpCount2.toFixed(0) + "%");
            var tmpCount3 = (R03 / parseFloat($("#rsMoney3")[0].innerText)) * 100;
            $("#rsMoneyRatio3").html(tmpCount3.toFixed(0) + "%");
            var tmpCount4 = (R04 / parseFloat($("#rsMoney4")[0].innerText)) * 100;
            var tmp_MoneyRatio4 = ($("#rsMoney4")[0].innerText == "" || Number($("#rsMoney4").html()) == 0) ? "0" : tmpCount4.toFixed(0);
            $("#rsMoneyRatio4").html(tmp_MoneyRatio4 + "%");

            //整體實支數
            $("#rsAllRealMoney").html((R01 + R02 + R03 + R04).toFixed(3));

            //整體實支率
            var allMoney = parseFloat($("#rsAllMoney")[0].innerText)
            var allRatio = ((R01 + R02 + R03 + R04) / allMoney) * 100;
            $("#rsAllMoneyRatio").html(allRatio.toFixed(0) + "%");
        }

        //日期轉民國年
        //參數格式: yyyy/MM/dd
        function toROC_Date(str) {
            var day = new Date(str);
            var y = parseInt(day.getFullYear()) - 1911;
            var m = ((day.getMonth() + 1) < 10) ? "0" + (day.getMonth() + 1) : (day.getMonth() + 1);
            var d = (day.getDate() < 10) ? "0" + day.getDate() : day.getDate();
            var rVal = y + "/" + m + "/" + d;
            return rVal;
        }

        //計算合計月份
        //*sd: 起始日期
        //*ed: 結束日期
        //參數格式: yyyy/MM/dd
        function countMonth(sd, ed) {
            var Sday = new Date(sd);
            var Eday = new Date(ed);
            var countYear = Eday.getFullYear() - Sday.getFullYear();
            var countMonth = (((Eday.getMonth() + 1) - (Sday.getMonth() + 1)) + 1);
            var rVal = countYear * 12 + countMonth;
            return rVal;
        }
        //格式化日期
        function formatDate(date) {
            if (date == "") {
                var d = new Date();
            } else {
                var d = new Date(date);
            }
            var month = '' + (d.getMonth() + 1);
            var day = '' + d.getDate();
            var year = d.getFullYear();

            if (month.length < 2) month = '0' + month;
            if (day.length < 2) day = '0' + day;

            return [year, month, day].join('/');
        }

        //送審
        function goCheck(id) {
            var chkType = "";
            if (id == "btn_ok") {
                chkType = "Y";
                checkUpdate(chkType);
            }
            if (id == "btn_notok") {
                chkType = "N";
                $("#submitbtn").show();
                $("#hidden_typeval").val(chkType);
                openMailOP();
            }



        }

        //更新季報審核通過/不通過
        function checkUpdate(strtype) {
            $.ajax({
                type: "POST",
                async: false, //在沒有返回值之前,不會執行下一步動作
                url: "../handler/ReportCheck.ashx",
                data: {
                    func: "reportCheck",
                    str_reportGuid: $("#hidden_reportGuid").val(),
                    str_checkType: strtype,
                    str_ReportType: "02",
                    yyyy: $("#hidden_year").val(),
                    season: $("#hidden_season").val()
                },
                error: function (xhr) {
                    alert("Error " + xhr.status);
                    console.log(xhr.responseText);
                },
                success: function (data) {
                    if (data == "timeout") {
                        alert("請重新登入");
                        location.href = "Login.aspx";
                    } else if (data.indexOf("Error") > -1) {
                        alert(data);
                    } else {
                        if (data == "success") {
                            if (strtype == "Y") {
                                alert("審核操作成功");
                                location.href = "ReviewSeason.aspx";
                            } else {
                                sendmail();
                            }
                            
                        }
                    }
                }
            });//ajax end
        }

        //打開輸入"審核不通過"dialog
        function openMailOP() {
            $.fancybox({
                href: "#opblock",
                title: "",
                //closeBtn: false,
                minWidth: "400",
                minHeight: "200",
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
                        locked: false   //開始fancybox時,背景是否回top
                        //closeClick: false //點背景關閉 fancybox
                    }
                }
            });
        }

        function sendmail() {
            var type = $("input[name='extype']:checked").val();
                $.ajax({
                    type: "POST",
                    async: false, //在沒有返回值之前,不會執行下一步動作
                    url: "../handler/SendMailCheck.aspx",
                    data: {
                        func: "load_projectbyperson",
                        reportGuid: $("#hidden_reportGuid").val(),
                        mailtype: "SnotOK",
                        mailbody: $("#textOp").val(),
                        yyyy: $.getParamValue('year'),
                        season: $.getParamValue('season')
                    },
                    error: function (xhr) {
                        alert("Error " + xhr.status);
                        console.log(xhr.responseText);
                    },
                    success: function (data) {
                        $.fancybox.close();
                        if (data == "timeout") {
                            alert("請重新登入");
                            location.href = "Login.aspx";
                        } else if (data.indexOf("Error") > -1) {
                            alert(data);
                        } else {
                            alert("審核操作成功");
                            location.href = "ReviewSeason.aspx";
                        }
                    }
                });//ajax end
                $.fancybox.close();
        }

        function openExport() {
            $.fancybox({
                href: "#exblock",
                title: "",
                //closeBtn: false,
                minWidth: "200",
                minHeight: "100",
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
                        locked: false   //開始fancybox時,背景是否回top
                        //closeClick: false //點背景關閉 fancybox
                    }
                }
            });
        }

        //將\n換行自原替換成<br />
        function replaceLine(reVal) {
            var strVal = "";
            //一次將該字串底下所有的\n一次Replace掉
            strVal = reVal.replace(/\n/g, "<br />");
            return strVal;
        }


    </script>
</asp:Content>

