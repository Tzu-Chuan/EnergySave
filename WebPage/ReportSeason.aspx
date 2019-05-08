<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="ReportSeason.aspx.cs" Inherits="WebPage_ReportSeason" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
  <script type="text/javascript">
      $(document).ready(function () {
          $("#rYear").html($.getParamValue('year') + " 年");
          ShowSeason($.getParamValue('season'));
          $("#rStage").html("第 " + $.getParamValue('stage') + " 期");

          getInfo();
          $("#Loading").hide();
          $("#ContentDiv").show();

          //附件檔
          getFileList("03", "#bwFileList");
          getFileList("04", "#pFileList");
          getFileList("05", "#sFileList");
          getFileList("07", "#exFileList");

          //取消
          $(document).on("click", "#btnCancel", function () {
              if (confirm("確定取消？")) {
                  location.href = "SeasonList.aspx";
              }
          });

          //限制只能輸入數字
          $(document).on("keyup", ".num", function () {
              if (/[^0-9.]/g.test(this.value)) {
                  this.value = this.value.replace(/[^0-9.]/g, '');
              }
          });

          //實支率 Change
          $(document).on("change", "#RS_Type01Real,#RS_Type02Real,#RS_Type03Real,#RS_Type04Real", function () {
              countMoney();
          });

          //累計實際進度 Change
          $(document).on("change", "input[name='CP_RealProcess']", function () {
              var nowType = $(this).attr("tp");
              var tmpVal = 0;
              $("input[name='CP_RealProcess']").each(function () {
                  if ($(this).attr("tp") == nowType) {
                      tmpVal += Number(this.value);
                  }
              });
              var rProcess = tmpVal + Number($("#tmpReal" + nowType).val());
              $("#RealProcess_" + nowType).html(Number(rProcess.toFixed(2)));

              //整體實際進度
              var allReal = (Number($("#RealProcess_01").html()) + Number($("#RealProcess_02").html()) + Number($("#RealProcess_03").html())) / 3;
              $("#allRealProcess").html(Number(allReal.toFixed(2)));

              //防呆-累計進度不可超過預定進度最大值
              var FinalProcess = $(this).closest('tr').prev().children("td:last-child").text().replace("%", "");
              if (parseFloat(this.value) > parseFloat(FinalProcess)) {
                  alert("最高不可超過 " + FinalProcess + " %");
                  this.value = FinalProcess;
              }
          });


          //儲存&送審
          $(document).on("click", "#btnSave,#btnSubReview", function () {
              //辦理情形&進度差異說明 convert to XML post to backend
              var xmldoc = document.createElement("root");
              $("textarea[name='PD_Summary']").each(function (i) {
                  var xNode = document.createElement("pdItem");
                  var Node = document.createElement("summary");
                  Node.textContent = this.value;
                  var Node2 = document.createElement("backward");
                  Node2.textContent = $("textarea[name='PD_BackwardDesc']")[i].value;
                  xNode.appendChild(Node);
                  xNode.appendChild(Node2);
                  xmldoc.appendChild(xNode);
              });

              var iframe = $('<iframe name="postiframe" id="postiframe" style="display: none" />');
              var year = $('<input type="hidden" name="year" id="year" value="' + $.getParamValue('year') + '" />');
              var season = $('<input type="hidden" name="season" id="season" value="' + $.getParamValue('season') + '" />');
              var stage = $('<input type="hidden" name="stage" id="stage" value="' + $.getParamValue('stage') + '" />');
              var tmpXML = $('<input type="hidden" name="tmpXML" id="tmpXML" value="' + encodeURIComponent(xmldoc.outerHTML) + '" />');
              var form = $("form")[0];

              //如果沒有重新導頁需要刪除上次資訊
              $("#postiframe").remove();
              $("input[name='year']").remove();
              $("input[name='season']").remove();
              $("input[name='stage']").remove();
              $("input[name='tmpXML']").remove();

              form.appendChild(iframe[0]);
              form.appendChild(year[0]);
              form.appendChild(season[0]);
              form.appendChild(stage[0]);
              form.appendChild(tmpXML[0]);

              if (this.id == "btnSubReview") {
                  var subbtn = $('<input type="hidden" name="subbtn" id="subbtn" value="Y" />');
                  $("input[name='subbtn']").remove();
                  form.appendChild(subbtn[0]);
              }

              form.setAttribute("action", "../handler/addSeason.ashx");
              form.setAttribute("method", "post");
              form.setAttribute("enctype", "multipart/form-data");
              form.setAttribute("encoding", "multipart/form-data");
              form.setAttribute("target", "postiframe");
              form.submit();
          });

          //上傳 button
          $(document).on("click", "#btn_up03,#btn_up04,#btn_up05,#btn_up07", function () {
              var type = "";
              switch (this.id) {
                  case "btn_up03":
                      type = "03";
                      break;
                  case "btn_up04":
                      type = "04";
                      break;
                  case "btn_up05":
                      type = "05";
                      break;
                  case "btn_up07":
                      type = "07";
                      break;
              }

              $.fancybox({
                  href: "File_Upload.aspx?v=" + $("#tmpguid").val() + "&tp=" + type,
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
                                  switch (str_atp) {
                                      case "03":
                                          getFileList(str_atp, "#bwFileList");
                                          break;
                                      case "04":
                                          getFileList(str_atp, "#pFileList");
                                          break;
                                      case "05":
                                          getFileList(str_atp, "#sFileList");
                                          break;
                                      case "07":
                                          getFileList(str_atp, "#exFileList");
                                          break;
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
      }); // js end

      function getInfo() {
          $.ajax({
              type: "POST",
              async: false, //在沒有返回值之前,不會執行下一步動作
              url: "../handler/getSeasonDetail.aspx",
              data: {
                  Year: $.getParamValue('year'),
                  Season: $.getParamValue('season'),
                  Stage: $.getParamValue('stage')
              },
              error: function (xhr) {
                  alert("Error " + xhr.status);
                  console.log(xhr.responseText);
              },
              success: function (data) {
                  if ($(data).find("Error").length > 0) {
                      alert($(data).find("Error").attr("Message"));
                  }
                  else {
                      //季報資料
                      if ($(data).find("season_item").length > 0) {
                          $(data).find("season_item").each(function (i) {
                              $("#RS_CostDesc").val($(this).children("RS_CostDesc").text().trim());
                              $("#RS_Type01Real").val($(this).children("RS_Type01Real").text().trim());
                              $("#RS_Type02Real").val($(this).children("RS_Type02Real").text().trim());
                              $("#RS_Type03Real").val($(this).children("RS_Type03Real").text().trim());
                              $("#RS_Type04Real").val($(this).children("RS_Type04Real").text().trim());
                              $("#RS_03Type01C").val($(this).children("RS_03Type01C").text().trim());
                              $("#RS_03Type02C").val($(this).children("RS_03Type02C").text().trim());
                              $("#RS_03Type03C").val($(this).children("RS_03Type03C").text().trim());
                              $("#RS_03Type04C").val($(this).children("RS_03Type04C").text().trim());
                              $("#RS_03Type05C").val($(this).children("RS_03Type05C").text().trim());
                          });
                      }

                      //基本資料
                      if ($(data).find("data_item").length > 0) {
                          $(data).find("data_item").each(function (i) {
                              $("#rsCity").html($(this).children("CityName").text().trim());
                              $("#rsOffice").html($(this).children("I_Office").text().trim());
                              var StartDate = $(this).children("I_" + $.getParamValue('stage') + "_Sdate").text().trim();
                              var EndDate = $(this).children("I_" + $.getParamValue('stage') + "_Edate").text().trim();
                              $("#rsStartDate").html(toROC_Date(StartDate));
                              $("#rsEndDate").html(toROC_Date(EndDate));
                              $("#rsTotalMonth").html(countMonth(StartDate, EndDate));
                              $("#rsMoney1").html($(this).children("I_Money_item1_" + $.getParamValue('stage')).text().trim());
                              $("#rsMoney2").html($(this).children("I_Money_item2_" + $.getParamValue('stage')).text().trim());
                              $("#rsMoney3").html($(this).children("I_Money_item3_" + $.getParamValue('stage')).text().trim());
                              $("#rsMoney4").html($(this).children("I_Money_item4_" + $.getParamValue('stage')).text().trim());
                              var tmpRS4 = ($("#rsMoney4").html() == "") ? 0 : parseFloat($("#rsMoney4").html());
                              var allMoney = (parseFloat($("#rsMoney1").html()) + parseFloat($("#rsMoney2").html()) + parseFloat($("#rsMoney3").html()) + tmpRS4).toFixed(3);
                              $("#rsAllMoney").html(allMoney);
                              $("#DR_Finish1").html($(this).children("I_Finish_item1_" + $.getParamValue('stage')).text().trim());
                              $("#DR_Finish2").html($(this).children("I_Finish_item2_" + $.getParamValue('stage')).text().trim());
                              $("#DR_Finish3").html($(this).children("I_Finish_item3_" + $.getParamValue('stage')).text().trim());
                              $("#DR_Finish4").html($(this).children("I_Finish_item4_" + $.getParamValue('stage')).text().trim());
                              $("#DR_Finish5").html($(this).children("I_Finish_item5_" + $.getParamValue('stage')).text().trim());
                          });
                      }

                      //預定進度&實際進度
                      if ($(data).find("process_item").length > 0) {
                          $(data).find("process_item").each(function (i) {
                              $("#allProcess").html(Number(Number($(this).children("sumValAll").text().trim()).toFixed(2)));
                              $("#allRealProcess").html(Number(Number($(this).children("sumRealValAll").text().trim()).toFixed(2)));
                              $("#Process_01").html(Number(Number($(this).children("sumVal01").text().trim()).toFixed(2)));
                              $("#RealProcess_01").html(Number(Number($(this).children("sumRealVal01").text().trim()).toFixed(2)));
                              $("#Process_02").html(Number(Number($(this).children("sumVal02").text().trim()).toFixed(2)));
                              $("#RealProcess_02").html(Number(Number($(this).children("sumRealVal02").text().trim()).toFixed(2)));
                              $("#Process_03").html(Number(Number($(this).children("sumVal03").text().trim()).toFixed(2)));
                              $("#RealProcess_03").html(Number(Number($(this).children("sumRealVal03").text().trim()).toFixed(2)));
                              $("#Process_04").html(Number(Number($(this).children("sumVal04").text().trim()).toFixed(2)));
                              $("#RealProcess_04").html(Number(Number($(this).children("sumRealVal04").text().trim()).toFixed(2)));
                          });
                      }

                      //累計實際進度扣除本季實際進度資料
                      if ($(data).find("process2_item").length > 0) {
                          $(data).find("process2_item").each(function (i) {
                              $("#tmpReal01").val($(this).children("sumRealVal01").text().trim());
                              $("#tmpReal02").val($(this).children("sumRealVal02").text().trim());
                              $("#tmpReal03").val($(this).children("sumRealVal03").text().trim());
                              $("#tmpReal04").val($(this).children("sumRealVal04").text().trim());
                          });
                      }

                      //三大項目動態資訊
                      $("#div_type01").append(getCheckPoint(data, "PushItem[P_Type='01']"));
                      $("#div_type02").append(getCheckPoint(data, "PushItem[P_Type='02']"));
                      $("#div_type03").append(getCheckPoint(data, "PushItem[P_Type='03']"));
                      $("#div_type04").append(getCheckPoint(data, "PushItem[P_Type='04']"));

                      //擴大補助預計完成數
                      CreateExFinishTable(data);

                      //頁尾人員基本資料
                      if ($(data).find("per_item").length > 0) {
                          $(data).find("per_item").each(function (i) {
                              $("#UserName").html($(this).children("UserName").text().trim());
                              $("#UserTel").html($(this).children("UserTel").text().trim());
                              $("#WriteDate").html(toROC_Date(new Date()));
                              $("#ManagerName").html($(this).children("ManagerName").text().trim());
                          });
                      }

                      countMoney();
                      $("#BtnPanel").show();
                  }
              }
          });
      }

      //推動項目與查核點 Html
      function getCheckPoint(xml, tag) {
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
                          pstr += '<td style="text-align:center;">累計預定進度</td>';
                          pstr += '<td style="text-align:center;">' + $(this).children("CP_Process").text().trim() + '%</td>';
                          realstr += '<tr>';
                          realstr += '<td style="text-align:center;">累計實際進度(%)</td>';
                          if ($(this).children("CP_ReserveMonth").text().trim() == $("#tmpMonth").val() && $(this).children("CP_ReserveYear").text().trim() == $.getParamValue('year')) {
                              realstr += '<td style="text-align:center;">';
                              realstr += '<input type="text" class="inputex width60 num" tp="' + $(this).parent().attr("P_Type") + '" name="CP_RealProcess" value="' + $(this).children("CP_RealProcess").text().trim() + '" /> ';
                              realstr += '<input type="hidden" name="cpGuid" value="' + $(this).children("CP_Guid").text().trim() + '" /></td>';
                          }
                          else
                              realstr += '<td style="text-align:center;"><input type="text" class="inputex width60" value="' + $(this).children("CP_RealProcess").text().trim() + '" disabled="disabled" /></td>';
                      }
                      else if ($(this).parent().children().length == (i + 1)) { //最後一筆資料
                          cpstr += '<td style="text-align:center;">' + $(this).children("CP_Point").text().trim() + '</td></tr>';
                          pstr += '<td style="text-align:center;">' + $(this).children("CP_Process").text().trim() + '%</td></tr>';
                          if ($(this).children("CP_ReserveMonth").text().trim() == $("#tmpMonth").val() && $(this).children("CP_ReserveYear").text().trim() == $.getParamValue('year')) {
                              realstr += '<td style="text-align:center;">';
                              realstr += '<input type="text" class="inputex width60 num" tp="' + $(this).parent().attr("P_Type") + '" name="CP_RealProcess" value="' + $(this).children("CP_RealProcess").text().trim() + '" /> ';
                              realstr += '<input type="hidden" name="cpGuid" value="' + $(this).children("CP_Guid").text().trim() + '" /></td>';
                          }
                          else
                              realstr += '<td style="text-align:center;"><input type="text" class="inputex width60" value="' + $(this).children("CP_RealProcess").text().trim() + '" disabled="disabled" /></td>';
                      }
                      else {
                          cpstr += '<td style="text-align:center;">' + $(this).children("CP_Point").text().trim() + '</td>';
                          pstr += '<td style="text-align:center;">' + $(this).children("CP_Process").text().trim() + '%</td>';
                          if ($(this).children("CP_ReserveMonth").text().trim() == $("#tmpMonth").val() && $(this).children("CP_ReserveYear").text().trim() == $.getParamValue('year')) {
                              realstr += '<td style="text-align:center;">';
                              realstr += '<input type="text" class="inputex width60 num" tp="' + $(this).parent().attr("P_Type") + '" name="CP_RealProcess" value="' + $(this).children("CP_RealProcess").text().trim() + '" /> ';
                              realstr += '<input type="hidden" name="cpGuid" value="' + $(this).children("CP_Guid").text().trim() + '" /></td>';
                          }
                          else
                              realstr += '<td style="text-align:center;"><input type="text" class="inputex width60" value="' + $(this).children("CP_RealProcess").text().trim() + '" disabled="disabled" /></td>';
                      }

                      //查核點進度說明-Body (Version 2)
                      //cpdesc += '<tr>';
                      //cpdesc += '<td>' + $(this).children("CP_Point").text().trim() + '  ' + $(this).children("CP_Desc").text().trim() + '</td>';

                      //if ($(this).children("CP_ReserveMonth").text().trim() == $("#tmpMonth").val() && $(this).children("CP_ReserveYear").text().trim() == $.getParamValue('year')) {
                      //    cpdesc += '<td><textarea rows="4" class="inputex width100" name="CP_Summary">'+$(this).children("CP_Summary").text().trim()+'</textarea></td>';
                      //    cpdesc += '<td><textarea rows="4" class="inputex width100" name="CP_BackwardDesc"">'+$(this).children("CP_BackwardDesc").text().trim()+'</textarea></td>';
                      //}
                      //else {
                      //    cpdesc += '<td><textarea rows="4" class="inputex width100" disabled="disabled">'+$(this).children("CP_Summary").text().trim()+'</textarea></td>';
                      //    cpdesc += '<td><textarea rows="4" class="inputex width100" disabled="disabled">'+$(this).children("CP_BackwardDesc").text().trim()+'</textarea></td>';
                      //}
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
                  tab2 += '<th style="width:30%;">辦理情形</th>';
                  tab2 += '<th style="width:30%;">進度差異說明</th>';
                  tab2 += '</tr></thead>';
                  /// 查核點進度說明
                  var pdstr = '';
                  var ParentID = $(this).attr("P_Guid");
                  if ($(xml).find('pd_item[PD_PushitemGuid="' + $(this).attr("P_Guid") + '"]').length > 0) {
                      // Rowspan
                      var rspan_tmp = $(xml).find('pd_item[PD_PushitemGuid="' + $(this).attr("P_Guid") + '"]').length;
                      // 若無當季資料 Rowspan + 1
                      if ($(xml).find('pd_item[PD_PushitemGuid="' + $(this).attr("P_Guid") + '"][PD_Year="' + $.getParamValue('year') + '"][PD_Season="' + $.getParamValue('season') + '"][PD_Stage="' + $.getParamValue('stage') + '"]').length == 0)
                          rspan_tmp += 1;
                      // 處理資料庫中現有資料
                      $(xml).find('pd_item[PD_PushitemGuid="' + $(this).attr("P_Guid") + '"]').each(function (i) {
                          pdstr += '<tr>';
                          if (i == 0)
                              pdstr += '<td rowspan="' + rspan_tmp + '">' + cpdesc + '</td>';
                          pdstr += '<td nowrap="nowrap" style="text-align:center;">' + $(this).attr("PD_Year") + '年<br>第' + $(this).attr("PD_Season") + '季</td>';
                          if ($.getParamValue('year') == $(this).attr("PD_Year") && $.getParamValue('season') == $(this).attr("PD_Season") && $.getParamValue('stage') == $(this).attr("PD_Stage")) {
                              pdstr += '<td><textarea rows="4" class="inputex width100" name="PD_Summary">' + $(this).attr("PD_Summary") + '</textarea><input type="hidden" name="pi_guid" value="' + ParentID + '" /></td>';
                              pdstr += '<td><textarea rows="4" class="inputex width100" name="PD_BackwardDesc">' + $(this).attr("PD_BackwardDesc") + '</textarea></td>';
                          }
                          else {
                              pdstr += '<td><textarea rows="4" class="inputex width100" disabled="disabled">' + $(this).attr("PD_Summary") + '</textarea></td>';
                              pdstr += '<td><textarea rows="4" class="inputex width100" disabled="disabled">' + $(this).attr("PD_BackwardDesc") + '</textarea></td>';
                          }
                          pdstr += '</tr>';
                      });
                      // 若無當季資料需再加上
                      if ($(xml).find('pd_item[PD_PushitemGuid="' + $(this).attr("P_Guid") + '"][PD_Year="' + $.getParamValue('year') + '"][PD_Season="' + $.getParamValue('season') + '"][PD_Stage="' + $.getParamValue('stage') + '"]').length == 0) {
                          pdstr += '<tr>';
                          pdstr += '<td nowrap="nowrap" style="text-align:center;">' + $.getParamValue('year') + '年<br>第' + $.getParamValue('season') + '季</td>';
                          pdstr += '<td><textarea rows="4" class="inputex width100" name="PD_Summary"></textarea><input type="hidden" name="pi_guid" value="' + ParentID + '" /></td>';
                          pdstr += '<td><textarea rows="4" class="inputex width100" name="PD_BackwardDesc"></textarea></td>';
                          pdstr += '</tr>';
                      }
                  }
                  else {
                      pdstr += '<tr>';
                      pdstr += '<td>' + cpdesc + '</td>';
                      pdstr += '<td nowrap="nowrap" style="text-align:center;">' + $.getParamValue('year') + '年<br>第' + $.getParamValue('season') + '季</td>';
                      pdstr += '<td><textarea rows="4" class="inputex width100" name="PD_Summary"></textarea><input type="hidden" name="pi_guid" value="' + ParentID + '" /></td>';
                      pdstr += '<td><textarea rows="4" class="inputex width100" name="PD_BackwardDesc"></textarea></td>';
                      pdstr += '</tr>';
                  }
                  tab2 += '<tbody>' + pdstr + '</tbody>';
                  tab2 += '</table>';
                  divstr += tab2 + '</div>';
              });
              return divstr;
          }
          else
              return '<span style="color:red; font-size:12pt;">' + $.getParamValue('year') + '年 第' + $.getParamValue('stage') + '期 查核點無資料</span>';
      }

       //擴大補助預計(累計)完成數
      function CreateExFinishTable(xml) {
          if ($(xml).find("PushItem[P_Type='04']").length > 0) {
              var tabstr = '';
              $("#exFinishTab tbody").empty();
              $(xml).find("PushItem[P_Type='04']").each(function (i) {
                  tabstr += '<tr>';
                  tabstr += '<td style="text-align:center;">' + $(this).attr("P_ItemName") + '</td>';
                  tabstr += '<td style="text-align:right;">' + $(this).attr("P_ExFinish") + '</td>';
                  tabstr += '<td style="text-align:right;"><input type="text" class="inputex width90" name="exRealFinish" value="' + $(this).attr("EF_Finish") + '" style="text-align:right;" />';
                  tabstr += '<input type="hidden" name="exGuid" value="' + $(this).attr("P_Guid") + '" /></td>';
                  tabstr += '</tr>';
              });
              $("#exFinishTab").append(tabstr);
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
                    pGuid: $("#tmpguid").val(),
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
                                $(".mfile").show();
                                $(tabName).find("tr").mouseover(function () { $(this).addClass("spe"); }).mouseout(function () { $(this).removeClass("spe"); });
                                $(tabName + " table:not(td > table) > tbody > tr:not('.spe'):even").addClass("alt");
                            }
                            else
                                $(tabName).append('<tr><td colspan="2">查詢無資料</td></tr>');
                        }
                    }
                }
            });
        }

      function upFile_feedback(type) {
          switch (type) {
              case "03":
                  getFileList(type, "#bwFileList");
                  break;
              case "04":
                  getFileList(type, "#pFileList");
                  break;
              case "05":
                  getFileList(type, "#sFileList");
                  break;
              case "07":
                  getFileList(type, "#exFileList");
                  break;
          }
      }

      //計算實支率
      function countMoney() {
          var R01 = ($("#RS_Type01Real").val() == "") ? 0 : parseFloat($("#RS_Type01Real").val());
          var R02 = ($("#RS_Type02Real").val() == "") ? 0 : parseFloat($("#RS_Type02Real").val());
          var R03 = ($("#RS_Type03Real").val() == "") ? 0 : parseFloat($("#RS_Type03Real").val());
          var R04 = ($("#RS_Type04Real").val() == "") ? 0 : parseFloat($("#RS_Type04Real").val());

          var tmpCount = (R01 / parseFloat($("#rsMoney1").html())) * 100;
          var tmp_MoneyRatio1 = ($("#rsMoney1").html() == "" || Number($("#rsMoney1").html()) == 0) ? "0" : tmpCount.toFixed(0);
          $("#rsMoneyRatio1").html(tmp_MoneyRatio1 + "%");
          var tmpCount2 = (R02 / parseFloat($("#rsMoney2").html())) * 100;
          var tmp_MoneyRatio2 = ($("#rsMoney2").html() == "" || Number($("#rsMoney2").html()) == 0) ? "0" : tmpCount2.toFixed(0);
          $("#rsMoneyRatio2").html(tmp_MoneyRatio2 + "%");
          var tmpCount3 = (R03 / parseFloat($("#rsMoney3").html())) * 100;
          var tmp_MoneyRatio3 = ($("#rsMoney3").html() == "" || Number($("#rsMoney3").html()) == 0) ? "0" : tmpCount3.toFixed(0);
          $("#rsMoneyRatio3").html(tmp_MoneyRatio3 + "%");
          var tmpCount4 = (R04 / parseFloat($("#rsMoney4").html())) * 100;
          var tmp_MoneyRatio4 = ($("#rsMoney4").html() == "" || Number($("#rsMoney4").html()) == 0) ? "0" : tmpCount4.toFixed(0);
          $("#rsMoneyRatio4").html(tmp_MoneyRatio4 + "%");

          //整體實支數
          $("#rsAllRealMoney").html((R01 + R02 + R03 + R04).toFixed(3));

          //整體實支率
          var allMoney = parseFloat($("#rsAllMoney").html())
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

      //季別 文字處理
      function ShowSeason(str) {
          switch (str) {
              case "1":
                  $("#rSeason").html("第一季");
                  $("#tmpMonth").val("3");
                  break;
              case "2":
                  $("#rSeason").html("第二季");
                  $("#tmpMonth").val("6");
                  break;
              case "3":
                  $("#rSeason").html("第三季");
                  $("#tmpMonth").val("9");
                  break;
              case "4":
                  $("#rSeason").html("第四季");
                  $("#tmpMonth").val("12");
                  break;
          }
      }

      function feedback(xStr) {
          if ($(xStr).find("Error").length > 0) {
              alert($(xStr).find("Error").attr("Message"));
          }
          else {
              if ($(xStr).find("Response").text() == "SubReview") {
                  alert("季報已送審");
                  location.href = "../WebPage/SeasonList.aspx";
              }
              else {
                  alert("儲存完成");
              }
          }
      }
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <input type="hidden" id="tmpguid" name="tmpguid" value="<%= sGuid %>" />
    <input type="hidden" id="autoStatus" value="<%= AutoSaveStatus %>" />
    <input type="hidden" id="tmpMonth" />
    <input type="hidden" id="tmpReal01" />
    <input type="hidden" id="tmpReal02" />
    <input type="hidden" id="tmpReal03" />
    <input type="hidden" id="tmpReal04" />
    <div class="container">
        <div class="twocol filetitlewrapper">
            <div class="left"><span class="filetitle font-size5">季報</span></div>
            <div class="right">進度報表 / 季報</div>
        </div>
        <!-- twocol -->
        <div id="Loading"><img src="../App_Themes/images/loading.gif" width="50" />資料讀取中...</div>
        <div id="ContentDiv" class="OchiTrasTable width100 TitleLength08" style="display:none;">
            <!-- 提報年分  提報季別-->
            <div class="OchiRow">
                <div class="OchiThird">
                    <div class="OchiCell OchiTitle TitleSetWidth">提報年份</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <span id="rYear"></span>
                        </div><!-- cell內容end -->
                    </div><!-- OchiCell -->
                </div>
                <div class="OchiThird">
                    <div class="OchiCell OchiTitle TitleSetWidth">提報季別</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <span id="rSeason"></span>
                        </div><!-- cell內容end -->
                    </div><!-- OchiCell -->
                </div>
                <div class="OchiThird">
                    <div class="OchiCell OchiTitle TitleSetWidth">期程</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <span id="rStage"></span>
                        </div><!-- cell內容end -->
                    </div><!-- OchiCell -->
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
            </div><!-- OchiRow -->

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
                        </div><!-- OchiTableInner -->
                        <!-- cell內容end -->
                    </div><!-- OchiCell -->
                </div>
                <div class="OchiHalf">
                    <div class="OchiCell OchiTitle TitleSetWidth">合計：</div>
                    <div class="OchiCell width100">
                        <!-- cell內容start -->
                        <div class="OchiTableInner width100">
                            <div class="OchiCellInner width100"><span id="rsTotalMonth"></span> 月</div>
                        </div>
                        <!-- OchiTableInner -->
                        <!-- cell內容end -->
                    </div>
                </div>
            </div><!-- OchiRow -->

        <div class="twocol margin5T margin5B" id="div_nodata"></div>
            <!--季報資料-->
            <!-- 單欄 預算狀態-->
            <div class="OchiRow">
                <!--<div class="font-size3 margin10T" style="font-size: 16px;">預算狀態</div>-->
                <div class="font-bold margin5B" style="font-size: 16pt;">預算狀態</div>
                <div class="margin5B">
                    <textarea class="inputex" id="RS_CostDesc" name="RS_CostDesc" rows="4" style="width:970px; max-width:970px;"></textarea><br />
                    備註:如議會已納入預算
                </div>
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
                                <th nowrap="nowrap" style="width:300px;">推動工作</th>
                                <th nowrap="nowrap" style="width:200px;">經費(千元) A</th>
                                <th nowrap="nowrap" style="width:200px;">實支數(千元) B</th>
                                <th nowrap="nowrap">實支率(%) C = B / A</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td style="text-align:center;">節電基礎工作</td>
                                <td style="text-align:right;"><span id="rsMoney1"></span></td>
                                <td style="text-align:right;"><input type="text" class="inputex width90 num" id="RS_Type01Real" name="RS_Type01Real" style="text-align:right;" /></td>
                                <td style="text-align:right;"><span id="rsMoneyRatio1"></span></td>
                            </tr>
                            <tr>
                                <td style="text-align:center;">因地制宜</td>
                                <td style="text-align:right;"><span id="rsMoney2"></span></td>
                                <td style="text-align:right;"><input type="text" class="inputex width90 num" id="RS_Type02Real" name="RS_Type02Real"  style="text-align:right;" /></td>
                                <td style="text-align:right;"><span id="rsMoneyRatio2"></span></td>
                            </tr>
                            <tr>
                                <td style="text-align:center;">設備汰換及智慧用電</td>
                                <td style="text-align:right;"><span id="rsMoney3"></span></td>
                                <td style="text-align:right;"><input type="text" class="inputex width90 num" id="RS_Type03Real" name="RS_Type03Real" style="text-align:right;" /></td>
                                <td style="text-align:right;"><span id="rsMoneyRatio3"></span></td>
                            </tr>
                            <tr>
                                <td style="text-align:center;">擴大補助</td>
                                <td style="text-align:right;"><span id="rsMoney4"></span></td>
                                <td style="text-align:right;"><input type="text" class="inputex width90 num" id="RS_Type04Real" name="RS_Type04Real" style="text-align:right;" /></td>
                                <td style="text-align:right;"><span id="rsMoneyRatio4"></span></td>
                            </tr>
                            <tr>
                                <td style="text-align:center;">整體工作</td>
                                <td style="text-align:right;"><span id="rsAllMoney"></span></td>
                                <td style="text-align:right;"><span id="rsAllRealMoney"></span></td>
                                <td style="text-align:right;"><span id="rsAllMoneyRatio"></span></td>
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
                <div class="OchiTableInner width100 margin10TB" style="margin-left:50px;">
                    <div class="OchiCellInner nowrap">預定執行進度：</div>
                    <div class="OchiCellInner width50"><span id="allProcess"></span>%</div>
                    <div class="OchiCellInner nowrap">實際執行進度：</div>
                    <div class="OchiCellInner width50"><span id="allRealProcess"></span>%</div>
                </div><!-- OchiTableInner -->
                <!-- cell內容end -->
            </div>

            <!--節電基礎工作-->
            <div class="OchiRow">
                <div class="margin10T" style="font-size:16pt; font-weight: bolder;">
                    壹、節電基礎工作&nbsp;<span style="color:blue;">(本季預定：<span id="Process_01"></span>%；本季實際 <span id="RealProcess_01"></span>%)</span>
                </div>
                <!--查核點動態資訊-->
                <div id="div_type01"></div>
                <!--檔案上傳-->
                <div class="twocol margin10TB">
                    <div class="right"><a href="javascript:void(0);" class="genbtnS" id="btn_up03">參考資料上傳</a></div>
                </div>
                <div class="mfile" style="margin-top:10px;display:none;">附件檔</div>
                <div class="stripeMe mfile" style="margin-bottom:10px;display:none;">
                    <table id="bwFileList" width="100%" border="0" cellspacing="0" cellpadding="0">
                        <thead>
                            <tr>
                                <th>檔案名稱</th>
                                <th>刪除</th>
                            </tr>
                        </thead>
                        <tr><td colspan="2">查詢無資料</td></tr>
                    </table>
                </div>
            </div>

            <!--因地制宜-->
            <div class="OchiRow">
                <div class="margin10T" style="font-size:16pt; font-weight: bolder;">
                    貳、因地制宜&nbsp;<span style="color:blue;">(本季預定：<span id="Process_02"></span>%；本季實際 <span id="RealProcess_02"></span>%)</span>
                </div>
                <!--查核點動態資訊-->
                <div id="div_type02"></div>
                <!--檔案上傳-->
                <div class="twocol margin10TB">
                    <div class="right"><a href="javascript:void(0);" class="genbtnS" id="btn_up04">參考資料上傳</a></div>
                </div>
                <div class="mfile" style="margin-top:10px;display:none;">附件檔</div>
                <div class="stripeMe mfile" style="margin-bottom:10px;display:none;">
                    <table id="pFileList" width="100%" border="0" cellspacing="0" cellpadding="0">
                        <thead>
                            <tr>
                                <th>檔案名稱</th>
                                <th>刪除</th>
                            </tr>
                        </thead>
                        <tr><td colspan="2">查詢無資料</td></tr>
                    </table>
                </div>
            </div>

            <!--設備汰換與智慧用電-->
            <div class="OchiRow">
                <div class="margin10T" style="font-size:16pt; font-weight: bolder;">
                    參、設備汰換與智慧用電&nbsp;<span style="color:blue;">(本季預定：<span id="Process_03"></span>%；本季實際 <span id="RealProcess_03"></span>%)</span>
                </div>
                <!--查核點動態資訊-->
                <div id="div_type03"></div>
                <!--檔案上傳-->
                <div class="twocol margin10TB">
                    <div class="right"><a href="javascript:void(0);" class="genbtnS" id="btn_up05">參考資料上傳</a></div>
                </div>
                <div class="mfile" style="margin-top:10px;display:none;">附件檔</div>
                <div class="stripeMe mfile" style="margin-bottom:10px;display:none;">
                    <table id="sFileList" width="100%" border="0" cellspacing="0" cellpadding="0">
                        <thead>
                            <tr>
                                <th>檔案名稱</th>
                                <th>刪除</th>
                            </tr>
                        </thead>
                        <tr><td colspan="2">查詢無資料</td></tr>
                    </table>
                </div>
                
                <div class="stripeMe margin10B font-normal">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <thead>
                            <tr>
                                <th nowrap="nowrap">項目</th>
                                <th nowrap="nowrap" style="width:200px;">本期預計完成數</th>
                                <th nowrap="nowrap" style="width:200px;">本期累計完成數</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td style="text-align:center;">無風管冷氣(kW)註：每台冷氣約4kW</td>
                                <td style="text-align:right;"><span id="DR_Finish1"></span></td>
                                <td style="text-align:right;"><input type="text" class="inputex width90" id="RS_03Type01C" name="RS_03Type01C" style="text-align:right;" /></td>
                            </tr>
                            <tr>
                                <td style="text-align:center;">老舊辦公室照明(具)</td>
                                <td style="text-align:right;"><span id="DR_Finish2"></span></td>
                                <td style="text-align:right;"><input type="text" class="inputex width90" id="RS_03Type02C" name="RS_03Type02C" style="text-align:right;" /></td>
                            </tr>
                            <tr>
                                <td style="text-align:center;">室內停車場智慧照明(盞)</td>
                                <td style="text-align:right;"><span id="DR_Finish3"></span></td>
                                <td style="text-align:right;"><input type="text" class="inputex width90" id="RS_03Type03C" name="RS_03Type03C" style="text-align:right;" /></td>
                            </tr>
                            <tr>
                                <td style="text-align:center;">中型能管系統(套)</td>
                                <td style="text-align:right;"><span id="DR_Finish4"></span></td>
                                <td style="text-align:right;"><input type="text" class="inputex width90" id="RS_03Type04C" name="RS_03Type04C" style="text-align:right;" /></td>
                            </tr>
                            <tr>
                                <td style="text-align:center;">大型能管系統(套)</td>
                                <td style="text-align:right;"><span id="DR_Finish5"></span></td>
                                <td style="text-align:right;"><input type="text" class="inputex width90" id="RS_03Type05C" name="RS_03Type05C" style="text-align:right;" /></td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>

            <!--擴大補助-->
            <div class="OchiRow">
                <div class="margin10T" style="font-size:16pt; font-weight: bolder;">
                    肆、擴大補助&nbsp;<span style="color:blue;">(本季預定：<span id="Process_04"></span>%；本季實際 <span id="RealProcess_04"></span>%)</span>
                </div>
                <!--查核點動態資訊-->
                <div id="div_type04"></div>
                <!--檔案上傳-->
                <div class="twocol margin10TB">
                    <div class="right"><a href="javascript:void(0);" class="genbtnS" id="btn_up07">參考資料上傳</a></div>
                </div>
                <div class="mfile" style="margin-top:10px;display:none;">附件檔</div>
                <div class="stripeMe mfile" style="margin-bottom:10px;display:none;">
                    <table id="exFileList" width="100%" border="0" cellspacing="0" cellpadding="0">
                        <thead>
                            <tr>
                                <th>檔案名稱</th>
                                <th>刪除</th>
                            </tr>
                        </thead>
                        <tr><td colspan="2">查詢無資料</td></tr>
                    </table>
                </div>

                <div class="stripeMe margin10B font-normal">
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

        <div id="BtnPanel" class="twocol margin15T margin5B" style="display:none;">
            <div class="right">
                <a href="javascript:void(0);" class="genbtn" id="btnSubReview" style="color:red;">送審</a>
                <a href="javascript:void(0);" class="genbtn" id="btnSave">儲存</a>
                <a href="javascript:void(0);" class="genbtn" id="btnCancel">取消</a>
            </div>
        </div>
        <!-- twocol -->
    </div>
    
    
    <script type="text/javascript">
        // 自動存檔
        $(document).ready(function () {
            $(document).on("change", "input,textarea", function () {
                if ($("#autoStatus").val() == "open") {
                    //辦理情形&進度差異說明 convert to XML post to backend
                    var xmldoc = document.createElement("root");
                    $("textarea[name='PD_Summary']").each(function (i) {
                        var xNode = document.createElement("pditem");
                        var Node = document.createElement("summary");
                        Node.textContent = this.value;
                        var Node2 = document.createElement("backward");
                        Node2.textContent = $("textarea[name='PD_BackwardDesc']")[i].value;
                        xNode.appendChild(Node);
                        xNode.appendChild(Node2);
                        xmldoc.appendChild(xNode);
                    });

                    var iframe = $('<iframe name="postiframe" id="postiframe" style="display: none" />');
                    var year = $('<input type="hidden" name="year" id="year" value="' + $.getParamValue('year') + '" />');
                    var season = $('<input type="hidden" name="season" id="season" value="' + $.getParamValue('season') + '" />');
                    var stage = $('<input type="hidden" name="stage" id="stage" value="' + $.getParamValue('stage') + '" />');
                    var form = $("form")[0];

                    //如果沒有重新導頁需要刪除上次資訊
                    $("#postiframe").remove();
                    $("input[name='year']").remove();
                    $("input[name='season']").remove();
                    $("input[name='stage']").remove();

                    form.appendChild(iframe[0]);
                    form.appendChild(year[0]);
                    form.appendChild(season[0]);
                    form.appendChild(stage[0]);

                    form.setAttribute("action", "../handler/AutoSave_Season.ashx");
                    form.setAttribute("method", "post");
                    form.setAttribute("enctype", "multipart/form-data");
                    form.setAttribute("encoding", "multipart/form-data");
                    form.setAttribute("target", "postiframe");
                    form.submit();
                }
            });
        });

        function AutoSaveFeedback(xStr) {
            if ($(xStr).find("Error").length > 0) {
                console.log("Error Message："+$(xStr).find("Error").attr("Message"));
            }
        }
    </script>
</asp:Content>

