// JavaScript Document
$(document).ready(MainFunction);//主程式:所有JS程式放在此
//主程式內容
	function MainFunction(){
	clonemenu();//RWD選單處理
	colorboxFun()//colorbox修改	
	
	$('.itemhint').powerTip({placement:'e',smartPlacement:true,});//powertip:tooltip
	$('.itemhinthold').powerTip({placement:'e',smartPlacement:true,mouseOnToPopup:'true'});//tooltip可點選
	//flatpickr datepicker
	/*
	$(".Jdatepicker").flatpickr({
		"locale": "zh",
		});
	*/
	$(".Jdatepicker").datetimepicker({
		format:'Y/m/d',
        timepicker: false,
		withoutBottomPanel: true,
		scrollInput:false,
        yearEnd: 2020
		});
	

	//網站內容預設為桌機版(no-js)下狀態,RWD平板、手機內容以動態方式取代加入,故平板、手機內容需在JS內以變數方式設定
	var Htmlsidemenu = '<img src="../App_Themes/images/icon-menu.png" id="toggle-sidebar" />';
	//設定斷點:判斷螢幕尺寸決定內容
	//桌機狀態:瀏覽器大於 980
	$.breakpoint({
        condition: function () {
            return window.matchMedia('all and (min-width:981px)').matches;
        },
        enter: function () {
            $("#mainmenu").show();
			$("#opensidemenu").html("");
			
        },
    });
	
	//平板手機狀態:瀏覽器介於 980~480
	$.breakpoint({
        condition: function () {
            return window.matchMedia('all and (max-width:980px) and (min-width:480px)').matches;
        },
        enter: function () {
            $("#mainmenu").hide();
			$("#opensidemenu").html(Htmlsidemenu);

        },

    });
	
	//手機狀態:瀏覽器小於 480 且為橫式
	$.breakpoint({
	condition: function () {
            return window.matchMedia('all and (max-width:479px) and (orientation:landscape)').matches;
        },
        enter: function () {
            $("#mainmenu").hide();
			$("#opensidemenu").html(Htmlsidemenu);
		
        },
	 });
	 
	 //手機狀態:尺寸小於 480 且為直立 orientation:portrait
	$.breakpoint({
        condition: function () {
            return window.matchMedia('all and (max-width:479px) and (orientation:portrait)').matches;
        },
        enter: function () {
            $("#mainmenu").hide();
			$("#opensidemenu").html(Htmlsidemenu);

        },

    });
	
	}//MainFunction

	
//RWD選單處理:桌機使用superfish，手機平板使用simplerSidebar，而simplerSidebar內的選單為複製主選單內容。
	function clonemenu(){
	//複製選單到sidebar
	var clonemenu = $("#mainmenu").clone(false);
	//移除superfish的id與class並給予新id
	clonemenu.remove("#mainmenu").appendTo($("#sidebar-wrapper")).attr("id", "sidemenu").removeClass("sf-menu");
	//啟動下拉選單superfish
	$("#mainmenu").superfish({delay:300,}).supposition();
	//啟動mmenu
	$("#sidebar-wrapper").mmenu({
		//設定下拉選單為直接向下展開,而非滑動
		slidingSubmenus:false,
		});
	var mmenuswitch = $("#sidebar-wrapper").data( "mmenu" ); 
	  
	  //由於開關是動態加入(利用JS加入),故要使用動態binding的方式加入函式動作
	  $(document).on("click", "#toggle-sidebar", function(){
		  mmenuswitch.open();
		});
	  
	//利用 設定滑動關閉選單
	$("#sidebar-wrapper").swipe({
	swipe:function(event, direction, distance, duration, fingerCount) {//事件，方向，距離（像素為單位），時間，手指數量
		if(direction == "left")//當向上滑動手指時令當前頁面記數器加1
			{mmenuswitch.close();}
		}
	});
	}//clonemenu END


//colorbox修改
function colorboxFun(){
	$(".colorboxS").colorbox({inline:true, width:"70%", maxWidth:"600", maxHeight:"90%", opacity:0.5});
	$(".colorboxM").colorbox({inline:true, width:"80%", maxWidth:"800", maxHeight:"90%", opacity:0.5});
	$(".colorboxL").colorbox({inline:true, width:"90%", maxWidth:"900", maxHeight:"90%", opacity:0.5});
	$(".colorboxiframe").colorbox({iframe:true, innerWidth:"900",innerHeight:"600",opacity:0.5});
	//修改外框
	$("#cboxTopLeft").hide();
	$("#cboxTopRight").hide();
	$("#cboxBottomLeft").hide();
	$("#cboxBottomRight").hide();
	$("#cboxMiddleLeft").hide();
	$("#cboxMiddleRight").hide();
	$("#cboxTopCenter").hide();
	$("#cboxBottomCenter").hide();
	$("#cboxContent").addClass("colorboxnewborder");
	//修正powertip與colorbox搶title問題
	$(".colorboxS,.colorboxM,.colorboxL,.colorboxGen,.colorboxiframe").colorbox({
		title: function(){
  		return $(this).attr('data-colorboxtitle');
	}});
	//colorobx關閉按扭
	$(".closecolorbox").click(function(){
		$.colorbox.close()
		})
}

