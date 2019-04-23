//絕對置底
$(document).ready(sizeContent);
$(window).resize(sizeContent);

function sizeContent() {
	var overHeight = $(".WrapperBody").height();// 內容大於視窗高度
    var newHeight = $("html").height() - $(".WrapperFooter").height();//內容小於視窗高度
if(overHeight > newHeight){
	$(".WrapperBody").css("min-height", overHeight + "px");
	}
	else{
	$(".WrapperBody").css("min-height", newHeight + "px");
	}
}