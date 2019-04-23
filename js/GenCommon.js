// JavaScript Document
$(document).ready(function(){
//針對IE版本做處理(ex:.ie8 .tabmenublock{padding-bottom:1px;})
if (/*@cc_on!@*/false) {
     document.documentElement.className += ' ie' + document.documentMode;
    }
if (/*@cc_on!@*/true) {
     document.documentElement.className += ' ie' + document.documentMode;
    }
//雙色表單
$(".stripeMe table:not(td > table) > tbody > tr:even").addClass("alt");
//$(".stripeMe tr").mouseover(function() {$(this).addClass("over");}).mouseout(function() {$(this).removeClass("over");});
//資料列表
$(".gentable tr:last-child td").css("border-bottom-width","0");
$(".OchiFixTable .OchiFixRow:last-child").css("border-bottom-width","0");
$(".OchiTrasTable .OchiTrasRow:last-child").css("border-bottom-width","0");
});