function upPage(p, t) {
    nowPage = p
    //javascript:scroll(0,0) <— 回到頂端
    if (nowPage == 0)
        strStart = '<span>首頁</span> ';
    else
        strStart = '<a href="javascript:scroll(0,0)" onclick="getData(0)">首頁</a> ';
    var PageNum_2 = PageNum % 2 == 0 ? Math.ceil(PageNum / 2) + 1 : Math.ceil(PageNum / 2)
    var PageNum_3 = PageNum % 2 == 0 ? Math.ceil(PageNum / 2) : Math.ceil(PageNum / 2) + 1
    var strC = "", startPage, endPage;
    if (PageNum >= PagesLen) { startPage = 0; endPage = PagesLen - 1 }
    else if (nowPage < PageNum_2) { startPage = 0; endPage = PagesLen - 1 > PageNum ? PageNum : PagesLen - 1 } //首頁 
    else { startPage = nowPage + PageNum_3 >= PagesLen ? PagesLen - PageNum - 1 : nowPage - PageNum_2 + 1; var tmp = startPage + PageNum; endPage = tmp > PagesLen ? PagesLen - 1 : tmp }
    for (var i = startPage; i <= endPage; i++) {
        if (i == nowPage) strC += '<span style="color:red;">' + (i + 1) + '</span>&nbsp;&nbsp;';
        else strC += '<a href="javascript:scroll(0,0)" onclick="getData(' + i + ')">' + (i + 1) + '</a>&nbsp;&nbsp;'
    }
    if (nowPage != PagesLen - 1) { strEnd = '<a href="javascript:scroll(0,0)" onclick="getData(' + (PagesLen - 1) + ')">最末頁</a> ' }
    else { strEnd = '<span class="disabled">最末頁</span> ' }
    strTotalData = "共 " + t + " 筆資料　"
    //strE3 = nowPage + 1 + "/" + PagesLen + "頁"
    strTotalPage = "共 " + PagesLen + " 頁"
    if (nowPage != 0) { prepage = ' <a href="javascript:scroll(0,0)" onclick="getData(' + (nowPage - 1) + ')">上一頁</a>&nbsp;&nbsp;' }
    else { prepage = ' <span>上一頁</span>&nbsp;&nbsp;' }
    if (nowPage != PagesLen - 1) { nextpage = '<a href="javascript:scroll(0,0)" onclick="getData(' + (nowPage + 1) + ')">下一頁</a> ' }
    else { nextpage = '<span>下一頁</span> ' }
    document.getElementById("changpage").innerHTML = strStart + prepage + strC + nextpage + strEnd + "，" + strTotalData + strTotalPage
}