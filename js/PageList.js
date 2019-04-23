/*-------------------------------------
Title：分頁
Author：Nick Lai
Create Date：2018/03/16
Modify Date：2018/03/16
Version：1.0
-------------------------------------*/

//參數設定
var PageOption = {
    Selector: '', // Selector
    ListNum: 10, // 每頁顯示資料筆數,預設10
    PageNum: 10, // 分頁頁籤顯示數,預設10
    HomeBtn: true, // 顯示首頁按鈕
    LastBtn: true, // 顯示最末頁按鈕
    PrevBtn: true, // 顯示上一頁按鈕
    NextBtn: true, // 顯示下一頁按鈕
    PrevStep: true, // 顯示上一階按鈕
    NextStep: true, // 顯示下一階按鈕
    JumpVisible: false, //顯示跳頁功能
    TotalDataVisible: true, // 顯示總筆數
    TotalPageVisible: true // 顯示總頁數
}

//Math.ceil -> 無條件進位
//javascript:scroll(0,0) 回到頂端
function CreatePage(nowPage, totalData) {
    var PageListStr = '';
    var PageNum = PageOption.PageNum;
    var PagesLen = Math.ceil(totalData / PageOption.ListNum);
    if (PagesLen <= 1)
        $(PageOption.Selector).hide();
    else
        $(PageOption.Selector).show();

    //首頁
    if (PageOption.HomeBtn == true) {
        if (nowPage == 0) PageListStr += '<span>首頁</span>&nbsp;&nbsp;';
        else PageListStr += '<a href="javascript:scroll(0,0)" onclick="getData(0)">首頁</a>&nbsp;&nbsp;';
    }

    //上一階
    if (PageOption.PrevStep == true) {
        if (nowPage == 0) { PageListStr += '<span>上' + PageNum + '頁</span>&nbsp;&nbsp;' }
        else if (nowPage >= 10) { PageListStr += '<a href="javascript:scroll(0,0)" onclick="getData(' + (nowPage - PageNum) + ')">上' + PageNum + '頁</a>&nbsp;&nbsp;' }
        else { PageListStr += '<a href="javascript:scroll(0,0)" onclick="getData(0)">上' + PageNum + '頁</a>&nbsp;&nbsp;' }
    }

    //上一頁
    if (PageOption.PrevBtn == true) {
        if (nowPage != 0) { PageListStr += '<a href="javascript:scroll(0,0)" onclick="getData(' + (nowPage - 1) + ')">上一頁</a>&nbsp;&nbsp;' }
        else { PageListStr += '<span>上一頁</span>&nbsp;&nbsp;' }
    }

    //頁碼
    var startPage, endPage;
    var pn = (PageNum - 1);
    var tmp = nowPage % PageNum;

    endPage = (nowPage - tmp) + pn;
    startPage = endPage - pn;

    //最後一頁不大於總頁數
    if (endPage > (PagesLen - 1)) endPage = (PagesLen - 1);

    for (var i = startPage; i <= endPage; i++) {
        if (i == nowPage) PageListStr += '<span style="color:red;">' + (i + 1) + '</span>&nbsp;&nbsp;';
        else PageListStr += '<a href="javascript:scroll(0,0)" onclick="getData(' + i + ')">' + (i + 1) + '</a>&nbsp;&nbsp;'
    }

    //下一頁
    if (PageOption.NextBtn == true) {
        if (nowPage != (PagesLen - 1)) { PageListStr += '<a href="javascript:scroll(0,0)" onclick="getData(' + (nowPage + 1) + ')">下一頁</a>&nbsp;&nbsp;' }
        else { PageListStr += '<span>下一頁</span>&nbsp;&nbsp;' }
    }

    //下一階
    if (PageOption.NextStep == true) {
        if (nowPage == (PagesLen - 1)) { PageListStr += '<span>下' + PageNum + '頁</span>&nbsp;&nbsp;' }
        else if ((nowPage + 10) >= (PagesLen - 1)) { PageListStr += '<a href="javascript:scroll(0,0)" onclick="getData(' + (PagesLen - 1) + ')">下' + PageNum + '頁</a>&nbsp;&nbsp;' }
        else { PageListStr += '<a href="javascript:scroll(0,0)" onclick="getData(' + (nowPage + PageNum) + ')">下' + PageNum + '頁</a>&nbsp;&nbsp;' }
    }

    //最末頁
    if (PageOption.LastBtn == true) {
        if (nowPage != (PagesLen - 1)) { PageListStr += '<a href="javascript:scroll(0,0)" onclick="getData(' + (PagesLen - 1) + ')">最末頁</a>' }
        else { PageListStr += '<span class="disabled">最末頁</span>' }
    }

    //跳頁
    if (PageOption.JumpVisible == true) {
        PageListStr += '，移至 <input type="text" name="PageJump" onkeypress="JumpFun(event)" style="width:25px; height:20px;" /> 頁'
    }

    //總筆數
    if (PageOption.TotalDataVisible == true) {
        PageListStr += "，共 " + totalData + " 筆資料"
    }

    //總頁數 
    if (PageOption.TotalPageVisible == true) {
        PageListStr += "，共 " + PagesLen + " 頁"
    }

    $(PageOption.Selector).html(PageListStr);
}

//跳頁Function
function JumpFun(e) {
    if (e.keyCode == 13) {
        var p = parseInt(e.target.value) - 1;

        if (!isNaN(p))
            getData(p);
        else
            e.target.value = "";
    }
}