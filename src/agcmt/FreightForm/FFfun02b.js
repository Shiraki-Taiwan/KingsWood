//-------------------- 初始游標位置 ------------------------------------
function CLocation() {
    if (!document.form) return;
    if (!document.form.ID) return;
    document.form.ID.focus();
}
//-----------------------重新查詢功能----------------------------------------
function OnReSearch() {
    document.form.VesselID.value = "";
    document.form.VesselLine.value = "";
    document.form.action = "FFfun02a.asp";
    document.form.submit();
}
function ReSearchByKeyPress() {
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        OnReSearch();
    }
    else {
        document.form[0].focus();
    }
}
//-----------------------核定功能----------------------------------------
function OnCheckIn() {
    document.form.Status.value = "CheckIn"
    document.form.action = "FFfun02c.asp";
    document.form.submit();
}
function CheckInByKeyPress() {
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        OnCheckIn();
    }
    else {
        document.form.Save.focus();
    }
}
//-----------------------取消核定功能----------------------------------------
function OnDelCheckIn() {
    document.form.Status.value = "DelCheckIn"
    document.form.action = "FFfun02c.asp";
    document.form.submit();
}
function DelCheckInByKeyPress() {
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        OnDelCheckIn();
    }
    else {
        document.form.ReSearch.focus();
    }
}
//-----------------------儲存功能----------------------------------------
function OnSave() {
    document.form.Status.value = "Modify"
    document.form.action = "FFfun02c.asp";
    document.form.submit();
}
function SaveByKeyPress() {
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        OnSave();
    }
    else {
        document.form.Move.focus();
    }
}
//-----------------------查詢倉單功能----------------------------------------
function SearchForm() {
    if (CheckKeyCode(KEY_CODE_ENTER)) // Check if press "Enter"
    {
        document.form.Status.value = "Search"
        document.form.action = "FFfun02b.asp";
        document.form.submit();
    }
    else if (CheckKeyCode(KEY_CODE_UP)) // 按"上"鍵
    {
        if (document.form.ID.value != "") {
            document.form.Status.value = "Search"
            document.form.action = "FFfun02b.asp?Find=Prev";
            document.form.submit();
        }
    }
    else if (CheckKeyCode(KEY_CODE_DOWN)) // 按"下"鍵
    {
        if (document.form.ID.value != "") {
            document.form.Status.value = "Search"
            document.form.action = "FFfun02b.asp?Find=Next";
            document.form.submit();
        }
    }
    //else if(CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    //{
    //    OnCheckIn();
    //}
}
function CheckIDLen() {
    if (CheckKeyCode(KEY_CODE_UP) ||
       CheckKeyCode(KEY_CODE_DOWN) ||
       CheckKeyCode(KEY_CODE_LEFT) ||
       CheckKeyCode(KEY_CODE_RIGHT) ||
       CheckKeyCode(KEY_CODE_PAGE_UP) ||
       CheckKeyCode(KEY_CODE_PAGE_DOWN) ||
       CheckKeyCode(KEY_CODE_ESC) ||
       CheckKeyCode(KEY_CODE_F2) ||
       CheckKeyCode(KEY_CODE_F7) ||
       CheckKeyCode(KEY_CODE_F8) ||
       CheckKeyCode(KEY_CODE_F9) ||
       CheckKeyCode(KEY_CODE_SHIFT) ||
       CheckKeyCode(KEY_CODE_CTRL) ||
       CheckKeyCode(KEY_CODE_ALT) ||
       CheckKeyCode(KEY_CODE_A) ||
       CheckKeyCode(KEY_CODE_S) ||
       CheckKeyCode(KEY_CODE_Z) ||
       CheckKeyCode(KEY_CODE_X) ||
       CheckKeyCode(KEY_CODE_RIGHT_DOT) ||
       CheckKeyCode(KEY_CODE_RIGHT_SLASH)
       ) {

    }
    else {
        // 18-Nov2004:若已輸入四碼, 直接查詢
        var strLen = document.form.ID.value.length

        if (strLen == 4) {
            document.form.Status.value = "Search"
            document.form.action = "FFfun02b.asp?PrevFoundID=" + document.form.PrevFoundID.value;
            document.form.submit();
        }
    }
}
//-----------------------預測體積功能----------------------------------------
function OnPredictVolume() {
    if (CheckKeyCode(KEY_CODE_ENTER)) // Check if press "Enter"
    {
        document.form.CheckIn.focus();
    }
}
//-----------------------刪除功能----------------------------------------
function OnDelete() {
    if (window.confirm("確定要刪除嗎？")) {
        document.form.Status.value = "Delete"
        document.form.action = "FFfun02c.asp";
        document.form.submit();
    }
}
function DeleteByKeyPress() {
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        OnDelete();
    }
    else {
        document.form.ReSearch.focus();
    }
}
//-----------------------移動功能----------------------------------------
function OnMove() {

    var i, nDataCounter, szStr1, szStr2, iIndex, nCounter;
    nDataCounter = document.form.DataCounter.value;

    iIndex = 6;
    nCounter = 0;
    szStr2 = "";

    for (i = 0; i < nDataCounter; i++) {
        if (document.form[iIndex + i * 15].checked == true) {
            szStr2 = szStr2 + "&SN_" + String(nCounter) + "=" + document.form[iIndex + i * 15 - 1].value
            nCounter = nCounter + 1;
        }
    }

    szStr = "FFfun02d.asp?FormID=" + document.form.ID_0.value;
    szStr = szStr + "&OriginalVesselListID=" + document.form.VesselListID.value;
    szStr = szStr + "&DataCounter=" + String(nCounter) + szStr2;

    document.form.Status.value = "Move"
    document.form.VesselID.value = ""
    document.form.action = szStr;
    document.form.submit();
}
function MoveByKeyPress() {
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        OnMove();
    }
    else {
        document.form.Delete.focus();
    }
}
// -----------------------for最左邊的欄位-----------------------
function OnDirectionKey_LField(a_iCurElement, a_nDataCounter, a_iLineCounter) {
    var nMaxDataCount;
    nMaxDataCount = 14 + (a_nDataCounter - 2) * 15;

    if (CheckKeyCode(KEY_CODE_LEFT))  // 左
    {
        if (a_iCurElement > 7) {
            document.form[a_iCurElement - 3].focus();
        }
    }
    else if (CheckKeyCode(KEY_CODE_RIGHT))  // 右
    {
        document.form[a_iCurElement + 1].focus();
    }
    else if (CheckKeyCode(KEY_CODE_UP))  // 上
    {
        if (a_iLineCounter > 0) {
            document.form[a_iCurElement - 15].focus();
        }
    }
    else if (CheckKeyCode(KEY_CODE_DOWN))  // 下
    {
        if (a_iLineCounter < (a_nDataCounter - 1)) {
            document.form[a_iCurElement + 15].focus();
        }
    }
}
// -----------------------for最右邊的欄位-----------------------
function OnDirectionKey_RField(a_iCurElement, a_nDataCounter, a_iLineCounter) {
    var nMaxDataCount;
    nMaxDataCount = a_nDataCounter * 15 - 1;

    if (CheckKeyCode(KEY_CODE_LEFT))  // 左
    {
        document.form[a_iCurElement - 1].focus();
    }
    else if (CheckKeyCode(KEY_CODE_RIGHT))  // 右
    {
        if (a_iCurElement < nMaxDataCount) {
            document.form[a_iCurElement + 3].focus();
        }
    }
    else if (CheckKeyCode(KEY_CODE_UP))  // 上
    {
        if (a_iLineCounter > 0) {
            document.form[a_iCurElement - 15].focus();
        }
    }
    else if (CheckKeyCode(KEY_CODE_DOWN))  // 下
    {
        if (a_iLineCounter < (a_nDataCounter - 1)) {
            document.form[a_iCurElement + 15].focus();
        }
    }
}
// -----------------------for中間的欄位-----------------------
function OnDirectionKey_MField(a_iCurElement, a_nDataCounter, a_iLineCounter) {
    var nMaxDataCount;
    nMaxDataCount = 14 + (a_nDataCounter - 1) * 15 + 4;

    if (CheckKeyCode(KEY_CODE_LEFT))  // 左
    {
        document.form[a_iCurElement - 1].focus();
    }
    else if (CheckKeyCode(KEY_CODE_RIGHT))  // 右
    {
        document.form[a_iCurElement + 1].focus();
    }
    else if (CheckKeyCode(KEY_CODE_UP))  // 上
    {
        if (a_iLineCounter > 0) {
            document.form[a_iCurElement - 15].focus();
        }
    }
    else if (CheckKeyCode(KEY_CODE_DOWN))  // 下
    {
        if (a_iLineCounter < (a_nDataCounter - 1)) {
            document.form[a_iCurElement + 15].focus();
        }
    }
}
function OnDirectionKey_MField_OnlyRL(a_iCurElement, a_nDataCounter, a_iLineCounter) {
    var nMaxDataCount;
    nMaxDataCount = 14 + (a_nDataCounter - 1) * 15 + 4;

    if (CheckKeyCode(KEY_CODE_LEFT))  // 左
    {
        document.form[a_iCurElement - 1].focus();
    }
    else if (CheckKeyCode(KEY_CODE_RIGHT))  // 右
    {
        document.form[a_iCurElement + 1].focus();
    }
}
//-----------------------進入欄位"頁次"時----------------------------------------
function OnPageNoFocus(a_nCurIndex) {
    if (a_nCurIndex > 19)    //第二行之後才執行
    {
        document.form[a_nCurIndex].value = document.form[a_nCurIndex - 15].value;
        document.form[a_nCurIndex + 1].value = document.form[a_nCurIndex - 14].value;
    }
}
//-----------------------偵測HotKey----------------------------------------
function CheckPrivateHotKey() {
    //偵測是否要switch到新增倉單
    if (CheckKeyCode(KEY_CODE_F8)) {
        document.form.action = "FFfun01a.asp?Status=Add&VesselID=" + document.form.VesselListID.value + "&PrevFoundID=" + document.form.PrevFoundID.value;
        document.form.submit();
    }
        //按空白健核定功能
    else if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        OnCheckIn();
    }
        //按Page Up換頁功能
    else if (CheckKeyCode(KEY_CODE_PAGE_UP)) {
        if (document.form.ID.value != "") {
            var nNewPageNum;
            nNewPageNum = Number(document.form.PageNum.value) - 1;

            if (nNewPageNum >= 1) {
                document.form.Status.value = "Search"
                document.form.action = "FFfun02b.asp?ID" + document.form.ID.value + "&VesselListID=" + document.form.VesselListID.value + "&PageNum=" + nNewPageNum;
                document.form.submit();
            }
        }
    }
        //按Page DOWN換頁功能
    else if (CheckKeyCode(KEY_CODE_PAGE_DOWN)) {
        if (document.form.ID.value != "") {
            var nNewPageNum;
            nNewPageNum = Number(document.form.PageNum.value) + 1;

            if (nNewPageNum <= document.form.TotalPageNum.value) {
                document.form.Status.value = "Search"
                document.form.action = "FFfun02b.asp?ID=" + document.form.ID.value + "&VesselListID=" + document.form.VesselListID.value + "&PageNum=" + nNewPageNum;
                document.form.submit();
            }
        }
    }
    else if (CheckKeyCode(KEY_CODE_ESC)) {
        if (document.form.IsChecked.value == 0) {
            document.form.Piece_0.focus();
        }
        else {
            alert("請取消核定再修改!");
        }
    }
    else if (CheckKeyCode(KEY_CODE_A)) // Ctrl+A
    {
        if (event.ctrlKey) {
            // e-mail全部
            location.href = "../FaxService/FSfun01c.asp?ReportType=0&SelectType=2&HotKey=MailAll&VesselLine=" + document.form.VesselLine.value + "&VesselListID=" + document.form.VesselListID.value;
        }
    }
    else if (CheckKeyCode(KEY_CODE_S)) // Ctrl+S
    {
        if (event.ctrlKey) {
            // e-mail新增
            location.href = "../FaxService/FSfun01c.asp?ReportType=0&SelectType=2&HotKey=MailNew&VesselLine=" + document.form.VesselLine.value + "&VesselListID=" + document.form.VesselListID.value;
        }
    }
    else if (CheckKeyCode(KEY_CODE_Z)) // Ctrl+Z
    {
        if (event.ctrlKey) {
            // 總表查詢
            location.href = "../FaxService/FSfun03c.asp?ReportType=1&VesselLine=" + document.form.VesselLine.value + "&VesselListID=" + document.form.VesselListID.value;
        }
    }
    else if (CheckKeyCode(KEY_CODE_X)) // Ctrl+X
    {
        if (event.ctrlKey) {
            // 詳細尺寸查詢
            location.href = "../FaxService/FSfun03c.asp?ReportType=2&VesselLine=" + document.form.VesselLine.value + "&VesselListID=" + document.form.VesselListID.value;
        }
    }
}