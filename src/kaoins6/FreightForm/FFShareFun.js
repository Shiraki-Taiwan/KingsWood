var g_bNeededForestryHasFocus = false;
var g_bPredictVolumeHasFocus = false;
var g_bWeightHasFocus = false;
var g_bTotalWeightHasFocus = false;
var g_bF2Counter = 0;
function sleep(milliseconds) {
    var start = new Date().getTime();
    for (var i = 0; i < 1e7; i++) {
        if ((new Date().getTime() - start) > milliseconds) {
            break;
        }
    }
}
//
function UpdateNeededForestryHasFocus() {
    g_bNeededForestryHasFocus = true;
}

function UpdateNeededForestryLoseFocus() {
    g_bNeededForestryHasFocus = false;
}


//
function UpdatePredictVolumeHasFocus() {
    g_bPredictVolumeHasFocus = true;
}

function UpdatePredictVolumeLoseFocus() {
    g_bPredictVolumeHasFocus = false;
}

//
function UpdateWeightHasFocus() {
    g_bWeightHasFocus = true;
}

function UpdateWeightLoseFocus() {
    g_bWeightHasFocus = false;
}

//
function UpdateTotalWeightHasFocus() {
    g_bTotalWeightHasFocus = true;
}

function UpdateTotalWeightLoseFocus() {
    g_bTotalWeightHasFocus = false;
}

//-----------------------顯示正確的IsPL文字----------------------------------------
function OnIsPL(a_iIndex, a_iCurElement, a_iNextElement) {
    var bYes, bNo, nKeyCode;
    var bChanged;
    bChanged = false;

    var nTotalLineCounter;
    // 必須隨著實際行數而定
    nTotalLineCounter = 50;

    //if (CheckKeyCode(KEY_CODE_UP_0) || CheckKeyCode(KEY_CODE_RIGHT_0)) // Check if press "0": 否
    if (CheckKeyCode(KEY_CODE_UP_STAR) || CheckKeyCode(KEY_CODE_RIGHT_STAR)) // Check if press "*": 是
    {
        nKeyCode = 1;
    }
    else if (CheckKeyCode(KEY_CODE_UP_0) || CheckKeyCode(KEY_CODE_RIGHT_0))   // Check if press "0": 否=>空白
    {
        nKeyCode = 0;
    }


    //先把所有的設成 unselected
    document.form[a_iCurElement].options[0].selected = false;
    document.form[a_iCurElement].options[1].selected = false;

    if (nKeyCode >= 0 && nKeyCode <= 1) {
        document.form[a_iCurElement].options[nKeyCode].selected = true;
    }

    ChangeFocus(a_iNextElement);
}


//-----------------------顯示正確的包裝文字----------------------------------------
function OnPackageStyleID(a_nPackageStyleCount, a_iIndex, a_iCurElement, a_iNextElement) {
    var bYes, bNo, nKeyCode;
    var i;

    for (i = 48; i < 58; i++) {
        if (CheckKeyCode(i) || CheckKeyCode(i + 48)) {
            nKeyCode = i - 48;
            break;
        }
    }

    if (nKeyCode >= 0 && nKeyCode <= a_nPackageStyleCount) {
        //先把所有的設成 unselected
        for (i = 0 ; i < a_nPackageStyleCount; i++) {
            document.form[a_iCurElement].options[i].selected = false;
        }

        document.form[a_iCurElement].options[nKeyCode].selected = true;
    }

    ChangeFocus(a_iNextElement);
}


//-----------------------存檔熱鍵---------------------------------------
function CheckSaveHotKey() {

    if (CheckKeyCode(KEY_CODE_F2)) // F2
    {
        g_bF2Counter = g_bF2Counter + 1;

        if (document.form.IsChecked.value == 0) {
            if (g_bF2Counter == 1) {
                OnSave();

                // 妨止按太多次F2
                sleep(1000);
                g_bF2Counter = 0;
            }
        }
    }

}

function CheckSaveHotKey1() {

    if (CheckKeyCode(KEY_CODE_F2)) // F2
    {
        g_bF2Counter = g_bF2Counter + 1;

        if (g_bF2Counter == 1) {
            OnSave();

            // 妨止按太多次F2
            sleep(1000);
            g_bF2Counter = 0;
        }
    }

}


//-----------------------擋掉某些鍵, 使其不作用----------------------------------
function CheckForbiddenKey() {
    if (CheckKeyCode(KEY_CODE_RIGHT_SLASH)) {
        return false;
    }
    else if (CheckKeyCode(KEY_CODE_RIGHT_DOT)) {
        // 當focus在需求才積或預測體積時, 不擋小數點
        // [2007.03.31] 單重與總重也不擋小數點
        if (g_bNeededForestryHasFocus || g_bPredictVolumeHasFocus || g_bWeightHasFocus || g_bTotalWeightHasFocus) {
            return true;
        }
        else {
            return false;
        }
    }
    else if (CheckKeyCode(KEY_CODE_UP) || CheckKeyCode(KEY_CODE_DOWN)) {
        return false;
    }
    else {
        return true;
    }

}


function CheckForbiddenKey_UpAndDown() {
    if (CheckKeyCode(KEY_CODE_UP) || CheckKeyCode(KEY_CODE_DOWN)) {
        alert(document.form.IsPL_0.value);
        return false;
    }
    else {
        return true;
    }

}

function VolumeCalculator2(a_nIndex) {

    var fVolumn = 0;
    var DivNum = 1000000;
    var nBoard = 1, nPiece = 1, nRow = a_nIndex * nFieldDiff;
    var nBoardValue =
        document.form[11 + nRow].value;
    var bIsPL =
        document.form[12 + nRow].value;
    var nPieceValue =
        document.form[13 + nRow].value;
    var nLength =
        document.form[15 + nRow].value
    var nWidth =
        document.form[16 + nRow].value
    var nHeight =
        document.form[17 + nRow].value

    if (nBoardValue == "") { nBoard = 1 }
    else { nBoard = parseInt(nBoardValue) }

    if (nPieceValue == "") { nPiece = 1 }
    else { nPiece = parseInt(nPieceValue) }

    if (nLength != "" && nWidth != "" && nHeight != "") {
        fVolumn = nLength * nWidth * nHeight / DivNum;

        if (bIsPL == "0") {
            // 不是堆量，要乘件數
            fVolumn = fVolumn * nPiece;
        } else {
            // 堆量, 要乘板數
            fVolumn = fVolumn * nBoard;
        }
        document.form[18 + nRow].value = formatFloat(fVolumn, 2);

        ChangeTextColor2(a_nIndex, nLength, nWidth, nHeight, fVolumn, bIsPL);
    }

    // 09-May2005: update上方的總件數與總體積
    UpdateTotalPieceAndVolume();
}

// -----------------------update上方的總件數與總體積----------------------------------------
function UpdateTotalPieceAndVolume() {
    var i, nPieceSum = 0, fVolumeSum = 0, nPiece, bSkipAnyPiece = false;

    for (i = 0; i < document.form.DataCounter.value; i++) {
        nPiece = document.form[13 + i * nFieldDiff].value;

        if (nPiece == "") {
            nPiece = 0
            bSkipAnyPiece = true
        }

        nPieceSum += parseFloat(nPiece);
        fVolumeSum += parseFloat(document.form[18 + i * nFieldDiff].value);
    }
    document.form.StoreSum_Piece.value = nPieceSum;

    if (bSkipAnyPiece) {
        document.all.td_TotalPiece.innerHTML = '<font size="6"></font>';
        document.all.td_TotalVolume.innerHTML = '<font size="6"></font>';
    } else {
        document.all.td_TotalPiece.innerHTML = '<font size="6">' + document.form.StoreSum_Piece.value + '</font>';

        if (document.form.NeededForestry.value == "") {
            fVolumeSum = formatFloat(fVolumeSum, 2);
            document.form.StoreSum_Volume.value = fVolumeSum;
            document.all.td_TotalVolume.innerHTML = "<font size=6>" + document.form.StoreSum_Volume.value + "</font>";
        } else {
            document.all.td_TotalVolume.innerHTML = "<font size=6 color=#FF0000>" + formatFloat(document.form.NeededForestry.value, 2) + "</font>";
        }
    }
}
function isNumber(n) {
    return !isNaN(parseFloat(n)) && isFinite(n);
}
function formatFloat(num, pos) {
    var size = Math.pow(10, pos);
    return Math.round(num * size) / size;
}
function ChangeTextColor2(a_nIndex, nLength, nWidth, nHeight, fVolume, bIsPLValue) {
    var ColorYellow = "#FFFFA0", ColorWhite = "#ffffff", ColorPink = "#FFC8FF";
    var bChangeBg = 0, bIsPL;

    if (bIsPLValue == "") { bIsPL = 0 }
    else { bIsPL = Boolean(bIsPLValue) }

    // 若有板數,體積先除以板數
    var nBoard;

    if (document.form[11 + a_nIndex * nFieldDiff].value != "") {
        nBoard = document.form[11 + a_nIndex * nFieldDiff].value

        if (isNumber(nBoard)) {
            nBoard = parseInt(nBoard);
        } else {
            // just for aborting error
            nBoard = 1;
        }
        if (nBoard > 0) {
            fVolume = fVolume / nBoard;
        }
    }

    // 體積大於3.8, 或小於0.1 用不同頻色表示
    var ColorVolume = "";
    if ((fVolume > 3.8) || (fVolume < 0.1)) {
        bChangeBg = 1; ColorVolume = ColorPink;
    }

    // 27-Apr2005: 若有堆量, 且長寬高小於30cm, 變色
    var ColorLength = "";
    if (nLength > 600 || nLength < 10 || (bIsPL = 1 && nLength < 30)) {
        bChangeBg = 1; ColorLength = ColorPink;
    }

    var ColorWidth = "";
    if (nWidth > 600 || nWidth < 10 || (bIsPL = 1 && nWidth < 30)) {
        bChangeBg = 1; ColorWidth = ColorPink;
    }

    var ColorHeight = "";
    if (nHeight > 226 || nHeight < 10 || (bIsPL = 1 && nHeight < 30)) {
        bChangeBg = 1; ColorHeight = ColorPink;
    }

    if (bChangeBg == 1) {
        var nRow = a_nIndex * nFieldDiff;

        document.form[8 + nRow].style.backgroundColor = ColorYellow;
        document.form[9 + nRow].style.backgroundColor = ColorYellow;
        document.form[10 + nRow].style.backgroundColor = ColorYellow;
        document.form[11 + nRow].style.backgroundColor = ColorYellow;
        document.form[12 + nRow].style.backgroundColor = ColorYellow;
        document.form[13 + nRow].style.backgroundColor = ColorYellow;
        document.form[14 + nRow].style.backgroundColor = ColorYellow;

        if (ColorLength == "") {
            document.form[15 + nRow].style.backgroundColor = ColorYellow;
        } else {
            document.form[15 + nRow].style.backgroundColor = ColorLength;
        }

        if (ColorWidth == "") {
            document.form[16 + nRow].style.backgroundColor = ColorYellow;
        } else {
            document.form[16 + nRow].style.backgroundColor = ColorWidth;
        }

        if (ColorHeight == "") {
            document.form[17 + nRow].style.backgroundColor = ColorYellow;
        } else {
            document.form[17 + nRow].style.backgroundColor = ColorHeight;
        }

        if (ColorVolume == "") {
            document.form[18 + nRow].style.backgroundColor = ColorYellow;
        } else {
            document.form[18 + nRow].style.backgroundColor = ColorPink;
        }

        document.form[19 + nRow].style.backgroundColor = ColorYellow;
        document.form[20 + nRow].style.backgroundColor = ColorYellow;
    } else {
        var nRow = a_nIndex * nFieldDiff;
        document.form[8 + nRow].style.backgroundColor = ColorWhite
        document.form[9 + nRow].style.backgroundColor = ColorWhite
        document.form[10 + nRow].style.backgroundColor = ColorWhite
        document.form[11 + nRow].style.backgroundColor = ColorWhite
        document.form[12 + nRow].style.backgroundColor = ColorWhite
        document.form[13 + nRow].style.backgroundColor = ColorWhite
        document.form[14 + nRow].style.backgroundColor = ColorWhite
        document.form[15 + nRow].style.backgroundColor = ColorWhite
        document.form[16 + nRow].style.backgroundColor = ColorWhite
        document.form[17 + nRow].style.backgroundColor = ColorWhite
        document.form[18 + nRow].style.backgroundColor = ColorWhite
        document.form[19 + nRow].style.backgroundColor = ColorWhite
        document.form[20 + nRow].style.backgroundColor = ColorWhite
    }
}