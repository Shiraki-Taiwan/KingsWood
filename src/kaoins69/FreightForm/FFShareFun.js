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