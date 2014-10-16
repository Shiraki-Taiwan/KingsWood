var KEY_CODE_BACK_SPACE = 8;
var KEY_CODE_ENTER = 13;
var KEY_CODE_SHIFT = 16;
var KEY_CODE_CTRL = 17;
var KEY_CODE_ALT = 18;
var KEY_CODE_ESC = 27;
var KEY_CODE_SPACE = 32;
var KEY_CODE_PAGE_UP = 33;
var KEY_CODE_PAGE_DOWN = 34;
var KEY_CODE_LEFT = 37;
var KEY_CODE_UP = 38;
var KEY_CODE_RIGHT = 39;
var KEY_CODE_DOWN = 40;
var KEY_CODE_UP_0 = 48;
var KEY_CODE_UP_STAR = 56;
var KEY_CODE_A = 65;
var KEY_CODE_Q = 81;
var KEY_CODE_S = 83;
var KEY_CODE_X = 88;
var KEY_CODE_Z = 90;
var KEY_CODE_RIGHT_0 = 96;
var KEY_CODE_RIGHT_STAR = 106;
var KEY_CODE_RIGHT_DOT = 110;
var KEY_CODE_RIGHT_SLASH = 111;
var KEY_CODE_F2 = 113;
var KEY_CODE_F7 = 118;
var KEY_CODE_F8 = 119;
var KEY_CODE_F9 = 120;
var KEY_CODE_F12 = 123;

var KEY_CODE_0 = 48;
var KEY_CODE_1 = 49;
var KEY_CODE_2 = 50;
var KEY_CODE_3 = 51;
var KEY_CODE_4 = 52;
var KEY_CODE_5 = 53;
var KEY_CODE_6 = 54;
var KEY_CODE_7 = 55;
var KEY_CODE_8 = 56;
var KEY_CODE_9 = 57;

function FormatString_JS(szStr, nDesiredLen) {
    var i, nStrLen, szNewStr;
    nStrLen = String(szStr).length;
    szNewStr = "";
    for (i = nStrLen; i <= nDesiredLen - 1; i++) { szNewStr = szNewStr + "0"; }
    szNewStr = szNewStr + szStr;
    return szNewStr;
}
//-------------------- 改變游標位置 ------------------------------------
function ChangeFocus(index) {
    if (CheckKeyCode(KEY_CODE_ENTER)) {
        // check if press "Enter"
        document.form[index].focus();
    }
}
//-------------------- 改變游標位置: 方向鍵、Enter ------------------------------
function ChangeFocusByPressedKey(aIPreIndex, aINextIndex, aIEnterIndex) {
    if (CheckKeyCode(KEY_CODE_ENTER)) {
        // check if press "Enter"
        document.form[aIEnterIndex].focus();
    }
    else if (CheckKeyCode(KEY_CODE_LEFT)) {
        // 左方向鍵
        document.form[aIPreIndex].focus();
    }
    else if (CheckKeyCode(KEY_CODE_RIGHT)) {
        // 右方向鍵
        document.form[aINextIndex].focus();
    }
}
//-------------------- 全選文字 ------------------------------------
function SelectText(index) {
    document.form[index].select();
}
//-------------------- Check是否按下"Enter" ------------------------------------
function CheckKeyCode(aKeyCode) {
    window.NS = (document.layers) ? true : false;
    var code;
    if (window.NS) code = event.which; else code = event.keyCode;

    if (code == aKeyCode) return true; else return false;
}
//-------------------- Check HotKey ------------------------------------
function CheckHotKey() {
    window.NS = (document.layers) ? true : false;
    var code;
    if (window.NS) code = event.which; else code = event.keyCode;

    //if (event.ctrlKey && code == KEY_CODE_7) {
    //    // CTRL + 7
    //    // if (code == KEY_CODE_F7) // F7
    //    // 航次資料設定
    //    location.href = "../Voyage/VOfun01a.asp";
    //}
    //else if (code == KEY_CODE_BACK_SPACE) {
    //    // Back Space
    //    // 取消Windows對Back Space的預設功能 (好像回上一頁吧...)
    //    event.keyCode = 0;
    //}
    if (code == KEY_CODE_BACK_SPACE) {
        // Back Space
        // 取消Windows對Back Space的預設功能 (好像回上一頁吧...)
        event.keyCode = 0;
    }
}
//-----------------------檢查輸入的月份是否有誤-----------------------------
function CheckMonth() {
    if (document.form.Month.value == "") { return; }
    if (document.form.Month.value < 1 || document.form.Month.value > 12) {
        alert("您輸入的月份不在正常範圍內!!");
        form.Month.focus();
    }
}
//-----------------------檢查輸入的日期是否有誤-----------------------------
function CheckDay() {
    if ((document.form.Month.value == "") && (document.form.Day.value == "")) { return; }

    switch (document.form.Month.value) {
        case "1":
        case "3":
        case "5":
        case "7":
        case "8":
        case "10":
        case "12":
            {
                if (document.form.Day.value > 31) {
                    alert("您輸入的日期不在正常範圍內!!");
                    form.Day.focus();
                }

                break;
            }

        case "4":
        case "6":
        case "9":
        case "11":
            {
                if (document.form.Day.value > 30) {
                    alert("您輸入的日期不在正常範圍內!!");
                    form.Day.focus();
                }

                break;
            }

        case "2":
            {
                // 潤年
                if ((eval(document.form.Year.value) % 4) == 0) {
                    if (document.form.Day.value > 29) {
                        alert("您輸入的日期不在正常範圍內!!");
                        form.Day.focus();
                    }
                }
                    // 非潤年
                else {
                    if (document.form.Day.value > 28) {
                        alert("您輸入的日期不在正常範圍內!!");
                        form.Day.focus();
                    }
                }

                break;
            }
    }
}
//-------------------------Set UI style for the object that has focus--------------------------
function SetFocusStyle(aObjControl, aBHasFocus, aBIsButton) {
    if (aBHasFocus) {
        aObjControl.style.borderColor = '#ff0000';

        if (aBIsButton) {
            aObjControl.style.borderWidth = '0.5mm';
            aObjControl.style.borderStyle = 'outset';
        }
        else {
            aObjControl.style.borderWidth = '0.3mm';
        }
    }
    else {
        aObjControl.style.borderColor = '';

        if (aBIsButton) {
            aObjControl.style.borderWidth = '0.5mm';
            aObjControl.style.borderStyle = 'outset';
            aObjControl.style.borderColor = '#C9E0F8';
        }
        else {
            aObjControl.style.borderWidth = '';
        }
    }
}