//-------------------- 初始游標位置 ------------------------------------
function OnLoad() {
	document.form.ID.focus();
}

//-------------------- Submit前欄位驗證 ----------------------------
function checkform() {
	var bFlag = 0;
	var szErrorString = "";

	// 驗證是否必填欄位已填
	if (document.form.ID.value == "") {
		bFlag = 1;
		szErrorString = szErrorString + "【帳號】尚未填寫" + "\n";
	}

	if (document.form.Password.value == "") {
		bFlag = 1;
		szErrorString = szErrorString + "【密碼】尚未填寫" + "\n";
	}

	if (bFlag == 1) {
		alert(szErrorString);
		return (false);
	}
	else {
		return (true);
	}
}

//-----------------------新增功能----------------------------------------
function OnLogin() {
	if (checkform()) {
		$.post("Login.aspx", $('form').serialize(), function () {
			document.form.action = "Loginb.asp";
			document.form.submit();
		});
	}
}

//
function OnLoginByKeyPress() {
	if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
	{
		OnLogin();
	}
	else {
		ChangeFocus(3);
	}
}

//-----------------------清除功能----------------------------------------
function Reset() {
	document.form.ID.value = "";
	document.form.Password.value = "";

	document.form.action = "Login.asp";
	document.form.submit();
}

//
function ResetByKeyPress() {
	if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
	{
		Reset();
	}
	else {
		ChangeFocus(0);
	}
}