//-------------------- 初始游標位置 ------------------------------------
function CLocation() {
	document.form[0].focus();
}


//-----------------------重新查詢功能----------------------------------------
function OnReSearch() {
	document.form.VesselListID.value = "";
	document.form.VesselLine.value = "";
	document.form.action = "FSfun03b.asp";
	document.form.submit();
}

//
function ReSearchByKeyPress() {
	if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
	{
		OnReSearch();
	}
	else {
		document.form[0].focus();
	}
}


//-----------------------偵測HotKey----------------------------------------
function CheckPrivateHotKey() {
	if (CheckKeyCode(KEY_CODE_Z)) // Ctrl+Z
	{
		if (event.ctrlKey) {
			// 總表查詢
			location.href = "FSfun03c.asp?ReportType=1&VesselLine=" + document.form.VesselLine.value + "&VesselListID=" + document.form.VesselListID.value;
		}
	}
	else if (CheckKeyCode(KEY_CODE_X)) // Ctrl+X
	{
		if (event.ctrlKey) {
			// 詳細尺寸查詢
			location.href = "FSfun03c.asp?ReportType=2&VesselLine=" + document.form.VesselLine.value + "&VesselListID=" + document.form.VesselListID.value;
		}
	}
	else if (CheckKeyCode(KEY_CODE_F2)) {
		// 儲存
		OnSave();
	}
	else if (CheckKeyCode(KEY_CODE_F8)) {
		OnReSearch();
	}
}


//-----------------------儲存功能----------------------------------------
function OnSave() {
	document.form.action = "FSfun03e.asp";
	document.form.submit();
}

//
function SaveByKeyPress() {
	if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
	{
		OnSave();
	}
	else {
		// 15-Mar2005: 按Enter不要改變focus
		//document.form.BackBtn.focus();
	}
}