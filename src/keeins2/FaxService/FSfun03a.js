//-------------------- 初始游標位置 ------------------------------------
function CLocation()
{
	document.form[0].focus();
}

//-------------------- Submit前欄位驗證 ----------------------------
function checkform()
{
    
}


//-----------------------查詢功能----------------------------------------
function Search()
{
    document.form.action="FSfun03b.asp";
    document.form.submit();
}

//
function SearchByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        Search();
    }
    else
    {
        document.form.ReportType.focus();
    }
}



