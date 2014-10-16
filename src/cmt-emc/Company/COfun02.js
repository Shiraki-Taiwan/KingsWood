//-------------------- 初始游標位置 ------------------------------------
function CLocation()
{
	form.ID.focus();   
}

//-----------------------查詢功能----------------------------------------
function Search()
{
    document.form.action="COfun02b.asp";
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
        ChangeFocus(4);
    }
}

//-----------------------清除功能----------------------------------------
function Reset()
{
    document.form.ID.value   = "";
	document.form.ChineseName.value = "";
	document.form.EnglishName.value = "";
    
	document.form.action="COfun02a.asp";
	document.form.submit();
}


//
function ResetByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        Reset();
    }
    else
    {
        ChangeFocus(0);
    }
}