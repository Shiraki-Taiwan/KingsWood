//-------------------- 初始游標位置 ------------------------------------
function CLocation()
{
	document.form.Search.focus();
}

//-----------------------查詢功能----------------------------------------
function OnSearch()
{
    document.form.action="FSfun02b.asp?Status=Show";
    document.form.submit();
}

//
function SearchByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
         OnSearch();
    }
    else
    {
         document.form.Delete.focus();
    }
}


//-----------------------清除功能----------------------------------------
function OnReset()
{
    document.form.action="FSfun02a.asp";
    document.form.submit();
}

//
function ResetByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        OnReset();
    }
    else
    {
        document.form.Search.focus();
    }
}


//-----------------------刪除功能----------------------------------------
function OnDelete()
{
    document.form.action="FSfun02b.asp?Status=Delete";
    document.form.submit();
}

//
function DeleteByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        On();
    }
    else
    {
        document.form.Reset.focus();
    }
}