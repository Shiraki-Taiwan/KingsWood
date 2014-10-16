//-------------------- 初始游標位置 ------------------------------------
function CLocation()
{
    document.form.Back.focus();
}


//-----------------------回上一頁功能----------------------------------------
function OnBack()
{
    document.form.action="FFfun03b.asp";
	document.form.submit();
}

//
function BackByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        OnBack();
    }
    else
    {
        document.form[0].focus();
    }
}
