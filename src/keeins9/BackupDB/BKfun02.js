//-------------------- 初始游標位置 ------------------------------------
function CLocation()
{
	form.Import.focus();   
}

//-------------------- Submit前欄位驗證 ----------------------------
function checkform()
{
    
}

//-----------------------匯入功能----------------------------------------
function Send()
{
    if (confirm('確定要匯入備份資料庫?'))
    {
        document.form.action="BKfun02b.asp?Status=1";
    	document.form.submit();
    }
}

//
function SendByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        Send();
    }
    else
    {
        //form.Restore.focus();
    }
}


//-----------------------回復功能----------------------------------------
function OnRestore()
{
    if (confirm('確定要回復成原始資料庫?'))
    {
        document.form.action="BKfun02b.asp?Status=2";
    	document.form.submit();
    }
}

//
function RestoreByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        OnRestore();
    }
    else
    {
        //form.import.focus();
    }
}