//-------------------- 初始游標位置 ------------------------------------
function CLocation()
{
	form.MailServer.focus();   
}

//-------------------- Submit前欄位驗證 ----------------------------
function checkform()
{
    var bFlag = 0;
    var szErrorString = "";
       
    // 驗證是否必填欄位已填
    if ( document.form.MailServer.value == "" )
    {
        bFlag = 1; 
        szErrorString = szErrorString  + "【Mail Server】尚未填寫" + "\n";
    }
    
    if ( document.form.SenderMail.value == "" )
    {
        bFlag = 1; 
        szErrorString = szErrorString  + "【Sender's Mail】尚未填寫" + "\n";
    }
    
    if ( bFlag == 1 )
    {
        alert(szErrorString); 
        return(false);
    }
    else
    {
        return(true);
    }
}

//-----------------------新增功能----------------------------------------
function Add()
{
    if (checkform())
    {
        document.form.action="FSfun04b.asp";
    	document.form.submit();
    }
}

//
function AddByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        Add();
    }
    else
    {
        ChangeFocus(3);
    }
}

//-----------------------清除功能----------------------------------------
function Reset()
{
    document.form.MailServer.value   = "";
	document.form.SenderMail.value = "";
	
	document.form.action="FSfun04a.asp";
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