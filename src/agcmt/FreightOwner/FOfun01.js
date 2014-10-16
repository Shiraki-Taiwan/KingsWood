//-------------------- 初始游標位置 ------------------------------------
function CLocation()
{
	form.ID.focus();   
}

//-------------------- Submit前欄位驗證 ----------------------------
function checkform()
{
    var bFlag = 0;
    var szErrorString = "";
       
    // 驗證是否必填欄位已填
    if ( document.form.ID.value == "" )
    {
        bFlag = 1; 
        szErrorString = szErrorString  + "【編號】尚未填寫" + "\n";
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
        document.form.action="FOfun01b.asp";
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
        ChangeFocus(12);
    }
}

//-----------------------查詢功能----------------------------------------
function Search()
{
    document.form.action="FOfun01a.asp?Status=Search";
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
        ChangeFocus(13);
    }
}
//-----------------------刪除功能----------------------------------------
function Delete()
{
    if (document.form.Status.value == "Modify") // 只有在Modify的狀況下才能修改或刪除
    {
        if (window.confirm("確定要刪除嗎？"))
    	{
    		document.form.action="FOfun01b.asp?Status=Delete";
    	    document.form.submit();
    	}
    }
}

//
function DeleteByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        Delete();
    }
    else
    {
        ChangeFocus(14);
    }
}

//-----------------------清除功能----------------------------------------
function Reset()
{
    document.form.ID.value          = "";
	document.form.Name.value        = "";
    document.form.FaxNo_1.value     = "";
	document.form.FaxNo_2.value     = "";
	document.form.Contact_1.value   = "";
	document.form.Contact_2.value   = "";
	document.form.Phone_1.value     = "";
	document.form.Phone_2.value     = "";
	document.form.Address.value     = "";
	document.form.MailAddr.value    = "";
	document.form.MailAddrCC.value  = "";

    document.form.action="FOfun01a.asp?Status=Add";
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


//-----------------------Check Keypress----------------------------------------
function CheckKeyPress()
{
    if(CheckKeyCode(KEY_CODE_UP)) // 按"上"鍵
    {
        if (document.form.ID.value != "")
        {
            document.form.Status.value = "Modify"
            document.form.action="FOfun01a.asp?Find=Prev";
            document.form.submit();
        }
    }
    else if(CheckKeyCode(KEY_CODE_DOWN)) // 按"下"鍵
    {
        if (document.form.ID.value != "")
        {
            document.form.Status.value = "Modify"
            document.form.action="FOfun01a.asp?Find=Next";
            document.form.submit();
        }
    } 
}