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
        document.form.action="COfun01b.asp";
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
        ChangeFocusByPressedKey(12,14,14);
    }
}

//-----------------------查詢功能----------------------------------------
function Search()
{
    document.form.action="COfun01a.asp?Status=Search";
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
        ChangeFocusByPressedKey(13,15,15);
    }
}

//-----------------------刪除功能----------------------------------------
function Delete()
{
    if (document.form.Status.value == "Modify") // 只有在Modify的狀況下才能修改或刪除
    {
        if (window.confirm("確定要刪除嗎？"))
    	{
    		document.form.action="COfun01b.asp?Status=Delete";
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
        ChangeFocusByPressedKey(14,16,16);
    }
}
//-----------------------清除功能----------------------------------------
function Reset()
{
    document.form.ID.value   = "";
	document.form.ChineseName.value = "";
	document.form.EnglishName.value = "";
    document.form.FaxNo_1.value = "";
	document.form.FaxNo_2.value = "";
	document.form.Phone_1.value = "";
	document.form.Phone_2.value = "";
	document.form.Contact_1.value = "";
	document.form.Contact_2.value = "";
	document.form.Mobile_1.value = "";
	document.form.Mobile_2.value = "";
	document.form.Address.value = "";
	
	document.form.action="COfun01a.asp?Status=Add";
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
        ChangeFocusByPressedKey(14,0,0);
    }
}