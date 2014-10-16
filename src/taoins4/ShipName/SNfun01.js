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
        szErrorString = szErrorString  + "【船碼】尚未填寫" + "\n";
    }
    
    if ( document.form.Name.value == "" )
    {
        bFlag = 1; 
        szErrorString = szErrorString  + "【船名】尚未填寫" + "\n";
    }
    
    if ( document.form.Owner.value == "" )
    {
        bFlag = 1; 
        szErrorString = szErrorString  + "【船公司】尚未填寫" + "\n";
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
        document.form.action="SNfun01b.asp";
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
        ChangeFocus(4);
    }
}
//-----------------------查詢功能----------------------------------------
function Search()
{
    document.form.action="SNfun01a.asp?Status=Search";
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
        ChangeFocus(5);
    }
}
//-----------------------刪除功能----------------------------------------
function Delete()
{
    if (document.form.Status.value == "Modify") // 只有在Modify的狀況下才能修改或刪除
    {
        if (window.confirm("確定要刪除嗎？"))
    	{
    		document.form.action="SNfun01b.asp?Status=Delete";
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
        ChangeFocus(6);
    }
}
//-----------------------清除功能----------------------------------------
function Reset()
{
    document.form.ID.value   = "";
	document.form.Name.value = "";
	document.form.Owner.value = "";
	
    document.form.action="SNfun01a.asp?Status=Add";
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