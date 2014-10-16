//-------------------- 初始游標位置 ------------------------------------
function CLocation()
{
	document.form[0].focus(); 
}

//-------------------- Submit前欄位驗證 ----------------------------
function checkform()
{
    var bFlag = 0;
    var szErrorString = "";
       
    // 驗證是否必填欄位已填
    if ( document.form.FormID.value == "" )
    {
        bFlag = 1; 
        szErrorString = szErrorString  + "【單號】尚未填寫" + "\n";
    }
    
    if ( document.form.OwnerID.value == "0" )
    {
        bFlag = 1; 
        szErrorString = szErrorString  + "【攬貨商】尚未填寫" + "\n";
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
        document.form.action="FFfun03c.asp";
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

//-----------------------查詢功能----------------------------------------
function Search()
{
    document.form.action="FFfun03b.asp?Status=Search&VesselLine=" + document.form.VesselLine.value;
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

//-----------------------刪除功能----------------------------------------
function Delete()
{
    //if (document.form.Status.value == "Modify") // 只有在Modify的狀況下才能修改或刪除
    if (document.form.FormID.value != "")
    {
        if (window.confirm("確定要刪除嗎？"))
    	{
    		document.form.action="FFfun03c.asp?Status=Delete";
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
        ChangeFocus(5);
    }
}
//-----------------------清除功能----------------------------------------
function Reset()
{
    document.form.FormID.value   = "";
    document.form.OwnerID.value   = "";
    document.form.action="FFfun03b.asp?Status=Add";
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
        ChangeFocus(6);
    }
}

//-----------------------回上一頁功能----------------------------------------
function Back()
{
    document.form.action="FFfun03a.asp";
	document.form.submit();
}


//
function BackByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        Back();
    }
    else
    {
        document.form[0].focus();
    }
}