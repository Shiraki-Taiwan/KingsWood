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
    if ( document.form.VesselListID.value == "" )
    {
        bFlag = 1; 
        szErrorString = szErrorString  + "【單號】尚未填寫" + "\n";
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


//-----------------------查詢功能----------------------------------------
function Search()
{
    document.form.action="FFfun02a.asp";
    document.form.submit();
}

//
function SearchByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
         Search();
    }
    else if (CheckKeyCode(KEY_CODE_ENTER))
    {
         ChangeFocus(10);
    }
}


//-----------------------清除功能----------------------------------------
function Reset()
{
    document.form.SerialNum.value = "";
    document.form.VesselName.value = "";
    document.form.Owner.value = "";
    document.form.CheckInID.value = "";
    document.form.VesselName.value = "";
    document.form.VesselNo.value = "";
    document.form.Year.value = "";
    document.form.Month.value = "";
    document.form.Day.value = "";
    document.form.VesselLine.value=0;
    
    document.form.action="FFfun02a.asp";
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
        document.form[0].focus();
    }
}


//---------------
function OnSerialNum()
{
    if (CheckKeyCode(KEY_CODE_ENTER)) // Check if press "Enter"
    {
        if (document.form.SerialNum.value != "")
        {
            document.form.action="../Voyage/VOfun01c.asp?SerialNum=" + document.form.SerialNum.value;
	        document.form.submit();   
	    }
	    else
	    {
	    	// 07-Apr2005: 若Serial Number是空白就跳到下一格 (即Enter的正常功能)
	    	ChangeFocus(1);
	    }
	}
}