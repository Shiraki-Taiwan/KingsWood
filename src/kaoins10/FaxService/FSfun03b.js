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
    document.form.action="FSfun03b.asp?ReportType=" + document.form.ReportType.value;
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
         ChangeFocusByPressedKey(7,10,10);
    }
}


//-----------------------清除功能----------------------------------------
function Reset()
{
    document.form.SerialNum.value = "";
    document.form.VesselName.value = "";
    document.form.VesselNo.value = "";
    document.form.Year.value = "";
    document.form.Month.value = "";
    document.form.Day.value = "";
    
    document.form.action="FSfun03b.asp";
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
        ChangeFocusByPressedKey(8,11,11);
    }
}


//-----------------------回上一頁功能----------------------------------------
function Back()
{
    document.form.action="FSfun03a.asp";
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
        ChangeFocusByPressedKey(9,0,0);
    }
}


//---------------
function OnSerialNum()
{
    if (CheckKeyCode(KEY_CODE_ENTER)) // Check if press "Enter"
    {
        if (document.form.SerialNum.value != "")
        {
            document.form.action="FSfun03d.asp?SerialNum=" + document.form.SerialNum.value + "&ReportType=" + document.form.ReportType.value;
	        document.form.submit();   
	    }
	    else
	    {
	    	// 07-Apr2005: 若Serial Number是空白就跳到下一格 (即Enter的正常功能)
	    	ChangeFocus(1);
	    }
	}
}