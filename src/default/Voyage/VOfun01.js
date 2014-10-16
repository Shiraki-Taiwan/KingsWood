//-------------------- 初始游標位置 ------------------------------------
function CLocation() {
	if (form.SerialNum)
		form.SerialNum.focus();
}

//-------------------- Submit前欄位驗證 ----------------------------
function checkform()
{
    var bFlag = 0;
    var szErrorString = "";
       
    // 驗證是否必填欄位已填
    //if ( document.form.ID.value == "" )
    //{
    //    bFlag = 1; 
    //    szErrorString = szErrorString  + "【編號】尚未填寫" + "\n";
    //}
    
    if ( (document.form.VesselName.value == "") && (document.form.VesselNo.value == "") &&
         (document.form.CheckInID.value == "")  && (document.form.Year.value == "") &&
         (document.form.Owner.value == "")) 
    {
        bFlag = 1; 
        szErrorString = szErrorString  + "請至少填一項資料" + "\n";
    }
    //else if  (document.form.VesselLine.value == "0")
    //{
    //    bFlag = 1; 
    //    szErrorString = szErrorString  + "【航線】尚未填寫" + "\n";
    //}
   
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
        document.form.action="VOfun01b.asp";
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
        ChangeFocus(10);
    }
}

//-----------------------查詢功能----------------------------------------
function Search()
{
    document.form.action="VOfun01a.asp?Status=Search";
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
        ChangeFocus(11);
    }
}

//-----------------------刪除功能----------------------------------------
function Delete()
{
    if (document.form.Status.value == "Modify") // 只有在Modify的狀況下才能修改或刪除
    {
        if (window.confirm("確定要刪除嗎？"))
    	{
    		document.form.action="VOfun01b.asp?Status=Delete";
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
        ChangeFocus(12);
    }
}


//-----------------------結關功能----------------------------------------
function Close()
{
    document.form.action="VOfun01b.asp?Status=Close";
    document.form.submit();
}

//
function CloseByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        Close();
    }
    else
    {
        ChangeFocus(12);
    }
}

//-----------------------清除功能----------------------------------------
function Reset()
{
    document.form.ID.value   = "";
	document.form.VesselName.value = "";
	document.form.VesselNo.value = "";
	document.form.CheckInID.value = "";
	document.form.VesselLine.value = "";
	document.form.Year.value = "";
	document.form.Month.value = "";
	document.form.Day.value = "";
	document.form.SerialNum.value = "";
	
	document.form.action="VOfun01a.asp?Status=Add";
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


//---------------
function OnSerialNum()
{
    if (CheckKeyCode(KEY_CODE_ENTER)) // Check if press "Enter"
    {
        if (document.form.SerialNum.value != "")
        {
            document.form.action="VOfun01c.asp?SerialNum=" + document.form.SerialNum.value;
	        document.form.submit();   
	    }
	    else
	    {
	    	// 07-Apr2005: 若Serial Number是空白就跳到下一格 (即Enter的正常功能)
	    	ChangeFocus(1);
	    }
	}
}