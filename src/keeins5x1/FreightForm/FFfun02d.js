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
    if ( document.form.VesselID.value == "" )
    {
        bFlag = 1; 
        szErrorString = szErrorString  + "請選擇航次" + "\n";
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
function OnSearch()
{
    document.form.action="FFfun02d.asp?OriginalVesselListID=" + document.form.OriginalVesselListID.value;
    document.form.submit();
}

//
function SearchByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
         OnSearch();
    }
    else if (CheckKeyCode(KEY_CODE_ENTER))
    {
         document.form.Reset.focus();
    }
}


//-----------------------清除功能----------------------------------------
function OnReset()
{
    document.form.VesselID.value = "";
    document.form.Owner.value = "";
    document.form.CheckInID.value = "";
    document.form.VesselName.value = "";
    document.form.VesselNo.value = "";
    document.form.Year.value = "";
    document.form.Month.value = "";
    document.form.Day.value = "";
    document.form.VesselLine.value = "";
    
    document.form.action="FFfun02d.asp?OriginalVesselListID=" + document.form.OriginalVesselListID.value;
    document.form.submit();
}

//
function ResetByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        OnReset();
    }
    else
    {
        document.form[0].focus();
    }
}


//-----------------------移動功能----------------------------------------
function OnMove()
{
    if (document.form.VesselID.value == "" )
    {
        alert ("請選擇航次");
    }
    else
    {
        // 21-Nov2004:航次相同, 則不用移動
        if (document.form.OriginalVesselListID.value == document.form.VesselID.value)
        {
            alert ("航次相同,不需移動!");
            return;
        }
        
        if (window.confirm("確定要移動嗎？"))
    	{
    	    var i, nDataCounter, szStr1, szStr2, iIndex, nCounter;
    	    nDataCounter = document.form.DataCounter.value;
    	    
    	    iIndex = 6;
    	    nCounter = 0;
    	    szStr2 = "";
    	    
    	    var szSN
    	    
    	    for (i=0; i<nDataCounter; i++)
    	    {
    	        szSN = eval("document.form.SN_" + i + ".value")
    	         
    	        szStr2 = szStr2 + "&SN_" + String(nCounter) + "=" + szSN;
    	        nCounter = nCounter + 1
    	    }
    	    
    	    szStr = "FFfun02e.asp?FormID=" + document.form.FormID.value;
            szStr = szStr + "&OriginalVesselListID=" + document.form.OriginalVesselListID.value;
            szStr = szStr + "&VesselID=" + document.form.VesselID.value;
            szStr = szStr + "&DataCounter=" + String(nCounter) + szStr2;
    	    
    	    document.form.action=szStr;
        	document.form.submit();
    	}
    }
}

//
function MoveByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        OnMove();
    }
    else
    {
        document.form.Search.focus();
    }
}