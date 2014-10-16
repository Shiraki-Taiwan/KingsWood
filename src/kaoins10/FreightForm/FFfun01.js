//-------------------- 初始游標位置 ------------------------------------
function CLocation()
{
	document.form[2].focus();
}

//-----------------------新增功能----------------------------------------
function OnSave()
{
    document.form.action="FFfun01b.asp";
    document.form.submit();
}

//
function AddByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        OnSave();
    }
    else
    {
        // 15-Mar2005: 按Enter不要改變focus
        //document.form.BackBtn.focus();
    }
}

//-----------------------刪除功能----------------------------------------
function Delete()
{
    if (document.form.Status.value == "Modify") // 只有在Modify的狀況下才能修改或刪除
    {
        if (window.confirm("確定要刪除嗎？"))
    	{
    		document.form.action="FFfun01b.asp?Status=Delete";
    	    document.form.submit();
    	}
    }
}

//-----------------------清除功能----------------------------------------
function Reset()
{
	document.form.action="FFfun01a.asp?Status=Add";
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
        // 15-Mar2005: 按Enter不要改變focus
        //ChangeFocus(2);
    }
}

//-----------------------清除部分功能----------------------------------------
function ResetPartByIndex(a_nIndex)
{
    document.form[a_nIndex].checked = false;
    document.form[a_nIndex+1].value = "";
    document.form[a_nIndex+2].value = "";
    document.form[a_nIndex+3].value = "";
    document.form[a_nIndex+4].value = "";
    document.form[a_nIndex+5][0].selected = true;
    document.form[a_nIndex+6].value = "";
    document.form[a_nIndex+7][0].selected = true;
    document.form[a_nIndex+8].value = "";
    document.form[a_nIndex+9].value = "";
    document.form[a_nIndex+10].value = "";
    document.form[a_nIndex+11].value = "";
    document.form[a_nIndex+12].value = "";
    document.form[a_nIndex+13].value = "";
    
    var ColorWhite;
    ColorWhite = "#ffffff";
    var i;
    
    for (i=1; i<14; i++)
    {
        document.form[a_nIndex+i].style.backgroundColor = ColorWhite;
    }    
}


function ResetPart()
{
    var index, nLineCount, i;
    nLineCount = 15;
    index = 1;
    
    for (i=0; i<=50; i++)
    {
        if (document.form(index).checked == true)
        {
            ResetPartByIndex(index);
        }
        
        index = index + nLineCount;
    }
}


function ResetPartByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        ResetPart();
    }
    else
    {
        // 15-Mar2005: 按Enter不要改變focus
        //document.form.ResetBtn.focus();
    }
}

//-----------------------回航次資料設定功能----------------------------------------
function Back()
{
    document.form.VesselID.value = "";
    document.form.action="../Voyage/VOfun01a.asp?szStatus=Add";
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
        // 15-Mar2005: 按Enter不要改變focus
        //document.form.ResetPartBtn.focus();
    }
}


// -----------------------for最左邊的欄位-----------------------
function OnDirectionKey_LField(a_iCurElement, a_iLineCounter)
{
    var nTotalLineCounter;
	// 必須隨著實際行數而定
    nTotalLineCounter = 50;
    
	if (CheckKeyCode(KEY_CODE_LEFT))  // 左
	{
		if (a_iCurElement > 2)
		{
		    document.form[a_iCurElement-3].focus();
		}
	}
	else if (CheckKeyCode(KEY_CODE_RIGHT))  // 右
	{
		document.form[a_iCurElement+1].focus();
	}
	else if (CheckKeyCode(KEY_CODE_UP))  // 上
	{
		if (a_iLineCounter > 0)
		{
			document.form[a_iCurElement-15].focus();
		}
	}
	else if (CheckKeyCode(KEY_CODE_DOWN))  // 下
	{
		if (a_iLineCounter < (nTotalLineCounter-1))
		{
			document.form[a_iCurElement+15].focus();
		}
	}
}

// -----------------------for最右邊的欄位-----------------------
function OnDirectionKey_RField(a_iCurElement, a_iLineCounter)
{
	var nTotalLineCounter;
	// 必須隨著實際行數而定
	nTotalLineCounter = 50;
	
	if (CheckKeyCode(KEY_CODE_LEFT))  // 左
	{
		document.form[a_iCurElement-1].focus();
	}
	else if (CheckKeyCode(KEY_CODE_RIGHT))  // 右
	{
		if (a_iCurElement < (nTotalLineCounter*15-1))
		{
		    document.form[a_iCurElement+3].focus();
		}
	}
	else if (CheckKeyCode(KEY_CODE_UP))  // 上
	{
		if (a_iLineCounter > 0)
		{
			document.form[a_iCurElement-15].focus();
		}
	}
	else if (CheckKeyCode(KEY_CODE_DOWN))  // 下
	{
		if (a_iLineCounter < (nTotalLineCounter-1))
		{
			document.form[a_iCurElement+15].focus();
		}
	}
}

// -----------------------for中間的欄位-----------------------
function OnDirectionKey_MField(a_iCurElement, a_iLineCounter)
{
	var nTotalLineCounter;
	// 必須隨著實際行數而定
	nTotalLineCounter = 50;
	
	if (CheckKeyCode(KEY_CODE_LEFT))  // 左
	{
		document.form[a_iCurElement-1].focus();
	}
	else if (CheckKeyCode(KEY_CODE_RIGHT))  // 右
	{
		document.form[a_iCurElement+1].focus();
	}
	else if (CheckKeyCode(KEY_CODE_UP))  // 上
	{
		if (a_iLineCounter > 0)
		{
			document.form[a_iCurElement-15].focus();
		}
	}
	else if (CheckKeyCode(KEY_CODE_DOWN))  // 下
	{
		if (a_iLineCounter < (nTotalLineCounter-1))
		{
			document.form[a_iCurElement+15].focus();
		}
	}
}


function OnDirectionKey_MField_OnlyRL(a_iCurElement, a_iLineCounter)
{
	var nTotalLineCounter, nCurSelected;
	// 必須隨著實際行數而定
	nTotalLineCounter = 50;
	
	if (CheckKeyCode(KEY_CODE_LEFT))  // 左
	{
		document.form[a_iCurElement-1].focus();
	}
	else if (CheckKeyCode(KEY_CODE_RIGHT))  // 右
	{
		document.form[a_iCurElement+1].focus();
	}
}


//-----------------------進入欄位"單號"時----------------------------------------
function OnIDFocus(a_nCurIndex)
{    
    if (a_nCurIndex>15)    //第二行之後才執行
    {
        if (document.form[a_nCurIndex-1].value == "" && document.form[a_nCurIndex].value == "" )
        {
            document.form[a_nCurIndex-1].value = document.form[a_nCurIndex-16].value;
            document.form[a_nCurIndex].value = document.form[a_nCurIndex-15].value;
        }        
    }
}



//-----------------------偵測是否要switch到查詢倉單----------------------------------------
function SwitchToSearchFForm()
{
    if (CheckKeyCode(KEY_CODE_F8))
    {
        // 04-Apr2005: 先儲存再跳至倉單查詢
        document.form.action="FFfun01b.asp?GotoFFSearch=1";          
        //document.form.action="FFfun02b.asp?Status=FirstLoad&VesselListID=" + document.form.VesselID.value;
        document.form.submit();
    }
}