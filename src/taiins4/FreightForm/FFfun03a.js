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
    if ( document.form.VesselLin.value == "0" )
    {
        bFlag = 1; 
        szErrorString = szErrorString  + "【航線】尚未填寫" + "\n";
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
    document.form.action="FFfun03b.asp?Status=Add";
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
        document.form[0].focus();
    }
}

