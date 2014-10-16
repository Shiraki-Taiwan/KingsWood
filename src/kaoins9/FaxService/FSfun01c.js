//-------------------- 初始游標位置 ------------------------------------
function CLocation()
{
    if (document.form.SelectType.value == 1)
    {
        document.form[1].focus();
    }
    else
    {
	    document.form[0].focus();
	}
}

//-----------------------查詢列印功能----------------------------------------
function Print()
{
    document.form.action="FSfun01e.asp?Status=Print";
    document.form.submit();
}

//
function PrintByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
         Search();
    }
    else if (CheckKeyCode(KEY_CODE_ENTER))
    {
         document.form.Re_Search.focus();
    }
}

//-----------------------查詢傳真功能----------------------------------------
function Fax()
{
    document.form.action="FSfun01d.asp?Status=Fax";
    document.form.submit();
}

//
function FaxByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
         Fax();
    }
    else if (CheckKeyCode(KEY_CODE_ENTER))
    {
         document.form.SendMail.focus();
    }
}

//-----------------------查詢Mail功能----------------------------------------
function Mail()
{
    var i;
/*   
    // 30-Mar2005: 不強迫填入mail address,若沒有, 就不send 
    for (i=1; i <= document.form.FoundCount.value; i++)
    {
        var objName = eval("document.form.MailAddr_" + String(i));
        
        if (objName.value == "")
        {
            alert ("請填妥e-Mail Address")
            return;
        }
    }
*/
    // 05-May2005: 若是依單號查詢，則必須填入 mail address
    if (document.form.SelectType.value == 1)
    {
        if (document.form.MailAddress.value == "")
        {
            alert ("請填妥e-Mail Address")
            return;
        }
    }
    
    document.form.action="FSfun01d.asp?Status=Mail";
    document.form.submit();
}

//
function MailByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        //Fax();
        Mail
    }
    else if (CheckKeyCode(KEY_CODE_ENTER))
    {
         document.form.SendPrint.focus();
    }
}
//-----------------------重新查詢功能----------------------------------------
function OnReSearch()
{
    document.form.action="FSfun01b.asp";
	document.form.submit();
}

//
function ReSearchByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        OnReSearch();
    }
    else if (CheckKeyCode(KEY_CODE_ENTER))
    {
   		document.form[0].focus();
    }
}

