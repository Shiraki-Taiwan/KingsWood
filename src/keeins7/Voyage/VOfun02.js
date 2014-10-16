//-------------------- 初始游標位置 ------------------------------------
function CLocation()
{
	form.VesselName.focus();   
}


//-----------------------查詢功能----------------------------------------
function Search()
{
    document.form.action="VOfun02b.asp";
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
        ChangeFocus(8);
    }
}

//-----------------------清除功能----------------------------------------
function Reset()
{
    document.form.ID.value   = "";
	document.form.VesselName.value = "";
	document.form.VesselNo.value = "";
	document.form.Owner.value = "";
	document.form.CheckInID.value = "";
	document.form.Year.value = "";
	document.form.Month.value = "";
	document.form.Day.value = "";
	
	
	document.form.action="VOfun02a.asp";
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


