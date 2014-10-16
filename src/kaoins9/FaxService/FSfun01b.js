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
    if ( document.form.ID.value == "" )
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
    if (document.form.SelectType[1].checked == true)
    {
        if (document.form.FreightOwnerID.value == "")
        {
            alert ("請輸入攬貨商號碼")
            return;
        }
        
        //if (document.form.Container.value == "")
        //{
        //    alert ("請輸入貨櫃場")
        //    return;
        //}            
    }
    
    document.form.action="FSfun01c.asp";
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
         ChangeFocusByPressedKey(5,7,7);
    }
}


//-----------------------重新查詢功能----------------------------------------
function OnReSearch()
{
    document.form.VesselListID.value = "";
    document.form.VesselLine.value = "";
    document.form.action="FSfun01a.asp";
	document.form.submit();
}

//
function ReSearchByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        OnReSearch();
    }
    else
    {
        ChangeFocusByPressedKey(8,0,0);
    }
}

//-----------------------For 攬貨商ID's keydown----------------------------------------
function OnFreightOwnerID()
{
    if (CheckKeyCode(KEY_CODE_ENTER) || CheckKeyCode(KEY_CODE_LEFT) || CheckKeyCode(KEY_CODE_RIGHT))
    {
         if (document.form.ReportType.disabled)
         {
            ChangeFocusByPressedKey(2,6,6);
         }
         else
         {
            ChangeFocusByPressedKey(2,5,5);
         }         
    }
    else
    {
        document.form.SelectType[0].checked = false;
        document.form.SelectType[1].checked = true;
        
        document.form.ReportType[0].selected = true;
    }
}


//-----------------------For 單號's keydown----------------------------------------
function OnFormID()
{    
    if (CheckKeyCode(KEY_CODE_ENTER) || CheckKeyCode(KEY_CODE_LEFT) || CheckKeyCode(KEY_CODE_RIGHT))
    {
         ChangeFocusByPressedKey(0,4,4);
    }
    else
    {
        document.form.SelectType[0].checked = true;
        document.form.SelectType[1].checked = false;
        
        
    }
}


//-----------------------加入一ReportType item----------------------------------------
function AddReportTypeItem()
{
    var oNewNode = document.createElement("option");
    oNewNode.ID=2;
    oNewNode.innerText="詳細尺寸資料";
    document.form.ReportType.appendChild(oNewNode);
}

//-----------------------刪除一ReportType item----------------------------------------
function RemoveReportTypeItem()
{
    document.form.ReportType.removeChild(document.form.ReportType.children(2));

}
//-----------------------For 單號's onfocus----------------------------------------
function OnSelectType0()
{
    document.form.SelectType[0].checked = true;
    document.form.SelectType[1].checked = false;
}

//-----------------------For 單號's onfocus----------------------------------------
function OnSelectType1()
{
    document.form.SelectType[0].checked = false;
    document.form.SelectType[1].checked = true;
    
    document.form.ReportType[0].selected = true;
}

