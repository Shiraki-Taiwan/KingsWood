//-------------------- 初始游標位置 ------------------------------------
function CLocation()
{
	form.Y1.focus();   
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
    
    if ((document.form.Y1.value == "") || (document.form.M1.value == "") || (document.form.D1.value == "")) 
    {
        bFlag = 1; 
        szErrorString = szErrorString  + "請填寫開始日期" + "\n";
    }
    
    if ((document.form.Y2.value == "") || (document.form.M2.value == "") || (document.form.D2.value == "")) 
    {
        bFlag = 1; 
        szErrorString = szErrorString  + "請填寫結束日期" + "\n";
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

//-----------------------備分功能----------------------------------------
function Send()
{
	var szStr
	szStr = "系統會在備份好資料庫後, 從現有資料庫中刪除";
	szStr = szStr + document.form.Y1.value +"/" + document.form.M1.value + "/" + document.form.D1.value;
	szStr = szStr + "到"
	szStr = szStr + document.form.Y2.value +"/" + document.form.M2.value + "/" + document.form.D2.value;
	szStr = szStr + "的資料, 您確定要備份嗎？";
	if (window.confirm(szStr))
	{
		if (checkform())
	    {
	        document.form.action="BKfun01b.asp";
	    	document.form.submit();
	    }
	}
}

//
function SendByKeyPress()
{
    if (CheckKeyCode(KEY_CODE_SPACE)) // Check if press "Space"
    {
        Send();
    }
    else
    {
        ChangeFocus(0);
    }
}


//-----------------------檢查輸入的月份是否有誤-----------------------------
function MonthChecking(Object)
{
    if ( Object.value == "")
    {
        return;
    }

    if ( Object.value < 1 || Object.value > 12 )
    {
        alert("您輸入的月份不在正常範圍內!!");
        Object.focus();
    }
}

//-----------------------檢查輸入的日期是否有誤-----------------------------
function DayChecking(Object)
{
    var year, month, day;
    
    if ( Object.name == "D1")  //起始日期
    {
        year = document.form.Y1.value;
        month = document.form.M1.value;
    }
    else    //結束日期
    {
        year = document.form.Y2.value;
        month = document.form.M2.value;
    }

    
    day = Object.value;

    
    if ( (month == "") && (day == ""))
    {
        return;
    }
    
    
    switch(month)
    {   
        case "1":
        case "3":
        case "5":
        case "7":
        case "8":
        case "10":
        case "12":
        {
            if (day > 31)
            {
                alert("您輸入的日期不在正常範圍內!!");
                Object.focus();
            }
            
            break;
        }
             
        case "4":
        case "6":
        case "9":
        case "11":
        {
            if (day > 30)
            {
                alert("您輸入的日期不在正常範圍內!!");
                Object.focus();
            }
            
            break;
        }
        
        case "2":
        {
            // 潤年
            if ( (eval(year) % 4) == 0)
            {
                if (day > 29)
                {
                    alert("您輸入的日期不在正常範圍內!!");
                    Object.focus();
                }
            }
            // 非潤年
            else 
            {
                if (day > 28)
                {
                    alert("您輸入的日期不在正常範圍內!!");
                    Object.focus();
                }
            }
        
            break;
        }
    }
}