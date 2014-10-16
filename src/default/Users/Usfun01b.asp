<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>航運系統</title>
</head>
<body>
<!--使用者設定-->
<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@DEFAULT") < 0 then
		response.redirect "../Login/Login.asp"
	end if

	'開啟ACCESS資料庫作資料庫的連結讀取
	Set ConnUser = Server.CreateObject("ADODB.Connection")
	ConnUser.open "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" & Server.MapPath("../App_Data/HAS.accdb")

    '===============接收參數===============
    dim szStatus
    szStatus = request("Status")    '編輯Status
    
    dim szID, szPassword, nGroupType, szOwner
    
    if szStatus = "Modify" then     '帳號
        szID = request("IDToModify")
    else
        szID = request("ID")        
    end if
    
    szPassword		= request ("Password")      '密碼
    nGroupType		= request ("GroupType")		'群組
	szOwner			= request ("Owner")			'所屬廠商

    '===============寫入資料庫===============  
    dim sql, rs 
    set rs = nothing 
  
    '=========新增=========
    if szStatus = "Add" then 
        '檢查編號是否重複  
        sql = "select ID from Users where ID='" + szID + "'"
        set rs = ConnUser.execute(sql)    
        
        if rs.eof then
			if Len(szOwner + "") > 0 then
				sql = "insert into Users(ID, Password, GroupType, Owner) values('" + szID + "','" + szPassword + "'," + CStr(nGroupType) + ", '" + szOwner + "')"
			else
				sql = "insert into Users(ID, Password, GroupType, Owner) values('" + szID + "','" + szPassword + "'," + CStr(nGroupType) + ", '-')"
			end if
            ConnUser.execute(sql)
            
            rs.close
            ConnUser.close 
            set rs=nothing
            set ConnUser=nothing
            response.redirect "Usfun01a.asp"
        else
            '編號重複, 回上一頁
            rs.close
            ConnUser.close 
            set rs=nothing
            set ConnUser=nothing
            response.write("<font size=6><br><meta http-equiv=""refresh"" content=""2; url=Usfun01a.asp""><center>帳號重複</center>")    
        end if
        
    '=========修改=========
    elseif szStatus = "Modify" then
        sql = "update Users set Password = '"+ szPassword +"', Owner = '"+ szOwner +"', GroupType = " + CStr(nGroupType) + " where ID = '"+ szID +"'"
      	ConnUser.execute(sql)
      	
      	ConnUser.close 
        set rs=nothing
        set ConnUser=nothing
        response.redirect "Usfun01a.asp"
            
    '=========刪除=========
    elseif szStatus = "Delete" then
        sql = "delete from Users where ID ='" + szID + "'"
      	ConnUser.execute(sql)
      	
      	ConnUser.close 
        set rs=nothing
        set ConnUser=nothing
        response.redirect "Usfun01a.asp"
        
    end if
%>
</body>
</html>