<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>航運系統</title>
</head>
<body>
<%
	'開啟ACCESS資料庫作資料庫的連結讀取
	Set ConnUser = Server.CreateObject("ADODB.Connection")
	ConnUser.open "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" & Server.MapPath("../App_Data/HAS.accdb")

	Dim szID, szPassword
	
	szID = request("ID")
	szPassword = request("Password")

	dim sql, rs
	'sql = "select GroupType, iif(isnull(Owner), '', Owner) AS Owner from Users where ID='" + szID + "' and Password = '" + szPassword + "'"
	sql = "select * from Users where ID='" + szID + "' and Password = '" + szPassword + "'"
	
	set rs = ConnUser.execute (sql)

	if not rs.eof then
		Session.Contents("GroupType@TAIINS2")		= rs("GroupType")
		Session.Contents("Owner@TAIINS2")			= ""

		if Len(CStr(rs("Owner"))) > 0 then
			if CStr(rs("Owner")) <> "-" then
				Session.Contents("Owner@TAIINS2")	= CStr(rs("Owner"))
			end if
		end if
	
		rs.close
		ConnUser.close 
    
		set rs=nothing
		set ConnUser=nothing

		response.Redirect("../main.htm")
	else
		rs.close
		ConnUser.close 
    
		set rs=nothing
		set ConnUser=nothing

		Session.Contents("GroupType@TAIINS2") = -1 
		response.write("<font size=6><br><meta http-equiv=""refresh"" content=""2; url=Login.asp""><center>帳號密碼錯誤</center>")    
	end if
%>
</body>
</html>
