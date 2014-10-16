<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>航運系統</title>
</head>
<body>
<!--顯示No data-->
<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@KEEINS4") < 0 then
		response.redirect "../Login/Login.asp"
	end if

    
    szVesselListID = request("VesselListID")
    szPrevFoundID = request("PrevFoundID")
    
    response.write "<font size=6><br>"
    response.write("<meta http-equiv=""refresh"" content=""1; url=FFfun02b.asp?Status=UsePrevFoundID&PrevFoundID=" + szPrevFoundID + "&VesselListID=" + szVesselListID + """><center>沒有此S/O號碼</center>")
    
%>
</body>
</html>
