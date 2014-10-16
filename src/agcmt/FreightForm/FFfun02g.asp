<!-- #include file = ../GlobalSet/conn.asp -->
<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@AGCMT") < 0 then
		response.redirect "../Login/Login.asp"
	end if

    nGroupType      = Session.Contents("GroupType@AGCMT")
%>
<html>
<head>
	<title>進倉單資料修改/核定</title>
</head>
<body>
<%
    dim sql, rs
    set rs = nothing
    
    '===========接參數===========
    dim szVesselListID
    
    szVesselListID				= request ("VesselListID")		'航次

    '===========查詢欲修改的資料=========== 
    dim szVesselLine
    
    '查船的資料
    sql							= "select VesselLine from VesselList where ID = '" + szVesselListID + "'"     
    set rs						= conn.execute(sql)
    if not rs.eof then
        szVesselLine			= rs("VesselLine")				'航線
    end if

	response.redirect "../FaxService/FSfun03c.asp?ReportType=1&VesselLine=" & szVesselLine & "&VesselListID=" & szVesselListID
%>
</body>
</html>
