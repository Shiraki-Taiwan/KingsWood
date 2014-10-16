<!-- #include file = ../GlobalSet/conn.asp -->
<!-- #include file = ../GlobalSet/ShareFun.asp -->
<!--攬貨報告書列印/傳真-->

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@TAIINS4") < 0 then
		response.redirect "../Login/Login.asp"
	end if


    '
    dim szReportType
    szReportType = request("ReportType")
    
    '序號
    dim nSerialNum, nSerialNumTmp
    nSerialNum = 1
    nSerialNumTmp = CLng(request("SerialNum"))
    
    dim szVesselListID, szVesselLine
    
    GetVesselListIDBySerialNum(nSerialNumTmp)
    
    if szVesselListID <> "" then
        response.write("<meta http-equiv=""refresh"" content=""0; url=FSfun03c.asp?ReportType=" & szReportType & "&VesselListID=" & szVesselListID & "&VesselLine=" & szVesselLine & """>")
    else
        response.write("<meta http-equiv=""refresh"" content=""1; url=FSfun03b.asp?ReportType=" & szReportType & """><font size=6><center>沒有此編號</center></font>")
    end if
    
    
%>