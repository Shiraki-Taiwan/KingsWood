<!-- #include file = ../GlobalSet/conn.asp -->
<!-- #include file = ../GlobalSet/ShareFun.asp -->
<!--航次資料刪除-->

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@TAIINS2") < 0 then
		response.redirect "../Login/Login.asp"
	end if


    '序號
    dim nSerialNum, nSerialNumTmp
    nSerialNum = 1
    nSerialNumTmp = CLng(request("SerialNum"))
    
    dim szVesselListID, szVesselLine
    
    GetVesselListIDBySerialNum(nSerialNumTmp)
    
    if szVesselListID <> "" then
    	response.write("<meta http-equiv=""refresh"" content=""0; url=VOfun04a.asp?Status=Modify&ID=" & szVesselListID & """>")
    else
        response.write("<meta http-equiv=""refresh"" content=""1; url=VOfun04a.asp""><font size=6><center>沒有此編號</center></font>")
    end if
    
    
%>