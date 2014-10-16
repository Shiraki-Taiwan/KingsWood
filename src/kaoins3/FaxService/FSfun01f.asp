<!-- #include file = ../GlobalSet/conn.asp -->
<!-- #include file = ../GlobalSet/ShareFun.asp -->
<!--攬貨報告書列印/傳真-->

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@KAOINS3") < 0 then
		response.redirect "../Login/Login.asp"
	end if


    '序號
    dim nSerialNum, nSerialNumTmp
    nSerialNum = 1
    nSerialNumTmp = CLng(request("SerialNum"))
    dim szVesselListID, szVesselLine
    
    GetVesselListIDBySerialNum(nSerialNumTmp)
    
    if szVesselListID <> "" then
        response.write("<meta http-equiv=""refresh"" content=""0; url=FSfun01b.asp?Status=Modify&VesselListID=" & szVesselListID & "&SerialNum=" & nSerialNumTmp & """>")
    else
        response.write("<meta http-equiv=""refresh"" content=""1; url=FSfun01a.asp""><font size=6><center>沒有此編號</center></font>")
    end if
    
    
%>