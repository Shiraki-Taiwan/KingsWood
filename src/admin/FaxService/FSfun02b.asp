<!--傳真結果查詢-->

<html>
<head>
<title>攬貨報告書列印/傳真</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<LINK href="../GlobalSet/sg.css" rel=stylesheet type=text/css>
</head>

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@ADMIN") < 0 then
		response.redirect "../Login/Login.asp"
	end if



    '===============接收參數===============
    dim szStatus
    szStatus = request("Status")    '編輯Status
    
        
    dim szLogFile(20), nFileCounter, szLogFileState(20)
    nFileCounter = request("FileCounter")    '檔案個數
    
    if nFileCounter = "" then
        nFileCounter = 0
    end if
    
    dim j
    for j = 0 to  nFileCounter-1
        szLogFile(j) = request ("LogFile_" + CStr(j))        
        szLogFileState(j) = request ("LogFileState_" + CStr(j))
    next
    
    
    Dim objFileSystem
    
    '===============刪除檔案=================
    if szStatus = "Delete" then
        
        Set objFileSystem = CreateObject ("Scripting.FileSystemObject")
        for j = 0 to  nFileCounter-1
            if szLogFileState(j) = 1 then
                if objFileSystem.FileExists(szLogFile(j)) then
                    objFileSystem.DeleteFile szLogFile(j), True
                end if
            end if
        next
    
        response.redirect "FSfun02a.asp"
    end if
%>

<form name="form" method="post" action="FSfun02a.asp">
<table width="100%" height=""100%"">

    <tr> 
        <td >
            <table border=0 cellspacing=1 width="100%" height=""100%"" align="center">
                           
                <tr>
                    <td align="center"><font size=5>傳真紀錄</font></td>
                </tr>
                <tr><td><hr></td></tr>
           </table>
       </td>
    <tr>
    <tr>
        <td>
                 
   
<%    
    '===============顯示檔案=================
    dim szStr, szDate, szTime, szTimeTmp, nPos, szVesselName, szVesselList, szVesselDate, szOwnerID, szOwnerName
    if szStatus = "Show" then
        Set objFileSystem = CreateObject ("Scripting.FileSystemObject")
                
        for j = 0 to  nFileCounter-1
            if szLogFileState(j) = 1 then
                Set objReadTextFile = objFileSystem.OpenTextFile(szLogFile(j))
                
                nPos = InStrRev (szLogFile(j), "\")
                szTimeTmp = Mid (szLogFile(j), nPos+1)
                
                szDate = Mid (szTimeTmp, 1, 2) & "/" & Mid (szTimeTmp, 3, 2) & "/" & Mid (szTimeTmp, 5, 2)
                szTime = Mid (szTimeTmp, 8, 2) & ":" & Mid (szTimeTmp, 10, 2)
                
                response.write "<table border=0 cellspacing=0 width=""100%"" height=""100%"" align=""center"">"
                response.write "<tr align=""center"">"
                response.write "<td align=""right""><span style=""letter-spacing: 8pt"">傳真時間</span>：</td>"
                response.write "<td align=""left"" colspan=5>" & szDate & "　" & szTime & "</td>"
                response.write "</tr>"
                
                
                while not objReadTextFile.AtEndOfLine
                    szStr = objReadTextFile.ReadLine()
                   
                    response.write "<tr align=""center"">"
                    
                                                           
                    if szStr = "[Failed]" then
                        response.write "<td align=""right"" width=""30%"">傳真失敗資料：</td>"
                        response.write "<td align=""left"" width=""10%"">船名</td>"
                        response.write "<td align=""left"" width=""10%"">航次</td>"
                        response.write "<td align=""left"" width=""10%"">結關日</td>"
                        response.write "<td align=""left"" width=""40%"">攬貨商</td>"
                           
                    elseif szStr = "[Success]" then
                        response.write "<td align=""right"" width=""30%"">傳真成功資料：</td>"
                        response.write "<td align=""left"" width=""10%"">船名</td>"
                        response.write "<td align=""left"" width=""10%"">航次</td>"
                        response.write "<td align=""left"" width=""10%"">結關日</td>"
                        response.write "<td align=""left"" width=""40%"">攬貨商</td>"                        
                    else
                        nPos = InStrRev (szStr, "=")
                        szStr = Mid (szStr, 1, nPos-1)
                        
                        nPos = InStrRev (szStr, "_")                        
                        szVesselDate = Mid (szStr, nPos+1)
                        szStr = Mid (szStr, 1, nPos-1)
                        
                        nPos = InStrRev (szStr, "_")                        
                        szVesselList = Mid (szStr, nPos+1)
                        szStr = Mid (szStr, 1, nPos-1)
                       
                        nPos = InStrRev (szStr, "_")                        
                        szVesselName = Mid (szStr, nPos+1)
                        szStr = Mid (szStr, 1, nPos-1)
                        
                        nPos = InStrRev (szStr, "_")                        
                        szOwnerName = Mid (szStr, nPos+1)
                        
                        szOwnerID = Mid (szStr, 1, nPos-1)
                        
                        response.write "<td align=""right""></td>"
                        response.write "<td align=""left"">" & szVesselName & "</td>"
                        response.write "<td align=""left"">" & szVesselList & "</td>"
                        response.write "<td align=""left"">" & szVesselDate & "</td>" 
                        response.write "<td align=""left"">" & szOwnerID & "　" & szOwnerName & "</td>"
                                               
                    end if
                    
                    response.write "</tr>"           
                wend
                
                objReadTextFile.Close
                
                response.write "</table>"
                response.write "<hr>"
            end if
        
        next
    end if
    
    
%>

        </td>
   </tr>
   <tr>
        <td align="center">
            <input type="submit" value="回上一頁" name="Back">
        </td>
   </tr>
</table>

 
</form>
</body>
</html>