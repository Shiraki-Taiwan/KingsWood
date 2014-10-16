
<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@ADMIN") < 0 then
		response.redirect "../Login/Login.asp"
	end if

	dim nStatus
	dim szFolderPath, szCurDBPath, szCurDBPathTmp, szBkDBPath

	Set objFileSystem	= CreateObject ("Scripting.FileSystemObject")
	
	nStatus				= request("Status")
	szFolderPath		= Request.ServerVariables ("PATH_TRANSLATED")	
	szFolderPath		= objFileSystem.GetParentFolderName (szFolderPath)
	szFolderPath		= objFileSystem.GetParentFolderName (szFolderPath)
	
	szBkDBPath			= szFolderPath & "\Database\" & request("DBName") & "\INSDB.mdb"
	'szBkDBPath			= szFolderPath & "\App_Data\" & request("DBName") & "\INSDB.mdb"
	'szFolderPath		= objFileSystem.GetParentFolderName (szFolderPath)

	'szCurDBPath		= szFolderPath & "\Database\INSDB.mdb"
	'szCurDBPathTmp		= szFolderPath & "\Database\INSDB_BK.mdb"
	szCurDBPath			= szFolderPath & "\App_Data\INSDB.mdb"
	szCurDBPathTmp		= szFolderPath & "\App_Data\INSDB_BK.mdb"

	'if NOT objFileSystem.FileExists (szCurDBPath) then
	'	Response.Write(szCurDBPath)
	'	Response.Write("<br />")
	'end if
	'
	'if NOT objFileSystem.FileExists (szCurDBPathTmp) then
	'	Response.Write(szCurDBPathTmp)
	'	Response.Write("<br />")
	'end if
	'Response.Write(nStatus)

	if nStatus = 1 then	'匯入備份資料庫
		'備份現有的DB
		if NOT objFileSystem.FileExists (szCurDBPathTmp) then
			objFileSystem.MoveFile szCurDBPath, szCurDBPathTmp
			'objFileSystem.CopyFile szCurDBPath, szCurDBPathTmp, FALSE
		end if
		
		'以備份db取代現有db
		objFileSystem.CopyFile szBkDBPath, szCurDBPath, TRUE
	elseif nStatus = 2 then	'回復成原始資料庫
		objFileSystem.CopyFile szCurDBPathTmp, szCurDBPath, TRUE
		objFileSystem.DeleteFile szCurDBPathTmp
	end if

	set objFileSystem	= nothing

    Response.Write("<font size=6><br><meta http-equiv=""refresh"" content=""2;url=BKfun02a.asp""><center>匯入成功!</center>")    
%>
