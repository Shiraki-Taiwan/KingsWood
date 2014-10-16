<!-- #include file = ../GlobalSet/conn.asp -->
<!-- #include file = ../GlobalSet/ShareFun.asp -->
<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@KAOINS3") < 0 then
		response.redirect "../Login/Login.asp"
	end if

	dim sql, rs
	
	sql = "select * from VesselList"
	Set rs = Conn.Execute(sql)
	if rs.eof then
		rs.close
		response.write("<font size=6><br><meta http-equiv=""refresh"" content=""2; url=BKfun01a.asp""><center>沒有資料需要備份!</center>")    
	  	response.end
	end if
		
	'================step1:先copy完整資料庫至指定folder=============
    dim szDstFilePath

	Set objFileSystem	= CreateObject ("Scripting.FileSystemObject")
    szDstFilePath		= Request.ServerVariables ("PATH_TRANSLATED")    
    szDstFilePath		= objFileSystem.GetParentFolderName (szDstFilePath)
	szDstFilePath		= objFileSystem.GetParentFolderName (szDstFilePath)
    
    'szDstFilePath		= szDstFilePath + "\Database\"
	szDstFilePath		= szDstFilePath + "\App_Data\"
    
    '以今天的日期時間開新資料夾
    dim nYear, nMonth, nDay, nHour, nMin, nSec
    nYear				= Year (now())
    nMonth				= FormatString(Month (now()), 2)
    nDay				= FormatString(Day (now()), 2)
    nHour				= FormatString(Hour (now()), 2)
    nMin				= FormatString(Minute (now()), 2)
    nSec				= FormatString(Second (now()), 2)
    
    szDstFilePath		= szDstFilePath & nYear & nMonth & nDay & "_" & nHour & nMin & nSec
    
    if not (objFileSystem.FolderExists(szDstFilePath)) then
        objFileSystem.CreateFolder(szDstFilePath)
    end if
    
    dim szTxtFilePath
    szTxtFilePath = szDstFilePath & "\Backup.txt"
    
    szDstFilePath = szDstFilePath + "\INSDB.mdb"
	
	'找Source File Path
	dim szSrcFilePath, szSrcFilePathTmp
	szSrcFilePath = Request.ServerVariables ("PATH_TRANSLATED")
	
	szSrcFilePath = objFileSystem.GetParentFolderName (szSrcFilePath)	
	szSrcFilePath = objFileSystem.GetParentFolderName (szSrcFilePath)	
	
	
	'szSrcFilePathTmp = szSrcFilePath & "\Database\"	'壓縮資料庫時會用到
	szSrcFilePathTmp = szSrcFilePath & "\App_Data\"	'壓縮資料庫時會用到
	
	'szSrcFilePath = szSrcFilePath & "\Database\INSDB.mdb"
	szSrcFilePath = szSrcFilePath & "\App_Data\INSDB.mdb"
	
	if not objFileSystem.FileExists(szSrcFilePath) then
		response.write("<font size=6><br><meta http-equiv=""refresh"" content=""2; url=BKfun01a.asp""><center>原始資料庫不存在!</center>")    
	  	response.end
	end if
	
	objFileSystem.CopyFile szSrcFilePath, szDstFilePath, TRUE
	
	
	'開啟備註用的txt檔
	if (objFileSystem.FileExists(szTxtFilePath)) then
        objFileSystem.DeleteFile(szTxtFilePath)
    end if
       
    Set objWritedTextFile = objFileSystem.CreateTextFile(szTxtFilePath, -1, 0)
    
    objWritedTextFile.WriteLine request("y1") & FormatString(request("m1"), 2) & FormatString(request("d1"), 2)
    objWritedTextFile.WriteLine request("y2") & FormatString(request("m2"),2) & FormatString(request("d2"), 2)
    
    objWritedTextFile.Close
	
	'================step2: 依指定日期清除原始資料庫中的data =============
	'先查出該日期區間內的所有船隻
	dim StartDate, EndDate
	
	StartDate=DateSerial(request("y1"),request("m1"),request("d1"))
	EndDate=DateSerial(request("y2"),request("m2"),request("d2"))
	
	sql = "select ID from VesselList where date between #" & StartDate & "# and #" & EndDate & "#"
	set rs = Conn.Execute(sql)
	
	'依船ID來刪除以下table中的data
	while not rs.eof
		'1. Delete from FreightForm
		sql = "delete from FreightForm where VesselID = '" + CStr(rs("ID")) + "'"
		Conn.Execute(sql)
		
		'2. Delete from StoreSum
		sql = "delete from StoreSum where VesselListID = '" + CStr(rs("ID")) + "'"
		Conn.Execute(sql)
		
		'3. Delete from SumReportSearchRemark
		sql = "delete from SumReportSearchRemark where VesselListID = '" + CStr(rs("ID")) + "'"
		Conn.Execute(sql)
		
		'4. Delete from VesselList
		sql = "delete from VesselList where ID = '" + CStr(rs("ID")) + "'"
		Conn.Execute(sql)
		
		rs.movenext
	wend

	rs.close
	conn.close
	set rs = nothing
	set conn = nothing

	'================step3: 壓縮資料庫 =============
	Set Cd =Server.CreateObject("JRO.JetEngine") 
	If Err.number<>0 Then 
		Response.Write ("無法壓縮，請檢查錯誤資訊<br>" & Err.number & "<br>" & Err.Description) 
		Err.Clear 
	End If 
	
	dim szCompactDatabaseSrc, szCompactDatabaseDst, szDstFilePathTmp
	
	szDstFilePathTmp = szSrcFilePathTmp & "Temp.mdb"
	
	szCompactDatabaseSrc = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & szSrcFilePath
	szCompactDatabaseDst = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & szDstFilePathTmp & ";Encrypt Database=True"
	
	call Cd.CompactDatabase(szCompactDatabaseSrc, szCompactDatabaseDst) 
	
	if objFileSystem.FileExists(szSrcFilePath) then
		objFileSystem.DeleteFile(szSrcFilePath)
	end if
	
	objFileSystem.CopyFile szDstFilePathTmp, szSrcFilePath
	objFileSystem.DeleteFile(szDstFilePathTmp)
		
	set objFileSystem	= nothing

	response.write("<font size=6><br><meta http-equiv=""refresh"" content=""2; url=BKfun01a.asp""><center>備份成功!</center>")    
%>