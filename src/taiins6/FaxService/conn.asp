<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	dim szSessionKey
	szSessionKey		= UCase(Split(Request.ServerVariables("URL"), "/")(1))

    dim szDbWorkingDir, szOpenDbPath, szOpenDbName
	dim objFSO, objFile, objFolder, objDbTextFile, objDateLast

	Set objFSO			= Server.CreateObject("Scripting.FileSystemObject")

	objDateLast			= Null
	szDbWorkingDir		= Request.ServerVariables("PATH_TRANSLATED")
	szOpenDbPath		= objFSO.GetParentFolderName(szDbWorkingDir)
	szOpenDbPath		= objFSO.GetParentFolderName(szOpenDbPath)
    szOpenDbPath		= szOpenDbPath & "\App_Data\connection.txt" 
	
	Set objFolder		= objFSO.GetFolder(Server.MapPath("../App_Data/"))
    Set objDbTextFile	= objFSO.OpenTextFile(szOpenDbPath, 1, 1, 0)

    szOpenDbName		= objDbTextFile.ReadLine

	For Each objFile in objFolder.Files
		if InStr(objFile.Name, "INSDB") > 0 then
			if IsNull(objDateLast) then
				objDateLast		= objFile.DateLastModified
				szOpenDbName	= objFile.Name
			elseif objDateLast < objFile.DateLastModified then
				objDateLast		= objFile.DateLastModified
				szOpenDbName	= objFile.Name
			end if
		end if
	Next

	Set objDbTextFile		= Nothing
	Set objFolder			= Nothing
	Set objFSO				= Nothing

	'開啟ACCESS資料庫作資料庫的連結讀取
	Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("../App_Data/" & szOpenDbName)
    ' 加入自定資料庫名稱
	'Conn.open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("../App_Data/INSDB.mdb")

	'Conn.open "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" & Server.MapPath("../App_Data/HAS.accdb")
	'Conn.open "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" & Server.MapPath("../App_Data/INSDB.mdb")
	'Conn.open "Driver={Microsoft Access Driver (*.mdb)};" & "dbq=" & Server.MapPath("../database/INSDB.mdb")

	'"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="& Server.MapPath("資料庫名稱")
	'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\GreensBox\Shiraki.ICE\App_Data\INSDB.mdb;Persist Security Info=True
%>
