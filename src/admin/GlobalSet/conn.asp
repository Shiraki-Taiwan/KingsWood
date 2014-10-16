<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	dim szSessionKey
	szSessionKey	= UCase(Split(Request.ServerVariables("URL"), "/")(1))

	'開啟ACCESS資料庫作資料庫的連結讀取
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("../App_Data/INSDB.mdb")

	'Conn.open "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" & Server.MapPath("../App_Data/HAS.accdb")
	'Conn.open "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" & Server.MapPath("../App_Data/INSDB.mdb")
	'Conn.open "Driver={Microsoft Access Driver (*.mdb)};" & "dbq=" & Server.MapPath("../database/INSDB.mdb")

	'"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="& Server.MapPath("資料庫名稱")
	'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Projects\GreensBox\Shiraki.ICE\App_Data\INSDB.mdb;Persist Security Info=True
%>
