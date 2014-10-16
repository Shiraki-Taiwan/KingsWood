<!-- #include file = ../GlobalSet/conn.asp -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>匯入資料庫</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="BKfun02.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
<!-- #include file = ../title.htm -->
<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@TWKIL") < 0 then
		response.redirect "../Login/Login.asp"
	end if

	dim szWorkingDir, szFolderPath, szINSDB_Bk
	dim objFolder, objSubFolders, SubFolder
	
	Set objFileSystem	= CreateObject ("Scripting.FileSystemObject")
	
	szWorkingDir		= Request.ServerVariables ("PATH_TRANSLATED")	
	szFolderPath		= objFileSystem.GetParentFolderName (szWorkingDir)
	szINSDB_Bk			= objFileSystem.GetParentFolderName (szWorkingDir)
	szINSDB_Bk			= objFileSystem.GetParentFolderName (szINSDB_Bk)

	'szFolderPath		= szFolderPath & "\Database"
	'szINSDB_Bk			= szINSDB_Bk & "\Database\INSDB_BK.mdb"
	
	szFolderPath		= szINSDB_Bk & "\App_Data"
	szINSDB_Bk			= szINSDB_Bk & "\App_Data\INSDB_BK.mdb"
	
	set objFolder		= objFileSystem.GetFolder(szFolderPath)
	set objSubFolders	= objFolder.SubFolders
%>

<form name="form" method="post" action="BKfun02b.asp" onsubmit=" javascript: return checkform();" OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td> 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                    <tr bgcolor=#3366cc > 
                        <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                        <td align=middle width="100%" ><div align="center"><font color="#FFFFFF">匯入資料庫</font></div></td>
                        <td width=1><img height=26 src="../image/coin2rtb.gif" width=20></td>
                    </tr>
                </tbody> 
            </table>
        </td>
    </tr>
    <tr> 
        <td> 
            <table cellspacing=0 cellpadding=1 width="100%" bgcolor=#000000 border=0 height="8">
                <tbody> 
                    <tr>
                        <td height="24"> 
                            <table cellspacing="0" cellpadding="0" width="100%" bgcolor="#ebebeb" border="0" height="38">
                                <tbody>
                                    
                                    <tr>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="20%"></td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="30%">備份日期</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="50%">資料日期</td>                                       
                                    </tr>
                             		<%
                             			dim szTxtFilePath, objReadTextFile, szStartDate, szEndDate
                             			for Each SubFolder in objSubFolders
                             				szTxtFilePath = szFolderPath & "\" & SubFolder.name & "\Backup.txt" 
                             				
                             				Set objReadTextFile = objFileSystem.OpenTextFile(szTxtFilePath, 1, 0, 0) 
                             				szStartDate = objReadTextFile.ReadLine
                             				szEndDate = objReadTextFile.ReadLine
                             				objReadTextFile.close
                             		%>
                                        <tr>
                                            <td align=right bgcolor=#C9E0F8 height="2">
                                                <input checked type="radio" name="DBName" value="<%=SubFolder.name%>" style="border-style: solid; border-color: #C9E0F8" />
                                            </td>
            	                            <td align=left bgcolor=#C9E0F8 height="2">
            	                            <%
            	                            	response.write Mid(SubFolder.name, 1, 4) & "/" & Mid(SubFolder.name, 5, 2) & "/" & Mid(SubFolder.name, 7, 2)
            	                            	response.write "　"
            	                            	response.write Mid(SubFolder.name, 10, 2) & ":" & Mid(SubFolder.name, 12, 2)
            	                            %>
            	                            </td>
                                            <td align=left bgcolor=#C9E0F8 height="2">
                                            <%
                                            	response.write Mid(szStartDate, 1, 4) & "/" & Mid(szStartDate, 5, 2) & "/" & Mid(szStartDate, 7, 2)
                                            	response.write "~"
                                            	response.write Mid(szEndDate, 1, 4) & "/" & Mid(szEndDate, 5, 2) & "/" & Mid(szEndDate, 7, 2)
                                            %>
                                            </td>                                           
                                                                                
                                        </tr>
                             		<%
                             			next
                             		%>
                                </tbody> 
                            </table>
                        </td>
                    </tr>
                </tbody>
            </table>
        </td>
    </tr>
    <tr> 
        <td> 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                    <tr> 
                        <td width=1><img height=38 src="../image/box1.gif" width=20></td>
                        <td width=1><img height=38 src="../image/box2.gif" width=9></td>
                        <td valign=center align=middle width=1 background="../image/box3.gif">&nbsp; </td>
                        <td width=1><img height=38 src="../image/box4.gif" width=27></td>
                        <td valign=center align=middle width="100%" bgcolor=#3366cc> 
                            <div align="center">
                                <input type="button" name="Import" style="background-color:#C9E0F8; border-color:#C9E0F8; border-style: outset; border-width: 2" value="匯入" OnKeyDown="SendByKeyPress()" OnMouseUp="Send()"  onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                                <%
                                	if objFileSystem.FileExists(szINSDB_Bk) then
                                %>
                                		<input type="button" name="Restore" style="background-color:#C9E0F8; border-color:#C9E0F8; border-style: outset; border-width: 2" value="回復成原始資料庫" OnKeyDown="RestoreByKeyPress()" OnMouseUp="OnRestore()"  onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                                <%		
                                	end if
                                %>
                            </div>
                        </td>
                        <td width=1><img height=38 src="../image/box5.gif" width=20></td>
                    </tr>
                </tbody>
            </table> 
        </td>
    </tr>
</table>   

<br />

<%
    conn.close
    set conn=nothing
%> 

</form>
</body>
</html>
