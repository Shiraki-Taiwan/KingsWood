<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>備份資料庫</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="BKfun01.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
<!-- #include file = ../title.htm -->
<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@TWKIL") < 0 then
		response.redirect "../Login/Login.asp"
	end if
	
    dim szSrcFilePath

	Set objFileSystem	= CreateObject ("Scripting.FileSystemObject")
	szSrcFilePath		= Request.ServerVariables ("PATH_TRANSLATED")
	szSrcFilePath		= objFileSystem.GetParentFolderName (szSrcFilePath)
	szSrcFilePath		= objFileSystem.GetParentFolderName (szSrcFilePath)
	
	'szSrcFilePath		= szSrcFilePath & "\Database\INSDB_BK.mdb"
	szSrcFilePath		= szSrcFilePath & "\App_Data\INSDB_BK.mdb"
	
	if objFileSystem.FileExists(szSrcFilePath) then
		response.write("<font size=6><br><meta http-equiv=""refresh"" content=""3; url=BKfun02a.asp""><center>您目前使用中的資料庫為備份資料庫, <br>請先回復成原始資料庫!</center>")    
		response.end
	end if
%>
<form name="form" method="post" action="BKfun01b.asp" onsubmit=" javascript: return checkform();" OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                    <tr bgcolor=#3366cc > 
                        <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                        <td align=middle width="100%" ><div align="center"><font color="#FFFFFF">備份資料庫</font></div></td>
                        <td width=1><img height=26 src="../image/coin2rtb.gif" width=20></td>
                    </tr>
                </tbody> 
            </table>
        </td>
    </tr>
    <tr> 
        <td > 
            <table cellspacing=0 cellpadding=1 width="100%" bgcolor=#000000 border=0 height="8">
                <tbody> 
                    <tr>
                        <td height="24"> 
                            <table cellspacing=0 cellpadding=0 width="100%" bgcolor=#ebebeb border=0 height="38">
                                <tbody>
                                    <tr>
                                    	<td align=right bgcolor=#C9E0F8 height="2"></td>
                                    	<td align=left bgcolor=#C9E0F8 height="2" colspan=2>請選擇在備份好資料庫後, 欲刪除的資料日期區間：</td>
                                    </tr>
                                    <tr>
                                    	<td align=right bgcolor=#C9E0F8 height="2" width="30%"></td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="10%">開始日期：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="60%"> 
                                            <input type="text" name="Y1" size="1" maxlength="4" onfocus=SelectText(0) OnKeyDown="ChangeFocus(1)"  onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">年
                                            <input type="text" name="M1" size="1" maxlength="2"  onfocus=SelectText(1) onblur=MonthChecking(this) OnKeyDown="ChangeFocus(2)"  onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">月
                                            <input type="text" name="D1" size="1" maxlength="2"  onfocus=SelectText(2) onblur=DayChecking(this) OnKeyDown="ChangeFocus(3)"  onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">日                
                                        </td>                                        
                                    </tr>
                                    
                                    <tr>
                                    	<td align=right bgcolor=#C9E0F8 height="2"></td>
                                        <td align=left bgcolor=#C9E0F8 height="2">結束日期：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="Y2" size="1" maxlength="4" onfocus=SelectText(3) OnKeyDown="ChangeFocus(4)"  onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">年
                                            <input type="text" name="M2" size="1" maxlength="2"  onfocus=SelectText(4) onblur=MonthChecking(this) OnKeyDown="ChangeFocus(5)"  onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">月
                                            <input type="text" name="D2" size="1" maxlength="2"  onfocus=SelectText(5) onblur=DayChecking(this) OnKeyDown="ChangeFocus(6)"  onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">日                
                                        </td>                                        
                                    </tr>
                                    
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
                                <input type="button" style="background-color:#C9E0F8; border-color:#C9E0F8; border-style: outset; border-width: 2" value="進行備份" OnKeyDown="SendByKeyPress()" OnMouseUp="Send()"  onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                            </div>
                        </td>
                        <td width=1><img height=38 src="../image/box5.gif" width=20></td>
                    </tr>
                </tbody>
            </table> 
        </td>
    </tr>
</table>   

<br> 


</form>
</body>
</html>
