<!-- #include file = ../GlobalSet/conn.asp -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>e-Mail設定</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="FSfun04a.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<%
    dim sql, rs
    set rs = nothing
    
    '===========查詢欲修改的資料===========
    dim szMailServer, szSenderMail        
    
    sql = "select * from MailInfo"
    set rs = conn.execute(sql)
    
    if not rs.eof then
        szMailServer   = rs("MailServer")
        szSenderMail = rs("SenderMailAddr")
        
    end if
    
%>

<form name="form" method="post" action="FSfun04b.asp" onsubmit=" javascript: return checkform();" OnKeyDown="CheckHotKey()">

<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                    <tr bgcolor=#3366cc > 
                        <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                        <td align=middle width="100%" ><div align="center"><font color="#FFFFFF" >e-Mail設定</font></div></td>
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
                                        <td align=right bgcolor=#C9E0F8 height="2" width="40%">Mail Server：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="60%"> 
                                            <input type="text" name="MailServer" size="30" maxlength="20" value="<%=szMailServer%>" onfocus=SelectText(0) OnKeyDown="ChangeFocus(1)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align=right bgcolor=#C9E0F8 height="2">Sender's Mail：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="SenderMail" size="30" maxlength="50" value="<%=szSenderMail%>" onfocus=SelectText(1) OnKeyDown="ChangeFocus(2)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
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
                                <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="儲存" OnKeyDown="AddByKeyPress()" OnMouseUp="Add()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                                <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="重置" OnKeyDown="ResetByKeyPress()" OnMouseUp="Reset()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                            </div>
                        </td>
                        <td width=1><img height=38 src="../image/box5.gif" width=20></td>
                    </tr>
                </tbody>
            </table> 
        </td>
    </tr>
</table>

<%
    rs.close
    conn.close
    
    set rs=nothing
    set conn=nothing
%>  
 
</form>
</body>
</html>
