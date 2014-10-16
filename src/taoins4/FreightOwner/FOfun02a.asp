<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>攬貨商資料列印</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="FOfun02.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="Print();">
	<!-- #include file = ../title.htm -->
<form name="form" method="post" action="FOfun02b.asp"  OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                    <tr bgcolor=#3366cc > 
                        <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                        <td align=middle width="100%" ><div align="center"><font color="#FFFFFF" >攬貨商資料列印</font></div></td>
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
                                        <td align=right bgcolor=#C9E0F8 height="2" width="40%">編　　號：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="60%"> 
                                            <input type="text" name="ID" size="30" maxlength="30" value="<%=szIDTmp%>" onfocus=SelectText(0) OnKeyDown="ChangeFocus(1)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
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
                                <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="查詢" OnKeyDown="SearchByKeyPress()" OnMouseUp="Search()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
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



</form>
</body>
</html>
