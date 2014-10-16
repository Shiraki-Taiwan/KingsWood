<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>倉單資料查詢</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="FSfun03a.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<form name="form" method="post" action="FSfun03b.asp" OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
<tr> 
    <td > 
        <table cellspacing=0 cellpadding=0 width="100%" border=0>
            <tbody> 
                <tr bgcolor=#3366cc > 
                    <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                    <td align=middle width="100%" ><div align="center"><font color="#FFFFFF" >倉單資料查詢</font></div></td>
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
                                    <td width="45%" align=right bgcolor=#C9E0F8 height="2">表單類別：</td>
                                    <td width="55%" align=left bgcolor=#C9E0F8 height="2">                                        
                                        <select size="1" name="ReportType" OnKeyDown="ChangeFocus(1)">             
                                            <option value="1">總表查詢</option>
                                            <option value="2">尺寸資料查詢</option>
                                        </select>
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
