<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>公司資料列印查詢</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="COfun02.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<form name="form" method="post" action="COfun02b.asp" OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                    <tr bgcolor=#3366cc > 
                        <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                        <td align=middle width="100%" ><div align="center"><font color="#FFFFFF" >公司資料列印查詢</font></div></td>
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
                            <table cellspacing=0 cellpadding=0 width="100%" bgcolor=#ebebeb border=0 height="38">
                                <tbody>
	                                <tr> 
                                        <td align=right bgcolor=#cccccc height="2" width="40%">編　　　　號：</td>
                                        <td align=left bgcolor=#cccccc height="2" width="60%"> 
                                            <input type="text" name="ID" size="10" maxlength="10" value="" OnKeyDown="ChangeFocus(1)">
                                        </td>                                   
                                    </tr>
        
                                    <tr> 
                                        <td align=right bgcolor=#cccccc height="2">公司中文名稱：</td>
                                        <td align=left bgcolor=#cccccc height="2"> 
                                            <input type="text" name="ChineseName" size="30" maxlength="20" value="" OnKeyDown="ChangeFocus(2)">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align=right bgcolor=#cccccc height="2">公司英文名稱：</td>
                                        <td align=left bgcolor=#cccccc height="2"> 
                                            <input type="text" name="EnglishName" size="30" maxlength="20" value="" OnKeyDown="ChangeFocus(3)">
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
                                <input type="button" style="border-style: outset; border-width: 2" value="查詢" OnKeyDown="SearchByKeyPress()" OnMouseUp="Search()">
                                <input type="button" style="border-style: outset; border-width: 2" value="重置" OnKeyDown="ResetByKeyPress()" OnMouseUp="Reset()">
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
