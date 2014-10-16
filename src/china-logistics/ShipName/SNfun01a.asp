<!-- #include file = ../GlobalSet/conn.asp -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>船名資料設定</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="SNfun01.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<%
    dim sql, rs
    set rs = nothing
    
    '===========接參數===========
    dim szID, szName, szStatus, szOwner
    szID     = request ("ID")
    szName   = request ("Name")
    szOwner  = request ("Owner")
    szStatus = request ("Status")
    
    if szStatus = "" then
        szStatus = "Add"    'Default value
    end if
    
    '===========查詢欲修改的資料===========
    dim szIDTmp, szNameTmp, szOwnerTmp        
    if szStatus = "Modify" then
        sql = "select * from VesselInfo where ID = '" + szID + "'"
        set rs = conn.execute(sql)
        
        if not rs.eof then
            szIDTmp    = rs("ID")
            szNameTmp  = rs("Name")
            szOwnerTmp = rs("Owner")
            
            '清除參數
            szID    = ""
            szName  = ""
        end if
    end if
    
%>

<form name="form" method="post" action="SNfun01b.asp" onsubmit=" javascript: return checkform();" OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
<tr> 
    <td > 
        <table cellspacing=0 cellpadding=0 width="100%" border=0>
            <tbody> 
                <tr bgcolor=#3366cc > 
                    <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                    <td align=middle width="100%" ><div align="center"><font color="#FFFFFF" >船名資料設定</font></div></td>
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
                                    <td align=right bgcolor=#cccccc height="2" width="45%">船　碼：</td>
                                    <td align=left bgcolor=#cccccc height="2" width="55%"> 
                                        <input type="text" name="ID" size="10" maxlength="10" value="<%=szIDTmp%>" onfocus=SelectText(0) OnKeyDown="ChangeFocus(1)">
                                    </td>
                                </tr>
                                <tr>
                                    <td align=right bgcolor=#cccccc height="2">船　名：</td>
                                    <td align=left bgcolor=#cccccc height="2"> 
                                        <input type="text" name="Name" size="20" maxlength="20" value="<%=szNameTmp%>" onfocus=SelectText(1) OnKeyDown="ChangeFocus(2)">
                                    </td>
                                </tr>
                                <tr>
                                    <td align=right bgcolor=#cccccc height="2" >船公司：</td>
                                    <td align=left bgcolor=#cccccc height="2" > 
                                        <input type="text" name="Owner" size="20" maxlength="20" value="<%=szOwnerTmp%>" onfocus=SelectText(2) OnKeyDown="ChangeFocus(3)">
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
                            <input type="button" style="border-style: outset; border-width: 2" value="儲存" OnKeyDown="AddByKeyPress()" OnMouseUp="Add()" >
                            <input type="button" style="border-style: outset; border-width: 2" value="查詢" OnKeyDown="SearchByKeyPress()" OnMouseUp="Search()">
                            <input type="button" style="border-style: outset; border-width: 2" value="刪除" OnKeyDown="DeleteByKeyPress()" OnMouseUp="Delete()">
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

<input type="hidden" name="Status" value=<%=szStatus%>>
<input type="hidden" name="IDToModify" value=<%=szIDTmp%>>

<br>

<%
    '查詢出底下的list
    set rs = nothing 
    sql = "select * from VesselInfo where ID like '%" + szID + "%' and Name like '%" + szName + "%' and Owner like '%" + szOwner + "%' order by ID"
    set rs = conn.execute(sql)   
%>

<A name=D></A>
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
        <table bgcolor=#f0ebdb border=0 bordercolor=#000000 cellspacing=1 width=100% align="center">
            <tr bgcolor="#3366cc" align="center"> 
                <td width="20%" align="left"><font color="#FFFFFF">編號</font></td>
                <td width="40%" align="left"><font color="#FFFFFF">名稱</font></td>
                <td width="40%" align="left"><font color="#FFFFFF">船公司</font></td>
            </tr>
        </table>
        </td>
    </tr>
    
    <tr> 
        <td > 
            <table width="100%" cellspacing=1>		
            <%
                while not rs.eof
            %>
                <tr align="center"> 
                    <td width="20%" bgcolor=#f0ebdb align="left"><a href="SNfun01a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("ID") %><a/></td>
                    <td width="40%" bgcolor=#f0ebdb align="left"><a href="SNfun01a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("Name") %><a/></td>
                    <td width="40%" bgcolor=#f0ebdb align="left"><a href="SNfun01a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("Owner") %><a/></td>
                </tr>
            <%
                    rs.movenext
                wend
            %>
            </table>
        </td>
    </tr>

<%
    rs.close
    conn.close 
    
    set rs=nothing
    set conn=nothing
%>  
 
</form>
</body>
</html>
