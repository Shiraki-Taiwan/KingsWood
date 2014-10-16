<!-- #include file = ../GlobalSet/conn.asp -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>航線資料設定</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="VOfun03.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<%
    dim sql, rs
    set rs = nothing
    
    '===========接參數===========
    dim szID, szName, szStatus
    szID     = request ("ID")
    szName   = request ("Name")
    szStatus = request ("Status")
    
    if szStatus = "" then
        szStatus = "Add"    'Default value
    end if
    
    '===========查詢欲修改的資料===========
    dim szIDTmp, szNameTmp        
    if szStatus = "Modify" then
        sql = "select * from VesselLine where ID = '" + szID + "'"
        set rs = conn.execute(sql)
        
        if not rs.eof then
            szIDTmp   = rs("ID")
            szNameTmp = rs("Name")
            
            '清除參數
            szID    = ""
            szName  = ""
        end if
    end if
    
%>

<form name="form" method="post" action="VOfun03b.asp" onsubmit=" javascript: return checkform();" OnKeyDown="CheckHotKey()">

<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                    <tr bgcolor=#3366cc > 
                        <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                        <td align=middle width="100%" ><div align="center"><font color="#FFFFFF" >航線資料設定</font></div></td>
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
                                        <td align=right bgcolor=#C9E0F8 height="2" width="30%">編　　號：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="20%"> 
                                            <input type="text" name="ID" size="5" maxlength="10" value="<%=szIDTmp%>" onfocus=SelectText(0) OnKeyDown="ChangeFocus(1)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2" width="15%">名　　稱：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="35%"> 
                                            <input type="text" name="Name" size="10" maxlength="20" value="<%=szNameTmp%>" onfocus=SelectText(1) OnKeyDown="ChangeFocus(2)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
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
                                <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="查詢" OnKeyDown="SearchByKeyPress()" OnMouseUp="Search()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                                <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="刪除" OnKeyDown="DeleteByKeyPress()" OnMouseUp="Delete()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
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

<input type="hidden" name="Status" value=<%=szStatus%>>
<input type="hidden" name="IDToModify" value=<%=szIDTmp%>>

<br>

<%
    '查詢出底下的list
    set rs = nothing 
    sql = "select * from VesselLine where ID like '%" + szID + "%' and Name like '%" + szName + "%' order by ID"
    set rs = conn.execute(sql)   
%>

<A name=D></A>
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
        <table bgcolor=#C9E0F8 border=0 bordercolor=#000000 cellspacing=1 width=100% align="center">
            <tr bgcolor="#3366cc" align="center"> 
                <td width="20%" align="left"><font color="#FFFFFF">編號</font></td>
                <td width="80%" align="left"><font color="#FFFFFF">名稱</font></td>
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
                    <td width="20%" bgcolor=#C9E0F8 align="left"><a href="VOfun03a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("ID") %><a/></td>
                    <td width="80%" bgcolor=#C9E0F8 align="left"><a href="VOfun03a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("Name") %><a/></td>
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
