<!-- #include file = ../GlobalSet/conn.asp -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>公司資料設定</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="COfun01.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<%
    dim sql, rs
    set rs = nothing
    
    '===========接參數===========
    dim szID, szChineseName, szEnglishName, szStatus
    szID            = request ("ID")            '編號
    szChineseName   = request ("ChineseName")   '中文名稱
    szEnglishName   = request ("EnglishName")   '英文名稱
    szStatus        = request ("Status")
    
    if szStatus = "" then
        szStatus = "Add"    'Default value
    end if
    
    '===========查詢欲修改的資料===========
    dim szIDTmp, szChineseNameTmp, szEnglishNameTmp, szFaxNo_1Tmp, szFaxNo_2Tmp, szPhone_1Tmp, szPhone_2Tmp, szAddressTmp        
    dim szContact_1Tmp, szContact_2Tmp, szMobile_1Tmp, szMobile_2Tmp, szContainerYardTmp
    
    if szStatus = "Modify" then
        sql = "select * from CompanyInfo where ID = '" + szID + "'"
        set rs = conn.execute(sql)
        
        if not rs.eof then
            szIDTmp          = rs("ID")
            szChineseNameTmp = rs("ChineseName")
            szEnglishNameTmp = rs("EnglishName")
            szFaxNo_1Tmp     = rs("FaxNo_1")
            szFaxNo_2Tmp     = rs("FaxNo_2")
            szPhone_1Tmp     = rs("Phone_1")
            szPhone_2Tmp     = rs("Phone_2")
            szContact_1Tmp   = rs("Contact_1")
            szContact_2Tmp   = rs("Contact_2")
            szMobile_1Tmp    = rs("Mobile_1")
            szMobile_2Tmp    = rs("Mobile_2")
            szAddressTmp     = rs("Address")
            szContainerYardTmp = rs("ContainerYard")
            
            '清除參數
            szID          = ""
            szChineseName = ""
            szEnglishName = ""
        end if
    end if
    
%>


<form name="form" method="post" action="COfun01b.asp" onsubmit=" javascript: return checkform();" OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                    <tr bgcolor=#3366cc > 
                        <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                        <td align=middle width="100%" ><div align="center"><font color="#FFFFFF" >公司資料設定</font></div></td>
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
                                        <td align=right bgcolor=#C9E0F8 height="2" width="20%">編　　　　號：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="17%"> 
                                            <input type="text" name="ID" size="10" maxlength="10" value="<%=szIDTmp%>" onfocus=SelectText(0) OnKeyDown="ChangeFocusByPressedKey(15,1,1)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2" width="15%"><span style="letter-spacing: 8pt">貨櫃場：</sapn></td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="17%"><input type="text" name="ContainerYard" size="15" maxlength="20" value="<%=szContainerYardTmp%>" onfocus=SelectText(1) OnKeyDown="ChangeFocusByPressedKey(0,2,2)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"> 
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2" width="11%"></td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="17%"> 
                                        </td>
                                    </tr>
        
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2">公司中文名稱：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" colspan="5"> 
                                            <input type="text" name="ChineseName" size="50" maxlength="50" value="<%=szChineseNameTmp%>" onfocus=SelectText(2) OnKeyDown="ChangeFocusByPressedKey(1,3,3)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align=right bgcolor=#C9E0F8 height="2">公司英文名稱：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" colspan="5"> 
                                            <input type="text" name="EnglishName" size="50" maxlength="50" value="<%=szEnglishNameTmp%>" onfocus=SelectText(3) OnKeyDown="ChangeFocusByPressedKey(2,4,4)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
                                    </tr>
        
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2"><span style="letter-spacing: 4pt">傳真電話1：</sapn></td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="FaxNo_1" size="15" maxlength="13" value="<%=szFaxNo_1Tmp%>" onfocus=SelectText(4) OnKeyDown="ChangeFocusByPressedKey(3,5,5)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2">傳真電話2：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" > 
                                            <input type="text" name="FaxNo_2" size="15" maxlength="13" value="<%=szFaxNo_2Tmp%>" onfocus=SelectText(5) OnKeyDown="ChangeFocusByPressedKey(4,6,6)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2"></td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                        </td>
                                    </tr>
        
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2"><span style="letter-spacing: 8pt">聯絡人1</sapn>：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="Contact_1" size="15" maxlength="10" value="<%=szContact_1Tmp%>" onfocus=SelectText(6) OnKeyDown="ChangeFocusByPressedKey(5,7,7)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2">聯絡電話1：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="Phone_1" size="15" maxlength="13" value="<%=szPhone_1Tmp%>" onfocus=SelectText(7) OnKeyDown="ChangeFocusByPressedKey(6,8,8)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2">手機1：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="Mobile_1" size="10" maxlength="10" value="<%=szMobile_1Tmp%>" onfocus=SelectText(8) OnKeyDown="ChangeFocusByPressedKey(7,9,9)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align=right bgcolor=#C9E0F8 height="2"><span style="letter-spacing: 8pt">聯絡人2</sapn>：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="Contact_2" size="15" maxlength="10" value="<%=szContact_2Tmp%>" onfocus=SelectText(9) OnKeyDown="ChangeFocusByPressedKey(8,10,10)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2">聯絡電話2：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="Phone_2" size="15" maxlength="13" value="<%=szPhone_2Tmp%>" onfocus=SelectText(10) OnKeyDown="ChangeFocusByPressedKey(9,11,11)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2">手機2：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="Mobile_2" size="10" maxlength="10" value="<%=szMobile_2Tmp%>" onfocus=SelectText(11) OnKeyDown="ChangeFocusByPressedKey(10,12,12)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
        
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2">地　　　　址：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" colspan="5"> 
                                            <input type="text" name="Address" size="50" maxlength="50" value="<%=szAddressTmp%>" onfocus=SelectText(12) OnKeyDown="ChangeFocusByPressedKey(11,13,13)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
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
                                <input type="button" style="background-color:#C9E0F8; border-style: outset; border-width: 2" value="儲存" OnKeyDown="AddByKeyPress()" OnMouseUp="Add()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                                <input type="button" style="background-color:#C9E0F8; border-style: outset; border-width: 2" value="查詢" OnKeyDown="SearchByKeyPress()" OnMouseUp="Search()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                                <input type="button" style="background-color:#C9E0F8; border-style: outset; border-width: 2" value="刪除" OnKeyDown="DeleteByKeyPress()" OnMouseUp="Delete()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                                <input type="button" style="background-color:#C9E0F8; border-style: outset; border-width: 2" value="重置" OnKeyDown="ResetByKeyPress()" OnMouseUp="Reset()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
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
    sql = "select * from CompanyInfo where ID like '%" + szID + "%' and ChineseName like '%" + szChineseName + "%'"
    sql = sql + " and EnglishName like '%" + szEnglishName + "%' order by ID"
    
    set rs = conn.execute(sql)   
%>

<A name=D></A>
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
        <table bgcolor=#C9E0F8 border=0 bordercolor=#000000 cellspacing=1 width=100% align="center">
            <tr bgcolor="#3366cc" align="center"> 
                <td width="10%" align="left"><font color="#FFFFFF">編號</font></td>
                <td width="25%" align="left"><font color="#FFFFFF">公司中文名稱</font></td>
                <td width="25%" align="left"><font color="#FFFFFF">公司英文名稱</font></td>
                <td width="20%" align="left"><font color="#FFFFFF">傳真電話</font></td>
                <td width="20%" align="left"><font color="#FFFFFF">聯絡電話</font></td>
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
                    <td width="10%" bgcolor=#C9E0F8 align="left"><a href="COfun01a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("ID")%><a/></td>
                    <td width="25%" bgcolor=#C9E0F8 align="left"><a href="COfun01a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("ChineseName")%><a/></td>
                    <td width="25%" bgcolor=#C9E0F8 align="left"><a href="COfun01a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("EnglishName")%><a/></td>
                    <td width="20%" bgcolor=#C9E0F8 align="left"><a href="COfun01a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("FaxNo_1")%><a/></td>
                    <td width="20%" bgcolor=#C9E0F8 align="left"><a href="COfun01a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("Phone_1")%><a/></td>
                </tr>
            <%
                    rs.movenext
                wend
            %>
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
