<!-- #include file = ../GlobalSet/conn.asp -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>攬貨商資料設定</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="FOfun01.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@KEEINS5") < 0 then
		response.redirect "../Login/Login.asp"
	end if


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
    
    
    dim szFind
    szFind = request ("Find")           '要往前或往後一單號?
    
    dim bFindID, szPrevID
    bFindID = 0
    szPrevID = szID                     '預設第一筆
    
    '找前/後一單號
    sql = "select ID from FreightOwner group by ID order by ID"
    set rs = conn.execute(sql)
    
    dim bGotNumberPart, szStrTmp, szPreStrTmp, szCurID
    
    if szFind = "Prev" then
        while not rs.eof or bFindID <> 1
            
            '把字串中的文字部分轉成大寫
            bGotNumberPart = 0
            szStrTmp = szID
            szPreStrTmp = ""
            while bGotNumberPart <> 1
            	if IsNumeric(szStrTmp) then
            		szStrTmp = szStrTmp
            		bGotNumberPart = 1
           		else
           			szPreStrTmp = szPreStrTmp & Mid(szStrTmp, 1, 1)
           			szStrTmp = Mid(szStrTmp, 2) 
            	end if
            wend
        
            szID = UCase(szPreStrTmp) & szStrTmp
            
            
            bGotNumberPart = 0
            
            if not rs.eof then
                szStrTmp = rs("ID")
            end if
            
            szPreStrTmp = ""
            while bGotNumberPart <> 1
            	if IsNumeric(szStrTmp) then
            		szStrTmp = szStrTmp
            		bGotNumberPart = 1
           		else
           			szPreStrTmp = szPreStrTmp & Mid(szStrTmp, 1, 1)
           			szStrTmp = Mid(szStrTmp, 2) 
            	end if
            wend
            
            szCurID = UCase(szPreStrTmp) & szStrTmp
                        
            
            if szID = szCurID then
                bFindID = 1     '找到現在的ID了
                
                if not rs.eof then
                    szID = szPrevID   
                    szPrevFoundID = szID                 
                end if
                
            end if
            
            if not rs.eof then
                szPrevID = rs("ID")
                rs.movenext
            end if
        wend 
        
        nPageNum = 1
        
    elseif szFind = "Next" then
        n = 0
        while not rs.eof or bFindID <> 1
            
            '把字串中的文字部分轉成大寫
            bGotNumberPart = 0
            szStrTmp = szID
            szPreStrTmp = ""
            while bGotNumberPart <> 1
            	if IsNumeric(szStrTmp) then
            		szStrTmp = szStrTmp
            		bGotNumberPart = 1
           		else
           			szPreStrTmp = szPreStrTmp & Mid(szStrTmp, 1, 1)
           			szStrTmp = Mid(szStrTmp, 2) 
            	end if
            wend
        
            szID = UCase(szPreStrTmp) & szStrTmp
                        
            bGotNumberPart = 0
            
            if not rs.eof then
                szStrTmp = rs("ID")
            end if
            
            szPreStrTmp = ""
            while bGotNumberPart <> 1
            	if IsNumeric(szStrTmp) then
            		szStrTmp = szStrTmp
            		bGotNumberPart = 1
           		else
           			szPreStrTmp = szPreStrTmp & Mid(szStrTmp, 1, 1)
           			szStrTmp = Mid(szStrTmp, 2) 
            	end if
            wend
            
            szCurID = UCase(szPreStrTmp) & szStrTmp
            
            if szID = szCurID then
                bFindID = 1     '找到現在的ID了
                'szID = ""
                if not rs.eof then
                    rs.movenext
                end if
                
                if not rs.eof then
                    szID = rs("ID")
                    szPrevFoundID = szID
                end if
            end if
            
            if not rs.eof then  
                szPrevID = rs("ID")          
                rs.movenext
            end if
        wend
        
        nPageNum = 1
    end if
    
    
    '===========查詢欲修改的資料===========
    dim szIDTmp, szNameTmp, szFaxNo_1Tmp, szFaxNo_2Tmp, szContact_1Tmp, szContact_2Tmp, szPhone_1Tmp, szPhone_2Tmp, szAddressTmp
    dim szMailAddrTmp, szMailAddrCCTmp
    if szStatus = "Modify" then
        sql = "select * from FreightOwner where ID = '" + szID + "'"
        set rs = conn.execute(sql)
        
        if not rs.eof then
            szIDTmp          = rs("ID")
            szNameTmp        = rs("Name")
            szFaxNo_1Tmp     = rs("FaxNo_1")
            szFaxNo_2Tmp     = rs("FaxNo_2")
            szContact_1Tmp   = rs("Contact_1")
            szContact_2Tmp   = rs("Contact_2")
            szPhone_1Tmp     = rs("Phone_1")
            szPhone_2Tmp     = rs("Phone_2")
            szMailAddrTmp    = rs("MailAddr")
            szMailAddrCCTmp  = rs("MailAddrCC")
            szAddressTmp     = rs("Address")
            
            '清除參數
            szID    = ""
            szName  = ""
        end if
    end if
    
%>

<form name="form" method="post" action="FOfun01b.asp" onsubmit=" javascript: return checkform();"  OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                    <tr bgcolor=#3366cc > 
                        <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                        <td align=middle width="100%" ><div align="center"><font color="#FFFFFF" >攬貨商資料設定</font></div></td>
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
                                        <td align=right bgcolor=#C9E0F8 height="2" width="17%">編　　　號：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="31%"> 
                                            <input type="text" name="ID" size="10" maxlength="10" value="<%=szIDTmp%>" onfocus=SelectText(0) OnKeyDown="ChangeFocus(1); CheckKeyPress();" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2" width="20%"></td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="32%"></td>
                                    </tr>
                                    <tr>
                                        <td align=right bgcolor=#C9E0F8 height="2" ><span style="letter-spacing: 4pt">公司名稱：</span></td>
                                        <td align=left bgcolor=#C9E0F8 height="2" colspan="3"> 
                                            <input type="text" name="Name" size="30" maxlength="20" value="<%=szNameTmp%>" onfocus=SelectText(1) OnKeyDown="ChangeFocus(2)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
            
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2"><span style="letter-spacing: 4pt">傳真電話：</span></td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="FaxNo_1" size="15" maxlength="13" value="<%=szFaxNo_1Tmp%>" onfocus=SelectText(2) OnKeyDown="ChangeFocus(3)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2">傳真電話：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="FaxNo_2" size="15" maxlength="13" value="<%=szFaxNo_2Tmp%>" onfocus=SelectText(3) OnKeyDown="ChangeFocus(4)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
            
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2">聯　絡　人：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="Contact_1" size="15" maxlength="10" value="<%=szContact_1Tmp%>" onfocus=SelectText(4) OnKeyDown="ChangeFocus(5)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2">聯絡電話：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" > 
                                            <input type="text" name="Phone_1" size="15" maxlength="13" value="<%=szPhone_1Tmp%>" onfocus=SelectText(5) OnKeyDown="ChangeFocus(6)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
            
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2">聯　絡　人：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="Contact_2" size="15" maxlength="10" value="<%=szContact_2Tmp%>" onfocus=SelectText(6) OnKeyDown="ChangeFocus(7)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2">聯絡電話：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="Phone_2" size="15" maxlength="13" value="<%=szPhone_2Tmp%>" onfocus=SelectText(7) OnKeyDown="ChangeFocus(8)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
                                    
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2"><span style="letter-spacing: 4.5pt">E-Mail：</span></td>
                                        <td align=left bgcolor=#C9E0F8 height="2" colspan="3"> 
                                            <input type="text" name="MailAddr" size="50" maxlength="50" value="<%=szMailAddrTmp%>" onfocus=SelectText(8) OnKeyDown="ChangeFocus(9)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2">E-Mail副本：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" colspan="3"> 
                                            <input type="text" name="MailAddrCC" size="50" maxlength="50" value="<%=szMailAddrCCTmp%>" onfocus=SelectText(9) OnKeyDown="ChangeFocus(10)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2">地　　　址：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" colspan="3"> 
                                            <input type="text" name="Address" size="50" maxlength="50" value="<%=szAddressTmp%>" onfocus=SelectText(10) OnKeyDown="ChangeFocus(11)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
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
                                <input type="button" style="background-color:#C9E0F8; border-style: outset; border-width: 2" value="儲存" OnKeyDown="AddByKeyPress()" OnMouseUp="Add()"  onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
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
    sql = "select * from FreightOwner where ID like '%" + szID + "%' and Name like '%" + szName + "%' order by ID"
    set rs = conn.execute(sql)   
%>

<A name=D></A>
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
        <table bgcolor=#C9E0F8 border=0 bordercolor=#000000 cellspacing=1 width=100% align="center">
            <tr bgcolor="#3366cc" align="center"> 
                <td width="10%" align="left"><font color="#FFFFFF">編號</font></td>
                <td width="30%" align="left"><font color="#FFFFFF">公司名稱</font></td>
                <td width="20%" align="left"><font color="#FFFFFF">傳真電話</font></td>
                <td width="20%" align="left"><font color="#FFFFFF">聯絡人</font></td>
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
                    <td width="10%" bgcolor=#C9E0F8 align="left"><a href="FOfun01a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("ID") %><a/></td>
                    <td width="30%" bgcolor=#C9E0F8 align="left"><a href="FOfun01a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("Name") %><a/></td>
                    <td width="20%" bgcolor=#C9E0F8 align="left"><a href="FOfun01a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("FaxNo_1") %><a/></td>
                    <td width="20%" bgcolor=#C9E0F8 align="left"><a href="FOfun01a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("Contact_1") %><a/></td>
                    <td width="20%" bgcolor=#C9E0F8 align="left"><a href="FOfun01a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("Phone_1") %><a/></td>
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
