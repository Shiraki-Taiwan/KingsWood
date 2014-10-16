<!-- #include file = ../GlobalSet/conn.asp -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>攬貨報告書列印/傳真</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="FSfun01a.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<%
    dim sql, rs
    set rs = nothing
    
    '===========接參數===========
    dim szVesselListID, szVesselName, szVesselNo, szYear, szMonth, szDay, szDate, szVesselLine
    szVesselListID = request ("VesselListID")
    szVesselName   = request ("VesselName")
    szVesselNo     = request ("VesselNo")
    szYear         = request ("Year")
    szMonth        = request ("Month")
    szDay          = request ("Day")
    szVesselLine   = request ("VesselLine")
    
    If Trim(szYear) <> "" and Trim(szMonth) <> "" and Trim(szDay) <> "" then 
        szDate = CStr(szYear) + "/" + CStr(szMonth) + "/" + CStr(szDay)
    else
        szDate = "0"
    end if
    
    '序號
    dim nSerialNum, nSerialNumTmp
    nSerialNum = 1
    nSerialNumTmp = request("SerialNum")
%>

<form name="form" method="post" action="FSfun01b.asp" onsubmit=" javascript: return checkform();"  OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
<tr> 
    <td > 
        <table cellspacing=0 cellpadding=0 width="100%" border=0>
            <tbody> 
                <tr bgcolor=#3366cc > 
                    <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                    <td align=middle width="100%" ><div align="center"><font color="#FFFFFF" >攬貨報告書列印/傳真</font></div></td>
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
                                    <td align=right bgcolor=#C9E0F8 height="2" width="17%">編　號：</td>
                                    <td align=left bgcolor=#C9E0F8 height="2" width="31%"> 
                                        <input type="text" name="SerialNum" size="10" maxlength="4" value="<%=nSerialNumTmp%>" onfocus="SelectText(0)" OnKeyDown="OnSerialNum()" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">                                            
                                    </td>
                                    <td align=right bgcolor=#C9E0F8 height="2" width="17%"></td>
                                    <td align=left bgcolor=#C9E0F8 height="2" width="31%"></td>
                                </tr>
        	                    <tr>
                                    <td align=right bgcolor=#C9E0F8 height="2" width="17%">船　名：</td>
                                    <td align=left bgcolor=#C9E0F8 height="2" width="31%">
                                        <input type="text" name="VesselName" size="30" maxlength="22" value="<%=szVesselName%>" onfocus=SelectText(1) OnKeyDown="ChangeFocusByPressedKey(9,2,2)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                    </td>   
                                    <td align=right bgcolor=#C9E0F8 height="2" width="17%">航　次：</td>
                                    <td align=left bgcolor=#C9E0F8 height="2" width="31%"> 
                                        <input type="text" name="VesselNo" size="8" maxlength="10" value="<%=szVesselNo %>" onfocus=SelectText(2) OnKeyDown="ChangeFocusByPressedKey(1,3,3)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                    </td>       
                                </tr>
    
    
                                <tr>
                                    <td align=right bgcolor=#C9E0F8 height="2">船公司：</td>
                                    <td align=left bgcolor=#C9E0F8 height="2"> 
                                        <input type="text" name="Owner" size="8" maxlength="15" value="<%=szOwner%>" onfocus=SelectText(3) OnKeyDown="ChangeFocusByPressedKey(2,4,4)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                    </td> 
                                    <td align=right bgcolor=#C9E0F8 height="2">船　掛：</td>
                                    <td align=left bgcolor=#C9E0F8 height="2"> 
                                        <input type="text" name="CheckInID" size="8" maxlength="20" value="<%=szCheckInID %>" onfocus=SelectText(4) OnKeyDown="ChangeFocusByPressedKey(3,5,5)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                    </td>
                                </tr>
                                <tr>
                                    <td align=right bgcolor=#C9E0F8 height="2">結關日：</td>
                                    <td align=left bgcolor=#C9E0F8 height="2"> 
                                        <input type="text" name="Year" size="1" maxlength="4" value="<%=szYear%>" onfocus=SelectText(5) OnKeyDown="ChangeFocusByPressedKey(4,6,6)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">年
                                        <input type="text" name="Month" size="1" maxlength="2"  onfocus=SelectText(6) onblur=CheckMonth() value="<%=szMonth%>" OnKeyDown="ChangeFocusByPressedKey(5,7,7)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">月
                                        <input type="text" name="Day" size="1" maxlength="2"  onfocus=SelectText(7) onblur=CheckDay() value="<%=szDay%>" OnKeyDown="ChangeFocusByPressedKey(6,8,8)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">日                
                                    </td>
                                    <td align=right bgcolor=#C9E0F8 height="2">航　線：</td>
                                    <td align=left bgcolor=#C9E0F8 height="2">                                        
                                        <select size="1" name="VesselLine" OnKeyDown="ChangeFocus(9)">             
                                            <option selected value="0">--請選擇--</option>
                                            <%   
                                                dim rs1
                                                set rs1 = nothing                                                
                                                
                                                '查詢航線
                                                sql = "select ID, Name from VesselLine order by ID"
                                                
                                                set rs1 = conn.execute(sql)
                                            
                                                while not rs1.eof
                                                    if rs1("ID") = szVesselLineTmp then 
                                                        response.write "<option selected value=" & rs1("ID") & ">" & rs1("Name") &"</option>"            		 	
                                                    else
                                                        response.write "<option value=" & rs1("ID") & ">" & rs1("Name") & "</option>"            		 	
                                                    end if            		 	
                                                    
                                            	    rs1.movenext
                                                wend
                                                
                                                rs1.close
                                                set rs1 = nothing
                                            %>
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
                            <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="重置" OnKeyDown="ResetByKeyPress()" OnMouseUp="Reset()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                        </div>
                    </td>
                    <td width=1><img height=38 src="../image/box5.gif" width=20></td>
                </tr>
            </tbody>
        </table> 
    </td>
</tr>

<%
    set rs = nothing 
    
    sql = "select * from VesselList where VesselNo like '%" + szVesselNo + "%'"
    
    if szVesselName <> "0" then
        sql = sql + " and VesselName like '%" + szVesselName + "%'"
    end if
    
    sql = sql + " and Owner like '%" + szOwner + "%' and CheckInID like '%" + szCheckInID + "%'"
    
    if szDate <> "0" then
        sql = sql + " and Date = #" + szDate + "#"
    end if
    
    if szVesselLine <> "0" and szVesselLine <> "" then
        sql = sql + " and VesselLine = '" + szVesselLine + "'"
    end if
    
    sql = sql + " order by Date DESC, DataCount DESC ,ID DESC"
    
    set rs = conn.execute(sql)
%>

<A name=D></A>
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
        <table bgcolor=#C9E0F8 border=0 bordercolor=#000000 cellspacing=1 width=100% align="center">
            <tr bgcolor="#3366cc" align="center"> 
                <td width="6%" align="left"><font color="#FFFFFF">編號</font></td>
                <td width="24%" align="left"><font color="#FFFFFF">船名</font></td>
                <td width="15%" align="left"><font color="#FFFFFF">航次</font></td>
                <td width="15%" align="left"><font color="#FFFFFF">船公司</font></td> 
                <td width="15%" align="left"><font color="#FFFFFF">船掛</font></td>                
                <td width="15%" align="left"><font color="#FFFFFF">結關日</font></td>
                <td width="10%" align="left"><font color="#FFFFFF">航線</font></td>
            </tr>
        </table>
        </td>
    </tr>
    
    <tr> 
        <td > 
            <table width="100%" cellspacing=1>		
            <%
                while not rs.eof
                
                    '20-Apr2005: Serial number改成4位數
                    if nSerialNum < 10 then
                        nSerialNum = "000" & nSerialNum
                    elseif nSerialNum < 100 then
                        nSerialNum = "00" & nSerialNum
                    elseif nSerialNum < 1000 then
                        nSerialNum = "0" & nSerialNum
                    end if
            %>
                <tr align="center"> 
                    <td width="6%" bgcolor=#C9E0F8 align="left"><a href="FSfun01b.asp?Status=FirstLoad&VesselListID=<%= rs("ID") %>&SerialNum=<%=nSerialNum%>"><%=nSerialNum%></a></td>
                    <td width="24%" bgcolor=#C9E0F8 align="left"><a href="FSfun01b.asp?Status=Modify&VesselListID=<%= rs("ID")%>&SerialNum=<%=nSerialNum%> "><%=rs("VesselName")%></a></td>
                    <td width="15%" bgcolor=#C9E0F8 align="left"><a href="FSfun01b.asp?Status=Modify&VesselListID=<%= rs("ID")%>&SerialNum=<%=nSerialNum%> "><%= rs("VesselNo")%></a></td>
                    <td width="15%" bgcolor=#C9E0F8 align="left"><a href="FSfun01b.asp?Status=Modify&VesselListID=<%= rs("ID")%>&SerialNum=<%=nSerialNum%> "><%=rs("Owner")%></a></td>                    
                    <td width="15%" bgcolor=#C9E0F8 align="left"><a href="FSfun01b.asp?Status=Modify&VesselListID=<%= rs("ID")%>&SerialNum=<%=nSerialNum%> "><%= rs("CheckInID")%></a></td>
                    <td width="15%" bgcolor=#C9E0F8 align="left"><a href="FSfun01b.asp?Status=Modify&VesselListID=<%= rs("ID")%>&SerialNum=<%=nSerialNum%> "><%= rs("Date")%></a></td>
                    <%
                        dim szLineName
                        if rs("VesselLine") <> "0" then
                            sql = "select Name from VesselLine where ID = '" + rs("VesselLine") + "'"                        
                            
                            set rs1 = conn.execute(sql)
                            
                            if not rs1.eof then
                                szLineName = rs1("Name")
                                
                                rs1.close                        	
                            end if
                            
                            set rs1 = nothing
                        else
                            szLineName = ""
                        end if
                    %>
                    <td width="10%" bgcolor=#C9E0F8 align="left"><a href="FSfun01b.asp?Status=Modify&VesselListID=<%= rs("ID") %>&SerialNum=<%=nSerialNum%>"><%= szLineName%></a></td>
                </tr>
            <%
                    szLineName = ""
                    nSerialNum = nSerialNum + 1
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
</table>   
</form>
</body>
</html>
