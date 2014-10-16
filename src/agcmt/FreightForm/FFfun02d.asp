<!-- #include file = ../GlobalSet/conn.asp -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>移動進倉單資料</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="FFfun02d.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
<!-- #include file = ../title.htm -->
<%
    dim sql, rs
    set rs = nothing
    
    '===========接參數===========
    dim szStatus
    szStatus = request ("Status")
    
    dim szVesselID, szVesselName, szVesselNo, szYear, szMonth, szDay, szDate, szOwner, szCheckInID, szVesselLine
    szVesselID     = request ("VesselID")
    szVesselName   = request ("VesselName")
    szVesselNo     = request ("VesselNo")
    szOwner        = request ("Owner")
    szCheckInID    = request ("CheckInID")
    szYear         = request ("Year")
    szMonth        = request ("Month")
    szDay          = request ("Day")
    
    '21-Nov2004:加入航線
    szVesselLine   = request ("VesselLine")
    
    If Trim(szYear) <> "" and Trim(szMonth) <> "" and Trim(szDay) <> "" then 
        szDate = CStr(szYear) + "/" + CStr(szMonth) + "/" + CStr(szDay)
    else
        szDate = "0"
    end if
    
    dim szIDTmp, szVesselNameTmp, szVesselNoTmp, szCheckInIDTmp, szDateTmp, szYearTmp, szMonthTmp, szDayTmp, szOwnerTmp
    dim szVesselLineTmp
    
    sql = "select * from VesselList where ID = '" + szVesselID + "'"
    set rs = conn.execute(sql)
    
    if not rs.eof then
        szIDTmp             = rs("ID")
        szVesselNameTmp     = rs("VesselName")
        szOwnerTmp          = rs("Owner")
        szVesselNoTmp       = rs("VesselNo")
        szCheckInIDTmp      = rs("CheckInID")
        szDateTmp           = rs("Date")
        szVesselLineTmp     = rs("VesselLine")  
        
        szYearTmp = Year (szDateTmp)
        szMonthTmp = Month (szDateTmp)
        szDayTmp = Day (szDateTmp)
        
        '清除參數
        szID          = ""
        szVesselName  = ""
        szOwner       = ""
        szVesselNo    = ""
        szCheckInID   = ""
        szVesselLine  = ""
    end if
    
    
%>

<form name="form" method="post" action="FFfun02e.asp" onsubmit=" javascript: return checkform();"  OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
<tr> 
    <td > 
        <table cellspacing=0 cellpadding=0 width="100%" border=0>
            <tbody> 
                <tr bgcolor=#3366cc > 
                    <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                    <td align=middle width="100%" ><div align="center"><font color="#FFFFFF" >請選擇航次來移動倉單</font></div></td>
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
                                    <td align=right bgcolor=#C9E0F8 height="2" width="17%">船　名：</td>
                                    <td align=left bgcolor=#C9E0F8 height="2" width="31%">
                                        <input type="text" name="VesselName" size="8" maxlength="15" value="<%=szVesselNameTmp%>" onfocus=SelectText(0) OnKeyDown="ChangeFocus(1)">
                                    </td>   
                                    <td align=right bgcolor=#C9E0F8 height="2" width="17%">船公司：</td>
                                    <td align=left bgcolor=#C9E0F8 height="2" width="31%"> 
                                        <input type="text" name="Owner" size="8" maxlength="15" value="<%=szOwnerTmp%>" onfocus=SelectText(1) OnKeyDown="ChangeFocus(2)">
                                    </td>        
                                </tr>
    
    
                                <tr> 
                                    <td align=right bgcolor=#C9E0F8 height="2">航　次：</td>
                                    <td align=left bgcolor=#C9E0F8 height="2"> 
                                        <input type="text" name="VesselNo" size="8" maxlength="10" value="<%=szVesselNoTmp %>" onfocus=SelectText(2) OnKeyDown="ChangeFocus(3)">
                                    </td>
                                    <td align=right bgcolor=#C9E0F8 height="2">船　掛：</td>
                                    <td align=left bgcolor=#C9E0F8 height="2"> 
                                        <input type="text" name="CheckInID" size="8" maxlength="20" value="<%=szCheckInIDTmp %>" onfocus=SelectText(3) OnKeyDown="ChangeFocus(4)">
                                    </td>
                                </tr>
                                <tr>
                                    <td align=right bgcolor=#C9E0F8 height="2">結關日：</td>
                                    <td align=left bgcolor=#C9E0F8 height="2"> 
                                        <input type="text" name="Year" size="1" maxlength="4" value="<%=szYearTmp%>" onfocus=SelectText(4) OnKeyDown="ChangeFocus(5)">年
                                        <input type="text" name="Month" size="1" maxlength="2"  onfocus=SelectText(5) onblur=CheckMonth() value="<%=szMonthTmp%>" OnKeyDown="ChangeFocus(6)">月
                                        <input type="text" name="Day" size="1" maxlength="2"  onfocus=SelectText(6) onblur=CheckDay() value="<%=szDayTmp%>" OnKeyDown="ChangeFocus(7)">日                
                                    </td>
                                    <td align=right bgcolor=#C9E0F8 height="2">航　線：</td>
                                    <td align=left bgcolor=#C9E0F8 height="2">                                        
                                        <select size="1" name="VesselLine" OnKeyDown="ChangeFocus(8)">             
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
                            <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="移動" name = "Move" OnKeyDown="MoveByKeyPress()" OnMouseUp="OnMove()">
                            <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="查詢" name = "Search" OnKeyDown="SearchByKeyPress()" OnMouseUp="OnSearch()">
                            <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="重置" name = "Reset" OnKeyDown="ResetByKeyPress()" OnMouseUp="OnReset()">
                        </div>
                    </td>
                    <td width=1><img height=38 src="../image/box5.gif" width=20></td>
                </tr>
            </tbody>
        </table> 
    </td>
</tr>
<%
    dim szOriginalVesselListID, szFormID
    szOriginalVesselListID = request ("OriginalVesselListID")
    szFormID = request ("FormID")
    
    response.write "<input type=""hidden"" name=OriginalVesselListID value=" + szOriginalVesselListID + ">"
    response.write "<input type=""hidden"" name=FormID value=" + szFormID + ">"
    
    
    dim nDataCounter, szSN(50), i, szElementName, szStr, szStr2
    
    nDataCounter = request ("DataCounter")
    if nDataCounter = "" then
        nDataCounter = 0
    end if
    
    response.write "<input type=""hidden"" name=DataCounter value=" + CStr(nDataCounter) + ">"
    
    for i = 0 to nDataCounter-1
        szElementName = "SN_" + CStr(i)
        szSN(i) = request(szElementName)
        response.write "<input type=""hidden"" name=" + szElementName + " value=" + CStr(szSN(i)) + ">"
        szStr2 = szStr2 + "&" + szElementName + "=" + CStr(szSN(i))
    next
    
    szStr = "&DataCounter=" + CStr(nDataCounter) + "&OriginalVesselListID=" + szOriginalVesselListID + "&FormID=" + szFormID + szStr2    
    
%>

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

<input type="hidden" name="VesselID" value=<%=szVesselID%>>

<A name=D></A>
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
        <table bgcolor=#C9E0F8 border=0 bordercolor=#000000 cellspacing=1 width=100% align="center">
            <tr bgcolor="#3366cc" align="center"> 
                <td width="20%" align="left"><font color="#FFFFFF">船名</font></td>
                <td width="18%" align="left"><font color="#FFFFFF">船公司</font></td>                
                <td width="19%" align="left"><font color="#FFFFFF">航次</font></td>
                <td width="15%" align="left"><font color="#FFFFFF">船掛</font></td>                
                <td width="18%" align="left"><font color="#FFFFFF">結關日</font></td>
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
            %>
                <tr align="center"> 
                    <td width="20%" bgcolor=#C9E0F8 align="left"><a href="FFfun02d.asp?Status=Move&VesselID=<%= rs("ID") %><%=szStr%>"><%=rs("VesselName")%></a></td>
                    <td width="18%" bgcolor=#C9E0F8 align="left"><a href="FFfun02d.asp?Status=Move&VesselID=<%= rs("ID") %><%=szStr%>"><%=rs("Owner")%></a></td>
                    <td width="19%" bgcolor=#C9E0F8 align="left"><a href="FFfun02d.asp?Status=Move&VesselID=<%= rs("ID") %><%=szStr%>"><%= rs("VesselNo")%></a></td>
                    <td width="15%" bgcolor=#C9E0F8 align="left"><a href="FFfun02d.asp?Status=Move&VesselID=<%= rs("ID") %><%=szStr%>"><%= rs("CheckInID")%></a></td>
                    <td width="18%" bgcolor=#C9E0F8 align="left"><a href="FFfun02d.asp?Status=Move&VesselID=<%= rs("ID") %><%=szStr%>"><%= rs("Date")%></a></td>
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
                    <td width="18%" bgcolor=#C9E0F8 align="left"><a href="FFfun02d.asp?Status=Move&VesselID=<%= rs("ID") %><%=szStr%>"><%=szLineName%></a></td>
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
</table>   
</form>
</body>
</html>
