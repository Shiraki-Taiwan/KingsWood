<!-- #include file = ../GlobalSet/conn.asp -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>航次資料列印查詢</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="VOfun02.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<%
    dim sql, rs
    set rs = nothing
    
    '===========接參數===========
    dim szID, szVesselName, szVesselNo, szCheckInID, szStatus, szYear, szMonth, szDay, szDate, szOwner, szVesselLine
    szID            = request ("ID")
    szVesselName    = request ("VesselName")
    szOwner         = request ("Owner")
    szVesselNo      = request ("VesselNo")
    szCheckInID     = request ("CheckInID")
    szStatus        = request ("Status")
    szYear          = request ("Year")
    szMonth         = request ("Month")
    szDay           = request ("Day")  
    '21-Nov2004:加入航線
    szVesselLine   = request ("VesselLine")  
    
    If Trim(szYear) <> "" and Trim(szMonth) <> "" and Trim(szDay) <> "" then 
        szDate = CStr(szYear) + "/" + CStr(szMonth) + "/" + CStr(szDay)
    else
        szDate = "0"
    end if
    
    
    '===========查詢資料===========
    dim szIDTmp, szVesselNameTmp, szVesselNoTmp, szCheckInIDTmp, szDateTmp, szYearTmp, szMonthTmp, szDayTmp, szOwnerTmp
    dim szVesselLineTmp  
    dim sql2, rs2
    Dim fTotalVolume, fTotalVolumeChecked
    
    if szStatus = "Search" then
        sql = "select * from VesselList where VesselList.ID = '" + szID + "'"
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
            VesselLine    = ""
        end if
    end if
    
%>

<form name="form" method="post" action="VOfun02b.asp" OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                    <tr bgcolor=#3366cc > 
                        <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                        <td align=middle width="100%" ><div align="center"><font color="#FFFFFF">航次資料列印查詢</font></div></td>
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
                                        <!--             
                                        <td align=right bgcolor=#C9E0F8 height="2" width="17%">編　　號：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="31%"> 
                                            <input type="text" name="ID" size="10" maxlength="10" value="<%=szIDTmp%>" onfocus=SelectText(0) OnKeyDown="ChangeFocus(1)">
                                        </td>
                                        -->
                                        <td align=right bgcolor=#C9E0F8 height="2" width="17%">船　名：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="31%">
                                            <input type="text" name="VesselName" size="30" maxlength="22" value="<%=szVesselNameTmp%>" onfocus=SelectText(0) OnKeyDown="ChangeFocus(1)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">                                            
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2" width="17%">航　次：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="31%"> 
                                            <input type="text" name="VesselNo" size="8" maxlength="10" value="<%=szVesselNoTmp %>" onfocus=SelectText(1) OnKeyDown="ChangeFocus(2)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>      
                                    </tr>
        
        
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2">船公司：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="Owner" size="8" maxlength="15" value="<%=szOwnerTmp%>" onfocus=SelectText(2) OnKeyDown="ChangeFocus(3)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2">船　掛：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="CheckInID" size="8" maxlength="20" value="<%=szCheckInIDTmp %>" onfocus=SelectText(3) OnKeyDown="ChangeFocus(4)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align=right bgcolor=#C9E0F8 height="2">結關日：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="Year" size="1" maxlength="4" value="<%=szYearTmp%>" onfocus=SelectText(4) OnKeyDown="ChangeFocus(5)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">年
                                            <input type="text" name="Month" size="1" maxlength="2"  onfocus=SelectText(5) onblur=CheckMonth() value="<%=szMonthTmp%>" OnKeyDown="ChangeFocus(6)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">月
                                            <input type="text" name="Day" size="1" maxlength="2"  onfocus=SelectText(6) onblur=CheckDay() value="<%=szDayTmp%>" OnKeyDown="ChangeFocus(7)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">日                
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
<br>
<input type="hidden" name="ID" value=<%=szIDTmp%>>

<%
    '查詢出底下的list
    set rs = nothing 
    sql = "select * from VesselList where VesselName like '%" + szVesselName + "%'"
    
    sql = sql + " and VesselNo like '%" + szVesselNo + "%' and CheckInID like '%" + szCheckInID + "%'"
    
    if szDate <> "0" then
        sql = sql + " and Date = #" + szDate + "#"
    end if
    
    sql = sql + " order by Date DESC, DataCount DESC ,ID DESC"
    
    set rs = conn.execute(sql)
   
%>

<A name=D></A>
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
        <table bgcolor=#f0ebdb border=0 bordercolor=#000000 cellspacing=1 width=100% align="center">
            <tr bgcolor="#3366cc" align="center"> 
                <td width="30%" align="left"><font color="#FFFFFF">船名</font></td>
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
            %>
                <tr align="center"> 
                    <td width="30%" bgcolor=#C9E0F8 align="left"><a href="VOfun02a.asp?Status=Search&ID=<%= rs("ID") %>"><%=rs("VesselName")%></a></td>
                    <td width="15%" bgcolor=#C9E0F8 align="left"><a href="VOfun02a.asp?Status=Search&ID=<%= rs("ID") %>"><%= rs("VesselNo")%></a></td>
                    <td width="15%" bgcolor=#C9E0F8 align="left"><a href="VOfun02a.asp?Status=Search&ID=<%= rs("ID") %>"><%= rs("Owner")%></a></td>
                    <td width="15%" bgcolor=#C9E0F8 align="left"><a href="VOfun02a.asp?Status=Search&ID=<%= rs("ID") %>"><%= rs("CheckInID")%></a></td>
                    <td width="15%" bgcolor=#C9E0F8 align="left"><a href="VOfun02a.asp?Status=Search&ID=<%= rs("ID") %>"><%= rs("Date")%></a></td>
                    <%
                        dim szLineName
                        if rs("VesselLine") <> "0" then
                            sql = "select Name from VesselLine where ID = '" + rs("VesselLine") + "'"                        
                            
                            set rs1 = conn.execute(sql)
                            
                            if not rs1.eof then
                                szLineName = rs1("Name")
                            end if
                            
                            rs1.close
                            set rs1 = nothing
                        else
                            szLineName = ""
                        end if
                    %>
                    <td width="10%" bgcolor=#C9E0F8 align="left"><a href="FFfun02b.asp?Status=FirstLoad&VesselListID=<%= rs("ID") %>"><%= szLineName%></a></td>
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
