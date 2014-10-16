<!-- #include file = ../GlobalSet/conn.asp -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>航次資料刪除</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="VOfun04.js"></script>
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
    
    '21-Nov2004: 加入航線
    szVesselLine    = request ("VesselLine")    
    
    If Trim(szYear) <> "" and Trim(szMonth) <> "" and Trim(szDay) <> "" then 
        szDate = CStr(szYear) + "/" + CStr(szMonth) + "/" + CStr(szDay)
    else
        szDate = "0"
    end if
    
    if szStatus = "" then
        szStatus = "Add"    'Default value
        szID     = ""
    end if
    
    '序號
    dim nSerialNum, nSerialNumTmp
    nSerialNum = 1
    nSerialNumTmp = request("SerialNum")
    
    '===========ID自動編號================
    dim nIDLen, nMaxID
    
    sql = "select max(ID) as MaxID, count(*) as IDCount from VesselList"
    set rs = conn.execute(sql)
    
    if not rs.eof then
        if rs("IDCount") = 0 then
            szIDTmp = "00001"
        else
            
            nMaxID = rs("MaxID") + 1
            
            if nMaxID < 10 then
                szIDTmp = "0000"
            elseif nMaxID < 100 then
                szIDTmp = "000"
            elseif nMaxID < 1000 then
                szIDTmp = "00"
            elseif nMaxID < 10000 then
                szIDTmp = "0"
            else
                szIDTmp = ""
            end if
            
            szIDTmp = szIDTmp + CStr(nMaxID)
        end if
    end if
    
    '===========查詢欲修改的資料===========
    dim szIDTmp, szVesselNameTmp, szVesselNoTmp, szCheckInIDTmp, szDateTmp, szYearTmp, szMonthTmp, szDayTmp, szOwnerTmp
    dim szVesselLineTmp  
    dim sql2, rs2
    Dim fTotalVolume, fTotalVolumeChecked
    
    if szStatus = "Modify" then
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
            
            '==========查詢該船的總體積======    
            sql = "select sum(Volume) as TotalVolume from FreightForm where VesselID = '" + szID + "'"
            set rs = conn.execute(sql)
            if not rs.eof then
                fTotalVolume = rs("TotalVolume")
            end if
            
            '==========查詢該船已核定過的總體積======    
            sql = "select sum(Volume) as TotalVolume from FreightForm where IsChecked = true and VesselID = '" + szID + "'"
            set rs = conn.execute(sql)
            if not rs.eof then
                fTotalVolumeChecked = rs("TotalVolume")
            end if
    
            '清除參數
            szID          = ""
            szVesselName  = ""
            szOwner       = ""
            szVesselNo    = ""
            szCheckInID   = ""
            szVesselLine  = ""
        end if
    end if
    
    '看轉換的文字檔是否存在
    dim szFilePath, szFileNameFormat1, szFileNameFormat2, szFileNameFormat3, szFileNameFormat4, bFileExist
    szFilePath = Request.ServerVariables ("PATH_TRANSLATED")
    
    nPos = InStrRev (szFilePath, "\")
    
    szFilePath = Mid (szFilePath, 1, nPos)
    
    szFilePath = szFilePath + "TextFiles\"
    
    szFilePath = szFilePath & szIDTmp & "\"
    
    szFileNameFormat1 = szFilePath & szIDTmp & ".txt"
    
    szFileNameFormat2 = szFilePath & "ZACT.txt"
    szFileNameFormat3 = szFilePath & "DAI_1.txt"
    szFileNameFormat4 = szFilePath & "DAI_2.txt"
    
    Set objFileSystem = CreateObject ("Scripting.FileSystemObject")
    
    if objFileSystem.FileExists(szFileNameFormat1) or objFileSystem.FileExists(szFileNameFormat2) or objFileSystem.FileExists(szFileNameFormat3) or objFileSystem.FileExists(szFileNameFormat4) then
        bFileExist = 1
    else
        bFileExist = 0
    end if
   
%>

<form name="form" method="post" action="VOfun04b.asp" onsubmit=" javascript: return checkform();" OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                    <tr bgcolor=#3366cc > 
                        <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                        <td align=middle width="100%" ><div align="center"><font color="#FFFFFF">航次資料刪除</font></div></td>
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
                                        <!--             
                                        <td align=right bgcolor=#C9E0F8 height="2" width="17%">編　　號：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="31%"> 
                                            <input type="text" name="ID" size="10" maxlength="10" value="<%=szIDTmp%>" onfocus=SelectText(0) OnKeyDown="ChangeFocus(1)">
                                        </td>
                                        -->
                                        <td align=right bgcolor=#C9E0F8 height="2" width="17%">船　名：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="31%">
                                            <input type="text" name="VesselName" size="30" maxlength="22" value="<%=szVesselNameTmp%>" onfocus="SelectText(1)" OnKeyDown="ChangeFocus(2)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">                                            
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2" width="17%">航　　次：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="31%"> 
                                            <input type="text" name="VesselNo" size="8" maxlength="10" value="<%=szVesselNoTmp %>" onfocus=SelectText(2) OnKeyDown="ChangeFocus(3)"  onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>  
                                    </tr>
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2">船公司：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="Owner" size="8" maxlength="15" value="<%=szOwnerTmp%>" onfocus=SelectText(3) OnKeyDown="ChangeFocus(4)"  onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>    
                                        
                                        <td align=right bgcolor=#C9E0F8 height="2">船　　掛：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="CheckInID" size="8" maxlength="6" value="<%=szCheckInIDTmp %>" onfocus=SelectText(4) OnKeyDown="ChangeFocus(5)"  onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align=right bgcolor=#C9E0F8 height="2">結關日：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" name="Year" size="1" maxlength="4" value="<%=szYearTmp%>" onfocus=SelectText(5) OnKeyDown="ChangeFocus(6)"  onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">年
                                            <input type="text" name="Month" size="1" maxlength="2"  onfocus=SelectText(6) onblur=CheckMonth() value="<%=szMonthTmp%>" OnKeyDown="ChangeFocus(7)"  onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">月
                                            <input type="text" name="Day" size="1" maxlength="2"  onfocus=SelectText(7) onblur=CheckDay() value="<%=szDayTmp%>" OnKeyDown="ChangeFocus(8)"  onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">日                
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2">船總體積：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"><%=fTotalVolume%></td>
                                    </tr>
                                    <tr>                                    
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
                                        <td align=right bgcolor=#C9E0F8 height="2">船總體積﹝</td>
                                        <td align=left bgcolor=#C9E0F8 height="2">已核對﹞：<%=fTotalVolumeChecked%></td>
                                    </tr>
                                    <tr>                                    
                                    <%
                                        if bFileExist = 1 then
                                            response.write "<td align=right bgcolor=#C9E0F8 height=""2"">文字檔：</td>"
                                            response.write "<td align=left colspan=3 bgcolor=#C9E0F8 height=""2"">"
                                            response.write "<a href=" & szFileNameFormat1 & ">1. " & szIDTmp & ".txt</a>　　"
                                            response.write "<a href=" & szFileNameFormat2 & ">2. ZACT.txt</a>　　"
                                            response.write "<a href=" & szFileNameFormat3 & ">3. DAI_1.txt</a>　　"
                                            response.write "<a href=" & szFileNameFormat4 & ">4. DAI_2.txt</a>　　"
                                            response.write "</td>"
                                            response.write "</tr>"
                                            response.write "<tr><td bgcolor=#C9E0F8 ></td>"
                                            response.write "<td align=left colspan=3 bgcolor=#C9E0F8 height=""2"">(請按""右鍵->另存目標""來存取檔案)　"
                                        else
                                            response.write "<td align=right bgcolor=#C9E0F8 height=""2""></td>"
                                            response.write "<td align=left colspan=3 bgcolor=#C9E0F8 height=""2""></td>"
                                        end if
                                    %>  
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
                                <!--<input type="button" style="background-color:#C9E0F8; border-color:#C9E0F8; border-style: outset; border-width: 2" value="儲存" OnKeyDown="AddByKeyPress()" OnMouseUp="Add()"  onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">-->
                                <input type="button" style="background-color:#C9E0F8; border-color:#C9E0F8; border-style: outset; border-width: 2" value="查詢" OnKeyDown="SearchByKeyPress()" OnMouseUp="Search()"  onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                                <input type="button" style="background-color:#C9E0F8; border-color:#C9E0F8; border-style: outset; border-width: 2" value="刪除" OnKeyDown="DeleteByKeyPress()" OnMouseUp="Delete()"  onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                                <!--<input type="button" style="background-color:#C9E0F8; border-color:#C9E0F8; border-style: outset; border-width: 2" value="結關" OnKeyDown="CloseByKeyPress()" OnMouseUp="Close()"  onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">-->
                                <input type="button" style="background-color:#C9E0F8; border-color:#C9E0F8; border-style: outset; border-width: 2" value="重置" OnKeyDown="ResetByKeyPress()" OnMouseUp="Reset()"  onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
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

<input type="hidden" name="Status" value=<%=szStatus%>>
<!--<input type="hidden" name="IDToModify" value=<%=szIDTmp%>>-->
<input type="hidden" name="ID" value=<%=szIDTmp%>>

<%
    '查詢出底下的list
    set rs = nothing 
    sql = "select * from VesselList where VesselName like '%" + szVesselName + "%'"
    
    sql = sql + " and VesselNo like '%" + szVesselNo + "%' and CheckInID like '%" + szCheckInID + "%'"
    
    if szDate <> "0" then
        sql = sql + " and Date = #" + szDate + "#"
    end if
    
    if szVesselLine <> "0" and szVesselLine <> "" then
        sql = sql + " and VesselLine = '" + szVesselLine + "'"
    end if
    
    sql = sql + " order by Date DESC, DataCount DESC ,ID DESC"
    
    set rs = conn.execute(sql)
    
    dim sqltmp
    sqltmp = sql
%>

<A name=D></A>
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
        <table bgcolor=#f0ebdb border=0 bordercolor=#000000 cellspacing=1 width=100% align="center">
            <tr bgcolor="#3366cc" align="center"> 
                <!--<td width="5%" align="left"><font color="#FFFFFF">新增</font></td>
                <td width="5%" align="left"><font color="#FFFFFF">查詢</font></td>-->
                <td width="6%" align="left"><font color="#FFFFFF">編號</font></td>
                <td width="29%" align="left"><font color="#FFFFFF">船名</font></td>                
                <td width="12%" align="left"><font color="#FFFFFF">航次</font></td>
                <td width="11%" align="left"><font color="#FFFFFF">船公司</font></td>
                <td width="10%" align="left"><font color="#FFFFFF">船掛</font></td>
                <td width="13%" align="left"><font color="#FFFFFF">結關日</font></td>
                <td width="9%" align="left"><font color="#FFFFFF">航線</font></td>
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
                    <!--<td width="5%" bgcolor=#C9E0F8 align="left"><a href="..\FreightForm\FFfun01a.asp?Status=Add&VesselID=<%= rs("ID") %>&SerialNum=<%=nSerialNum%>">新增</a></td>
                    <td width="5%" bgcolor=#C9E0F8 align="left"><a href="..\FreightForm\FFfun02b.asp?Status=FirstLoad&VesselListID=<%= rs("ID") %>&SerialNum=<%=nSerialNum%>">查詢</a></td>-->
                    <td width="6%" bgcolor=#C9E0F8 align="left"><a href="VOfun04a.asp?Status=Modify&ID=<%= rs("ID") %>&SerialNum=<%=nSerialNum%>"><%=nSerialNum%></a></td>
                    <td width="29%" bgcolor=#C9E0F8 align="left"><a href="VOfun04a.asp?Status=Modify&ID=<%= rs("ID") %>&SerialNum=<%=nSerialNum%>"><%=rs("VesselName")%></a></td>
                    <td width="12%" bgcolor=#C9E0F8 align="left"><a href="VOfun04a.asp?Status=Modify&ID=<%= rs("ID") %>&SerialNum=<%=nSerialNum%>"><%= rs("VesselNo")%></a></td>
                    <td width="11%" bgcolor=#C9E0F8 align="left"><a href="VOfun04a.asp?Status=Modify&ID=<%= rs("ID") %>&SerialNum=<%=nSerialNum%>"><%= rs("Owner")%></a></td>
                    <td width="10%" bgcolor=#C9E0F8 align="left"><a href="VOfun04a.asp?Status=Modify&ID=<%= rs("ID") %>&SerialNum=<%=nSerialNum%>"><%= rs("CheckInID")%></a></td>
                    <td width="13%" bgcolor=#C9E0F8 align="left"><a href="VOfun04a.asp?Status=Modify&ID=<%= rs("ID") %>&SerialNum=<%=nSerialNum%>"><%= rs("Date")%></a></td>
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
                    
                    <td width="9%" bgcolor=#C9E0F8 align="left"><a href="VOfun04a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= szLineName%></a></td>
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

<input type="hidden" name="Sql" value="<%=sqltmp%>">  <!---->

<%
    rs.close
    conn.close 
    
    set rs=nothing
    set conn=nothing
%> 

</form>
</body>
</html>
