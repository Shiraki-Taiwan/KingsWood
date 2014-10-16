<!-- #include file = ../GlobalSet/conn.asp -->
<!--航次資料列印-->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>航次資料列印</title>
	<!-- #include file = ../head_bundle.html -->
	<link href="../GlobalSet/sg.css" rel="stylesheet" type="text/css" />
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@KAOINS9") < 0 then
		response.redirect "../Login/Login.asp"
	end if


    '===============接收參數===============
    dim szID, szVesselName, szVesselNo, szCheckInID, nYear, nMonth, nDay, szVesselLine
        
    szID            = request("ID")               '編號
    szVesselName    = request("VesselName")       '船名
    szVesselNo      = request("VesselNo")         '航次
    szCheckInID     = request("CheckInID")        '船掛
    nYear           = request("Year")             '結關年
    nMonth          = request("Month")            '結關月
    nDay            = request("Day")              '結關日
    szVesselLine    = request("VesselLine")       '航線
    
    if szVesselID = "0" then
        szVesselID = ""
    end if
    
    if (nYear) = "" then
        nYear = "0"
    end if
    
    if (nMonth) = "" then
        nMonth = "0"
    end if
    
    if (nDay) = "" then
        nDay = "0"
    end if
    
    '連接年月日成一字串
    dim szDate
    If Trim(nYear) <> "0" and Trim(nMonth) <> "0" and Trim(nDay) <> "0" then 
        szDate = CStr(int(nYear)) + "/" + CStr(int(nMonth)) + "/" + CStr(int(nDay)) 
    else
        szDate = "0"
    end if
    
    

    '===============查詢資料庫===============  
    dim sql, rs 
    
    set rs = nothing 
    sql = "select * from VesselList where VesselName like '%" + szVesselName + "%'"
    sql = sql + " and VesselNo like '%" + szVesselNo + "%' and CheckInID like '%" + szCheckInID + "%'"
    
    if szDate <> "0" then
        sql = sql + " and Date = #" + szDate + "#"
    end if
    
    if szVesselLine <> "0" then
        sql = sql + " and VesselLine='" + szVesselLine + "'"
    end if
    
    
    sql = sql + " order by Date DESC, DataCount DESC ,ID DESC"

    set rs = conn.execute(sql) 
    
%>

<!------------------------列表----------------------------------->
<form name="form" OnKeyDown="CheckHotKey()">
<table width="100%" height=""100%"">
    <tr> 
        <td > 
        <table border=0 width="70%" cellspacing=1 height=""100%"" align="center">
            <tr>
                <td colspan=6 align="center"><font size=5>航次資料列表</font></td>
            </tr>
            <tr>
                <td colspan=6 align="center"><hr></td>
            </tr>
            <tr valign="center"> 
                <td width="15%" align="left">船公司</td>
                <td width="10%" align="left">船名</td>                
                <td width="10%" align="left">航次</td>
                <td width="15%" align="left">船掛</td>
                <td width="15%" align="left">結關日</td>
                <td width="35%" align="left">航線</td>
            </tr>
            <tr>
                <td colspan=6 align="center"><hr></td>
            </tr>
       
            <%
                while not rs.eof
            %>
                <tr align="center"> 
                    <td align="left"><%= rs("Owner")%></td>
                    <td align="left"><%= rs("VesselName")%></td>                    
                    <td align="left"><%= rs("VesselNo")%></td>
                    <td align="left"><%= rs("CheckInID")%></td>
                    <td align="left"><%= rs("Date")%></td>
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
                    <td align="left"><%=szLineName%></td>
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