<!-- #include file = ../GlobalSet/conn.asp -->
<!-- #include file = ../GlobalSet/ShareFun.asp -->
<!-- #include file = FSShare.asp -->
<!--攬貨報告書列印-->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>攬貨報告書列印</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="../GlobalSet/dmenu.js"></script>
	<script type="text/javascript" src="../GlobalSet/ShareFun.js"></script>
	<script type="text/javascript" src="FSfun01e.js"></script>
	<link rel="stylesheet" type="text/css" href="../GlobalSet/main.css" media="screen, print" />
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="Print();">
<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@KEEINS5") < 0 then
		response.redirect "../Login/Login.asp"
	end if

    '===============接收參數===============
    dim szVesselListID, szFreightOwnerID, szStatus, szFormID, szCompanyID
    dim szSelectType, szReportType
    dim szVesselLine, szContainer
    
    szVesselListID   = request("VesselListID")     '航次
    szFreightOwnerID = request ("FreightOwnerID")  '攬貨商
    szFormID = request ("FormID")                  '起始單號
    szCompanyID = request("CompanyID")             '公司名稱
    szStatus = request ("Status")
    szSelectType = request ("SelectType")          '依攬貨商或單號...
    szReportType = request ("ReportType")          '報表格式
    
    szVesselLine = request ("VesselLine")           '航線
    szContainer = request ("Container")             '貨櫃場
    
    dim szHotKey
    
    szHotKey = request ("HotKey")		'是否使用快速鍵
    
    '===========Parse 單號/攬貨商===================
    dim szSubStr
    
    if szSelectType = 1 then    '依單號
        szSubStr = szFormID        
    else                        '依攬貨商
        szSubStr = szFreightOwnerID
    end if
    
    ParseFormID()
        
    '===============查詢資料庫===============  
    dim sql, rs     
    
    '查詢公司資料
    dim szCmpChtName, szCmpEngName, szCmpFaxNo, szCmpPhoneNo, szContactName, szMobile
    sql = "select * from CompanyInfo where ID = '" + szCompanyID +"'"
    set rs = conn.execute(sql)
    if not rs.eof then
        szCmpChtName = rs("ChineseName")
        szCmpEngName = rs("EnglishName")
        szCmpFaxNo   = rs("FaxNo_1")
        szCmpPhoneNo = rs("Phone_1")
        szContactName= rs("Contact_1")
        szMobile     = rs("Mobile_1")
    end if
    
    '查航線
    dim szVesselLineName 
    set rs = nothing   
    sql = "select * from VesselLine where ID = '" + szVesselLine + "'"
    
    set rs = conn.execute(sql)
    if not rs.eof then
        szVesselLineName = rs("Name")
    end if
     
    '查詢航次的資料
    dim szVesselName, szVesselList, szVesselDate, szVesselOwner
    set rs = nothing
    sql = "select * from VesselList where ID = '" + szVesselListID + "'"
    
    set rs = conn.execute(sql)
    
    if not rs.eof then
        szVesselName = rs("VesselName")
        szVesselList = rs("VesselNo")
        szVesselOwner= rs("Owner")
        szVesselDate = rs("Date")
    end if
    
   
%>

<!------------------------列表----------------------------------->
<form name="form" method="post" action="FSfun01d.asp" onload="Print()">
   
<%
    dim nNeededVolume
        
    dim nFoundCount
    '===================詳細尺寸資料==============================
    if szReportType = 2 then            

%>
                <table border=0 cellspacing=1 width="90%" align="center">
                    <tr>
                        <td colspan=4 align="center"><font size=5><span style="letter-spacing: 10pt"><%=szCmpChtName%></span></font></td>
                    </tr>
                    <tr>
                        <td colspan=4 align="center"><font size=5><%=szCmpEngName%></font></td>
                    </tr>
                    <tr>
                        <td colspan=4 align="center"><br></td>
                    </tr>                
                    <tr>
                        <td colspan=4 align="center"><font size=5>度量資料</font></td>
                    </tr>
                    <tr>
                        <td colspan=4 align="center"><hr size="0"></td>
                    </tr>
                    <tr valign="center">
                        <td width="1%" align="left"></td>
                        <td width="20%" align="left"><font size=3>船名：<%=szVesselName%></font></td> 
                        <td width="20%" align="left"><font size=3>航次：<%=szVesselList%></font></td>
                        <td width="69%" align="left"><font size=3>結關日期：<%=szVesselDate%></font></td>
                    </tr>
               </table>
               <table border=0 cellspacing=0 width="90%" align="center">
                    <tr align="center">
                        <td align="right" width="7%"><font size=3>倉位</font></td> 
                        <td align="right" width="7%"><font size=3>單號</font></td>
                        <!--<td align="right" width="7%"><font size=3>倉位</font></td>-->
                        <td align="right" width="7%"><font size=3>板數</font></td>
                        <td align="right" width="7%"><font size=3>堆量</font></td>
                        <td align="right" width="8%"><font size=3>件數</font></td>
                        <td align="right" width="7%"><font size=3>包裝</font></td>
                        <td align="right" width="9%"><font size=3>長</font></td>
                        <td align="right" width="9%"><font size=3>寬</font></td>
                        <td align="right" width="9%"><font size=3>高</font></td>
                        <td align="right" width="10%"><font size=3>體積</font></td>                    
                        <td align="right" width="10%"><font size=3>單重</font></td>
                        <td align="right" width="10%"><font size=3>總重</font></td>
                    </tr>
                    
                    <tr align="center"> 
                        <td align="right"><font size=3>====</font></td>
                        <td align="right"><font size=3>====</font></td>
                        <!--<td align="right"><font size=3>====</font></td>-->
                        <td align="right"><font size=3>====</font></td>
                        <td align="right"><font size=3>====</font></td>
                        <td align="right"><font size=3>====</font></td>
                        <td align="right"><font size=3>====</font></td>
                        <td align="right"><font size=3>=====</font></td>
                        <td align="right"><font size=3>=====</font></td>
                        <td align="right"><font size=3>=====</font></td>
                        <td align="right"><font size=3>=====</font></td>
                        <td align="right"><font size=3>=====</font></td>
                        <td align="right"><font size=3>=====</font></td>
                    </tr>
                    
           <%
                nFoundCount = 0
                dim nPrevFormID, szBgColor
                nTotalPiece = 0
                fVolume = 0
                bIsFirst = TRUE            
                
                dim szStartIDTmpFmt, szEndIDTmpFmt
                
                for i = 0 to nCount                     
                    
                    '查倉單資料
                    set rs = nothing
                    
                    sql = "select FreightForm.* from FreightForm where VesselID = '" + szVesselListID + "'"
                    
                    if szStartIDTmp(i) <> "" and szEndIDTmp(i) <> "" then                    
                        '27-Apr2005: 把單號format成4位
                        szStartIDTmpFmt = szStartIDTmp(i)
                        szEndIDTmpFmt = szEndIDTmp(i)
                        
                        szStartIDTmpFmt = FormatString (szStartIDTmpFmt, 4)
                        szEndIDTmpFmt = FormatString (szEndIDTmpFmt, 4)
                        
                        sql = sql + " and ID >= '" + szStartIDTmpFmt + "' and ID <= '" + szEndIDTmpFmt + "'"
                    end if
                    
                    '20-Jul2005: 以輸入順序顯示 (The priority is higher than PageNo)
                    sql = sql + " order by ID, SN, PageNo"
if (szReportType = 2 AND szSelectType = 2) then
	sql = "select FreightForm.* from FreightForm , FormToOwner where VesselID = '" + szVesselListID + "' and FreightForm.ID = FormToOwner.FormID and FormToOwner.OwnerID ='"+ FormatString(szFreightOwnerID, 3) +"' and FormToOwner.VesselLine='" + szVesselLine + "' order by ID"                                                       	
end if
'response.write "SQL="&sql&"<BR>"

                  
                    set rs = conn.execute(sql)
                    while not rs.eof
                        nFoundCount = nFoundCount + 1
                        
                        'If rs("Length") > 600 OR rs("Width") > 600 OR rs("Height") > 226 OR rs("Volume") > 3.8 OR rs("Volume") < 0.1 Then
                        '    szBgColor = "#FFFFA0"
                        'else
                            szBgColor = "#FFFFFF"   
                        'end if
                        
                        if bIsFirst = TRUE then
                            nPrevFormID = rs("ID")
                            bIsFirst = FALSE
                        elseif nPrevFormID <> rs("ID") then
                            nPrevFormID = rs("ID")
            %>
                            <tr align="center"> 
                                <td align="right"><font size=3>====</font></td>
                                <td align="right"><font size=3>====</font></td>
                                <!--<td align="right"><font size=3>====</font></td>-->
                                <td align="right"><font size=3>====</font></td>
                                <td align="right"><font size=3>====</font></td>
                                <td align="right"><font size=3>====</font></td>
                                <td align="right"><font size=3>====</font></td>
                                <td align="right"><font size=3>=====</font></td>
                                <td align="right"><font size=3>=====</font></td>
                                <td align="right"><font size=3>=====</font></td>
                                <td align="right"><font size=3>=====</font></td>
                                <td align="right"><font size=3>=====</font></td>
                                <td align="right"><font size=3>=====</font></td>
                            </tr>                
                            <tr align="center"> 
                                <td align="center" colspan=3><font size=3>TOTAL：</font></td>
                                <!--<td align="right"></td>-->
                                <td align="right"></td>
                                <td align="right"><font size=3><%=nTotalPiece%></font></td>
                                <td align="right"></td>
                                <td align="right"></td>                    
                                <td align="right"></td>
                                <td align="right"></td>
                                <%
                                    if nNeededVolume <> 0 then
                                %>
                                        <td align="right"><font size=3><%=FormatNumber(nNeededVolume, 2)%></font></td>
                                <%
                                    else
                                %>
                                        <td align="right"><font size=3><%=FormatNumber(fTotalVolume, 2)%></font></td>
                                <%
                                    end if
                                %>
                                <td align="right"></td>
                                <td align="right"></td>
                            </tr>
                            <tr>
                                <td colspan=13><br></td>
                            </tr>
            <%
                            nTotalPiece = 0
                            fTotalVolume = 0
                        end if 
            %>                        
                        
                        <tr align="center">
                            <td align="right" style="background-color:<%=szBgColor%>"><font size=3><%=rs("Storehouse")%></font></td> 
                            <td align="right" style="background-color:<%=szBgColor%>"><font size=3><%=rs("ID")%></font></td>
                            <!--<td align="right" style="background-color:<%=szBgColor%>"><font size=3><%=rs("Storehouse")%></font></td>-->
                            
                            <%
                                '20-Nov2004:如果板數為0, 則不要show
                                dim BoardTmp                                
                                if rs("Board") = 0 then
                                    BoardTmp = ""
                                else
                                    BoardTmp = rs("Board")
                                end if
                            %>
                            <td align="right" style="background-color:<%=szBgColor%>"><font size=3><%=BoardTmp%></font></td>
                            <td align="right" style="background-color:<%=szBgColor%>"><font size=3>
                            <%
                                if rs("IsPL") = TRUE then
                                    response.write "堆量"
                                end if
                            %>
                            </font></td>
                            <td align="right" style="background-color:<%=szBgColor%>"><font size=3><%=rs("Piece")%></font></td>
                            <%
                                dim szPackageStyle, sql2, rs2
                                szPackageStyle = ""
                                if rs("PackageStyleID") <> "0" then
                                    sql2 = "select * from PackageStyle where ID = '" + CStr(rs("PackageStyleID")) + "'"
                                    set rs2 = conn.execute(sql2)
                                    if not rs2.eof then
                                        szPackageStyle = rs2("Name")                                
                                    end if
                                    
                                    rs2.close
                                    set rs2 = nothing
                                end if
                            %>
                            <td align="right" style="background-color:<%=szBgColor%>"><font size=3><%=szPackageStyle%></font></td>
                            <td align="right" style="background-color:<%=szBgColor%>"><font size=3><%=rs("Length")%></font></td>
                            <td align="right" style="background-color:<%=szBgColor%>"><font size=3><%=rs("Width")%></font></td>
                            <td align="right" style="background-color:<%=szBgColor%>"><font size=3><%=rs("Height")%></font></td>
                            <td align="right" style="background-color:<%=szBgColor%>"><font size=3><%=FormatNumber(rs("Volume"), 2)%></font></td>
                            <%
                                '20-Nov2004:如果單重為0, 則不要show
                                dim WeightTmp                                
                                if rs("Weight") = 0 then
                                    WeightTmp = ""
                                else
                                    WeightTmp = MyFormatNumber(rs("Weight"), 1)
                                end if
                            %>
                            <td align="right" style="background-color:<%=szBgColor%>"><font size=3><%=WeightTmp%></font></td>
                            <%
                                '20-Nov2004:如果總重為0, 則不要show
                                dim TotalWeightTmp                                
                                if rs("TotalWeight") = 0 then
                                    TotalWeightTmp = ""
                                else
                                    TotalWeightTmp = MyFormatNumber(rs("TotalWeight"), 1)
                                end if
                            %>
                            <td align="right" style="background-color:<%=szBgColor%>"><font size=3><%=TotalWeightTmp%></font></td>
                        </tr>                
            <%
                        nTotalPiece = nTotalPiece + rs("Piece")
                        fTotalVolume = fTotalVolume + rs("Volume")
                        
                        nNeededVolume = rs("NeededForestry")
                        
                        rs.movenext
                        
                    wend
                    
                next
            %>
                <tr align="center"> 
                    <td align="right"><font size=3>====</font></td>
                    <td align="right"><font size=3>====</font></td>
                    <!--<td align="right"><font size=3>====</font></td>-->
                    <td align="right"><font size=3>====</font></td>
                    <td align="right"><font size=3>====</font></td>
                    <td align="right"><font size=3>====</font></td>
                    <td align="right"><font size=3>====</font></td>
                    <td align="right"><font size=3>=====</font></td>
                    <td align="right"><font size=3>=====</font></td>
                    <td align="right"><font size=3>=====</font></td>
                    <td align="right"><font size=3>=====</font></td>
                    <td align="right"><font size=3>=====</font></td>
                    <td align="right"><font size=3>=====</font></td>
                </tr>         
                <tr align="center"> 
                    <td align="center" colspan=3><font size=3>TOTAL：</font></td>
                    <!--<td align="right"></td>-->
                    <td align="right"></td>
                    <td align="right"><font size=3><%=nTotalPiece%></font></td>
                    <td align="right"></td>
                    <td align="right"></td>                    
                    <td align="right"></td>
                    <td align="right"></td>
                <%
                    if nNeededVolume <> 0 then
                %>
                        <td align="right"><font size=3><%=FormatNumber(nNeededVolume, 2)%></font></td>
                <%
                    else
                %>
                        <td align="right"><font size=3><%=FormatNumber(fTotalVolume, 2)%></font></td>
                <%
                    end if
                %>
                    <td align="right"></td>
                    <td align="right"></td>
                </tr>
                <tr>
                    <td colspan=13><br></td>
                </tr>
                    
            </table>
      
 
<%
    '===================總表格式==============================
    elseif szReportType = 0 or szReportType = 1 then
        nFoundCount = 0
        
        dim nTotalPiece, nTotalVolume, nTotalWeight, j
        nNeededVolume = 0
        
        '************依單號************
        if szSelectType = 1 then
%>
        <table border=0 cellspacing=1 width="100%" height="100%" align="center">
            <tr>
                <td valign="top">
                    <table border=0 cellspacing=1 width="100%" align="center">
                        <tr>
                            <td colspan=4 align="center"><font size=5><span style="letter-spacing: 10pt"><%=szCmpChtName%></span></font></td>
                        </tr>
                        <tr>
                            <td colspan=4 align="center"><font size=5><%=szCmpEngName%></font></td>
                        </tr>
                        <tr>
                            <td colspan=4 align="center"><br></td>
                        </tr>
                        <tr>
                            <td colspan=4 align="center"><font size=5>LIST OF CARGO MEASURE & WEIGHT</font></td>
                        </tr>
                        <tr>
                            <td colspan=4 align="center">
                            <%
                                for i = 0 to 80
                                    response.write "-"
                                next
                            %>
                            </td>
                        </tr>
                        <tr valign="center"> 
                            <td width="30%" align="right">船名&航次 ：</td>
                            <td width="24%" align="left"><%=szVesselName%> - <%=szVesselList%></td>
                            <td width="16%" align="right">結關日期：</td>
                            <td width="30%" align="left"><%=szVesselDate%></td>
                        </tr>
                        <tr valign="center"> 
                            <td align="right">船　公　司：</td>
                            <td align="left"><%=szVesselOwner%></td>
                            <td align="right">電　　話：</td>
                            <td align="left"><%=szCmpPhoneNo%> <%=szContactName%></td>
                        </tr>
                        <tr valign="center"> 
                            <td align="right">貨　櫃　場：</td>
                            <td align="left"><%=szContainer%></td>
                            <td align="right">傳　　真：</td>
                            <td align="left"><%=szCmpFaxNo%> <%=szContactName%></td>
                        </tr>
                        <tr valign="center"> 
                            <td align="right">航線 / 港口：</td>
                            <td align="left"><%=szVesselLineName%></td>
                            <td align="right">手　　機：</td>
                            <td align="left"><%=szMobile%> <%=szContactName%></td>
                        </tr>
                        <tr>
                            <td colspan=4 align="center">
                            <%
                                for i = 0 to 80
                                    response.write "-"
                                next
                            %>
                            </td>
                        </tr>
                   </table>
                 
                   <table border=0 cellspacing=1 width="75%" align="center">
                        <tr align="center"> 
                            <td align="right" width="1%"></td>
                            <td align="right" width="10%"></td>
                            <td align="right" width="16%">核對</td>
                            <td align="right" width="16%">單號</td>
                            <td align="right" width="16%">件數</td>
                            <td align="right" width="16%">體積</td>
                            <td align="right" width="16%">重量</td>
                            <td align="right" width="9%"></td>
                        </tr>
                        
                        <tr align="center"> 
                            <td align="right"></td>
                            <td align="right">=====</td>
                            <td align="right">=====</td>
                            <td align="right">=====</td>
                            <td align="right">=====</td>
                            <td align="right">=====</td>
                            <td align="right">=====</td> 
                            <td align="right"></td>                   
                        </tr>
                        
                   <%                        
                        nTotalPiece = 0
                        nTotalVolume = 0
                        nTotalWeight = 0
                        
                        dim i
                        
                        for i = 0 to nCount
                                 
                            '查倉單資料
                            set rs = nothing                            
                            
                            sql = "select ID, IsChecked, NeededForestry, sum(Piece) as TotalPiece, sum(Volume) as TotalVolume, sum(TotalWeight) as TotalWeightSum"
                            sql = sql + " from FreightForm where VesselID = '" + szVesselListID + "' "
                            
                            if szStartIDTmp(i) <> "" and szEndIDTmp(i) <> "" then
                                '27-Apr2005: 把單號format成4位
		                        szStartIDTmpFmt = szStartIDTmp(i)
		                        szEndIDTmpFmt = szEndIDTmp(i)
		                        
		                        szStartIDTmpFmt = FormatString (szStartIDTmpFmt, 4)
		                        szEndIDTmpFmt = FormatString (szEndIDTmpFmt, 4)
		                        
		                        sql = sql + " and ID >= '" + szStartIDTmpFmt + "' and ID <= '" + szEndIDTmpFmt + "'"
                            end if
                               
                            if szReportType = 1 then    '僅核對過的
                                sql = sql + " and IsChecked = TRUE"
                            end if
                            
                            sql = sql + " group by ID, IsChecked, NeededForestry order by ID"
                            
                            set rs = conn.execute(sql) 
                            
                            while not rs.eof
                                nNeededVolume = rs("NeededForestry")
                                nFoundCount = nFoundCount + 1
                    %>
                            <tr align="center"> 
                                <td align="right"></td>
                                <td align="right"  width="10%">
                                <%
                                    sql2 = "select HasDelivered from FreightForm where ID = '" + CStr(rs("ID")) + "' and  VesselID = '" + szVesselListID + "'"
                                    sql2 = sql2 + " and HasDelivered = false"
                                    set rs2 = conn.execute(sql2) 
                                    
                                    if rs2.eof then
                                        response.write ""
                                    else
                                        response.write "新增"
                                    end if
                                    
                                    rs2.close
                                    set rs2 = nothing
                                %>
                                </td>
                                <td align="right">
                                <%
                                    if rs("IsChecked") = FALSE then
                                        response.write "未核對"
                                    end if
                                %>
                                </td>
                                <td align="right"><%=rs("ID")%></td>
                                <td align="right"><%=rs("TotalPiece")%></td>
                                <td align="right">
                                <%
                                    if nNeededVolume <> 0 then
                                        response.write FormatNumber(nNeededVolume, 2)
                                    else
                                        response.write FormatNumber(rs("TotalVolume"), 2)
                                    end if
                                %>
                                </td>
                                <td align="right"><%=MyFormatNumber(rs("TotalWeightSum"),1)%></td>
                            </tr>                
                    <%
                                nTotalPiece = nTotalPiece + rs("TotalPiece")
                                
                                if nNeededVolume <> 0 then
                                    nTotalVolume = nTotalVolume + nNeededVolume
                                else
                                    nTotalVolume = nTotalVolume + rs("TotalVolume")
                                end if
                                
                                nTotalWeight = nTotalWeight + rs("TotalWeightSum")
                            
                                '9-May2005: 註記成"已輸出過"
                                sql2 = "update FreightForm set HasDelivered = 1  where ID = '" + CStr(rs("ID")) + "' and  VesselID = '" + szVesselListID + "'"
                                conn.execute(sql2) 
                            
                                rs.movenext
                            wend
                            
                        next
                    %>
                    
                        <tr align="center"> 
                            <td align="right"></td>
                            <td align="right">=====</td>
                            <td align="right">=====</td>
                            <td align="right">=====</td>
                            <td align="right">=====</td>
                            <td align="right">=====</td>
                            <td align="right">=====</td>
                            <td align="right"></td>
                            
                        </tr>
                        <tr align="center"> 
                            <td align="right"></td>
                            <td align="right"></td>
                            <td align="right"></td>
                            <td align="right">TOTAL：</td>
                            <td align="right"><%=nTotalPiece%></td>
                            <td align="right"><%=FormatNumber(nTotalVolume, 2)%></td>
                            <td align="right"><%=MyFormatNumber(nTotalWeight,1)%></td>
                            <td align="right"></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>

    <%
        elseif szSelectType = 2 then    '依攬貨商
            dim nFreightOwnerID, iCurIndex
            dim bNeedSend
            
            dim szRemark, nFreightOwnerCount
            nFreightOwnerCount = 1
            
            nFoundCount = 0       
            iCurIndex = -1     
            
            for i = 0 to nCount
                if IsNumeric(szStartIDTmp(i)) then
                    for nFreightOwnerID = CLng(szStartIDTmp(i)) to CLng(szEndIDTmp(i))
                        
                        szFreightOwnerID = CStr(nFreightOwnerID)
                        
                        '查倉單資料
                        '************依攬貨商************
                        set rs = nothing                            
                        
                        sql = "select ID, IsChecked, NeededForestry, HasDelivered, sum(Piece) as TotalPiece, sum(Volume) as TotalVolume, sum(TotalWeight) as TotalWeightSum"
                        sql = sql + " from FreightForm, FormToOwner where VesselID = '" + szVesselListID + "' "
                        sql = sql + " and FreightForm.ID = FormToOwner.FormID and CLng(FormToOwner.OwnerID) = '" + szFreightOwnerID + "'"
                        sql = sql + " and FormToOwner.VesselLine='" + szVesselLine + "'"
                        
                        if szReportType = 1 then    '僅核對過的
                            sql = sql + " and IsChecked = TRUE"
                        end if
                        
                        if szHotKey = "MailNew" then
                        	sql = sql + " and HasDelivered = false"
                        end if
                
                        sql = sql + " group by ID, IsChecked, NeededForestry, HasDelivered"
                        
                        if szHotKey = "MailAll" then
                        	sql = sql + " order by HasDelivered desc, ID"
                        else
                        	sql = sql + " order by ID"
                        end if
                
                        set rs = conn.execute(sql)
                        
                        bNeedSend = FALSE           'default
                        if szHotKey = "MailAll" then
                            if not rs.eof then
                                while not rs.eof and not bNeedSend
                                    if rs("HasDelivered") = FALSE then
                                        bNeedSend = TRUE
                                    end if
                                    
                                    rs.movenext
                                wend
                                
                                rs.movefirst
                            end if
                        else
                            bNeedSend = TRUE
                        end if
                        
                        if not rs.eof and bNeedSend then                            
                            nFoundCount = nFoundCount + 1
                            
                            '查攬貨商資料
                            dim szPhone, szFaxNo, szContact, szFreightOwnerName, szMailAddr, sql1, rs1
                            sql1 = "select * from FreightOwner where CLng(ID) = '" + szFreightOwnerID + "'"  
                            set rs1 = conn.execute(sql1) 
                            
                            if not rs1.eof then
                                szPhone = rs1("Phone_1")
                                szFaxNo = rs1("FaxNo_1")
                                szContact = rs1("Contact_1")
                                szFreightOwnerName = rs1("Name")
                                szMailAddr = rs1("MailAddr")
                            end if
                            
                            rs1.close
                            set rs1 = nothing
            
            %>
                   
                        <table border=0 cellspacing=1 width="100%" height="100%" align="center">
                            <tr>
                                <td valign="top">
                                    <table border=0 cellspacing=1 width="100%" align="center">
                                        <tr>
                                            <td colspan=5 align="center"><font size=5><%=szFreightOwnerName%>　<%=szVesselLineName%>　<%=szFaxNo%>　<%=szContact%>　NO#<%=szFreightOwnerID%></font></td>
                                        </tr>
                                        <tr>
                                            <td colspan=5 align="left"><br></td>
                                        </tr>
                                        <tr>
                                            <td colspan=5 align="center"><font size=5><span style="letter-spacing: 10pt"><%=szCmpChtName%></span></font></td>
                                        </tr>
                                        <tr>
                                            <td colspan=5 align="center"><font size=5><%=szCmpEngName%></font></td>
                                        </tr>
                                        <tr>
                                            <td colspan=5 align="center"><br></td>
                                        </tr>
                                        <tr>
                                            <td colspan=5 align="center"><font size=5>LIST OF CARGO MEASURE & WEIGHT</font></td>
                                        </tr>
                                        <tr>
                                            <td colspan=5 align="center">
                                            <%
                                                for j = 0 to 80
                                                    response.write "-"
                                                next
                                            %>
                                            </td>
                                        </tr>
                                        <tr valign="center">
                                            <td width="12%" align="right"></td> 
                                            <td width="18%" align="right">船名&航次 ：</td>
                                            <td width="24%" align="left"><%=szVesselName%> - <%=szVesselList%></td>
                                            <td width="16%" align="right">結關日期：</td>
                                            <td width="30%" align="left"><%=szVesselDate%></td>
                                        </tr>
                                        <tr valign="center"> 
                                            <td></td>
                                            <td align="right">船　公　司：</td>
                                            <td align="left"><%=szVesselOwner%></td>
                                            <td align="right">電　　話：</td>
                                            <td align="left"><%=szCmpPhoneNo%> <%=szContactName%></td>
                                        </tr>
                                        <tr valign="center">
                                            <td></td> 
                                            <td align="right">貨　櫃　場：</td>
                                            <td align="left"><%=szContainer%></td>
                                            <td align="right">傳　　真：</td>
                                            <td align="left"><%=szCmpFaxNo%> <%=szContactName%></td>
                                        </tr>
                                        <tr valign="center">
                                            <td></td> 
                                            <td align="right">航線 / 港口：</td>
                                            <td align="left"><%=szVesselLineName%></td>
                                            <td align="right">手　　機：</td>
                                            <td align="left"><%=szMobile%> <%=szContactName%></td>
                                        </tr>
                                        <tr>
                                            <td colspan=5 align="center">
                                            <%
                                                for j = 0 to 80
                                                    response.write "-"
                                                next
                                            %>
                                            </td>
                                        </tr>
                                   </table>
                                 
                                   <table border=0 cellspacing=1 width="75%" align="center">
                                        <tr align="center"> 
                                            <td align="right" width="1%"></td>
                                            <td align="right" width="10%"></td>
                                            <td align="right" width="16%">核對</td>
                                            <td align="right" width="16%">單號</td>
                                            <td align="right" width="16%">件數</td>
                                            <td align="right" width="16%">體積</td>
                                            <td align="right" width="16%">重量</td>
                                            <td align="right" width="9%"></td>
                                        </tr>
                                        
                                        <tr align="center"> 
                                            <td align="right"></td>
                                            <td align="right">=====</td>
                                            <td align="right">=====</td>
                                            <td align="right">=====</td>
                                            <td align="right">=====</td>
                                            <td align="right">=====</td>
                                            <td align="right">=====</td> 
                                            <td align="right"></td>                   
                                        </tr>
                                        
                                   <%
                                        nTotalPiece = 0
                                        nTotalVolume = 0
                                        nTotalWeight = 0
                                        
                                        dim nLineCounter, nPageCounter
                                        nLineCounter = 0
                                        nPageCounter = 1
                                        
                                        while not rs.eof
                                            nNeededVolume = rs("NeededForestry")
                                            
                                            if ((nPageCounter = 1) and nLineCounter >= 21) or ((nPageCounter>1) and nLineCounter >= 36)  then
                                                nLineCounter = 0
                                                nPageCounter = nPageCounter + 1
                                                response.write "</table></td></tr></table>"
                                                response.write "<table border=0 cellspacing=1 width=100% height=100% align=center>"
                                                response.write "<tr><td valign=top>"
                                                response.write "<table border=0 cellspacing=1 width=75% align=center>"
                                            end if
                                    %>
                                            <tr align="center"> 
                                                <td align="right"  width="1%"></td>
                                                <td align="right"  width="10%">
                                                <%
                                                    if rs("HasDelivered") = TRUE then
                                                        response.write ""
                                                    else
                                                        response.write "新增"
                                                    end if
                                                %>
                                                </td>
                                                <td align="right" width="16%">
                                                <%
                                                    if rs("IsChecked") = FALSE then
                                                        response.write "未核對"
                                                    end if
                                                %>
                                                </td>
                                                <td align="right" width="16%"><%=rs("ID")%></td>
                                                <td align="right" width="16%"><%=rs("TotalPiece")%></td>
                                                <td align="right" width="16%">
                                                <%
                                                    if nNeededVolume <> 0 then
                                                        response.write FormatNumber(nNeededVolume, 2)
                                                    else
                                                        response.write FormatNumber(rs("TotalVolume"), 2)
                                                    end if
                                                %>
                                                </td>
                                                <td align="right" width="16%"><%=MyFormatNumber(rs("TotalWeightSum"),1)%></td>
                                                <td align="right" width="9%"></td>
                                            </tr>                
                                    <%
                                            nTotalPiece = nTotalPiece + rs("TotalPiece")
                                            
                                            if nNeededVolume <> 0 then
                                                nTotalVolume = nTotalVolume + nNeededVolume
                                            else
                                                nTotalVolume = nTotalVolume + rs("TotalVolume")
                                            end if
                                            
                                            nTotalWeight = nTotalWeight + rs("TotalWeightSum")
                                            
                                            nLineCounter = nLineCounter + 1
                                            
                                            rs.movenext
                                        wend
                                            
                                        
                                    %>
                                    
                                        <tr align="center"> 
                                            <td align="right"></td>
                                            <td align="right">=====</td>
                                            <td align="right">=====</td>
                                            <td align="right">=====</td>
                                            <td align="right">=====</td>
                                            <td align="right">=====</td>
                                            <td align="right">=====</td>
                                            <td align="right"></td>
                                            
                                        </tr>
                                        <tr align="center"> 
                                            <td align="right"></td>
                                            <td align="right"></td>
                                            <td align="right"></td>
                                            <td align="right">TOTAL：</td>
                                            <td align="right"><%=nTotalPiece%></td>
                                            <td align="right"><%=FormatNumber(nTotalVolume, 2)%></td>
                                            <td align="right"><%=MyFormatNumber(nTotalWeight,1)%></td>
                                            <td align="right"></td>
                                        </tr>
                                    </table>
                                    
                                    <table border=0 cellspacing=1 width="100%" align="center">
                                    <%
                                    	nLineCounter = nLineCounter + 2   '2行tab
                                    	
                                        '23-Apr2005: 順便列出超長超大的尺寸資料
                                        sql = "select * from FreightForm, FormToOwner where VesselID = '" + szVesselListID + "' "
                                        sql = sql + " and FreightForm.ID = FormToOwner.FormID and CLng(FormToOwner.OwnerID) = '" + szFreightOwnerID + "'"
                                        sql = sql + " and FormToOwner.VesselLine='" + szVesselLine + "'"
                                        
                                        if szReportType = 1 then    '僅核對過的
                                            sql = sql + " and IsChecked = TRUE"
                                        end if
                                        
                                        if szHotKey = "MailNew" then
                                        	sql = sql + " and HasDelivered = false"
                                        end if
                                        
                                        sql = sql + " order by ID"
                                        
                                        set rs = conn.execute(sql) 
                                        
                                        dim nBoardTmp, fFirstData
                                        
                                        fFirstData = 1
                                        
                                        while not rs.eof
                                            
                                            nBoardTmp = rs("Board")
                                            
                                            if nBoardTmp <= 1 then
                                                nBoardTmp = ""
                                            end if  
                                            
                                            If rs("Length") > 600 OR rs("Height") > 226 then
                                                '只有第一筆才show出欄位名稱
                                                if fFirstData = 1 then
                                                    fFirstData = 0
                                                    nLineCounter = nLineCounter + 2		'2行tab
                                      %>
                                      				<br><br>
                                                    <tr align="center">
                                                        <td align="center" colspan=3>超大超小尺寸資料:</td>
                                                    </tr>
                                                    <tr><td><br></td></tr>
                                                    <tr align="center">
                                                        <td align="right" width="7%"><font size=3>頁次</font></td> 
                                                        <td align="right" width="7%"><font size=3>單號</font></td>
                                                        <td align="right" width="7%"><font size=3>板數</font></td>
                                                        <td align="right" width="7%"><font size=3>堆量</font></td>
                                                        <td align="right" width="7%"><font size=3>件數</font></td>
                                                        <td align="right" width="7%"><font size=3>包裝</font></td>
                                                        <td align="right" width="8%"><font size=3>長</font></td>
                                                        <td align="right" width="8%"><font size=3>寬</font></td>
                                                        <td align="right" width="8%"><font size=3>高</font></td>
                                                        <td align="right" width="9%"><font size=3>體積</font></td>                    
                                                        <td align="right" width="9%"><font size=3>單重</font></td>
                                                        <td align="right" width="9%"><font size=3>總重</font></td>
                                                    </tr>
                                                    
                                                    <tr align="center"> 
                                                        <td align="right"><font size=3>====</font></td>
                                                        <td align="right"><font size=3>====</font></td>
                                                        <td align="right"><font size=3>====</font></td>
                                                        <td align="right"><font size=3>====</font></td>
                                                        <td align="right"><font size=3>====</font></td>
                                                        <td align="right"><font size=3>====</font></td>
                                                        <td align="right"><font size=3>=====</font></td>
                                                        <td align="right"><font size=3>=====</font></td>
                                                        <td align="right"><font size=3>=====</font></td>
                                                        <td align="right"><font size=3>=====</font></td>
                                                        <td align="right"><font size=3>=====</font></td>
                                                        <td align="right"><font size=3>=====</font></td>
                                                    </tr>
                                       <%
                                                    nLineCounter = nLineCounter + 3
                                                end if
                                       %>
                                                    <tr align="center"> 
                                                        <td align="right"><font size=3><%=rs("PageNo")%></font></td>
                                                        <td align="right"><font size=3><%=rs("ID")%></font></td>
                                                        <td align="right"><font size=3><%=nBoardTmp%></font></td>
                                                        <td align="right"><font size=3>
                                                        <%
                                                            if rs("IsPL") = TRUE then
                                                                response.write "堆量"
                                                            end if
                                                        %>
                                                        </font></td>
                                                        <td align="right"><font size=3><%=rs("Piece")%></font></td>
                                                        <td align="right"><font size=3>
                                                        <%
                                                            szPackageStyle = ""
                                                            if rs("PackageStyleID") <> "0" then
                                                                sql2 = "select * from PackageStyle where ID = '" + CStr(rs("PackageStyleID")) + "'"
                                                                set rs2 = conn.execute(sql2)
                                                                if not rs2.eof then
                                                                    szPackageStyle = rs2("Name")                                
                                                                end if
                                                                
                                                                rs2.close
                                                                set rs2 = nothing
                                                            end if
                                                            
                                                            if rs("Weight") = 0 then
                                                            	WeightTmp = ""
                                                            else
                                                            	WeightTmp = MyFormatNumber(rs("Weight"),1)
                                                            end if
                                                        %>
                                                        </font><%=szPackageStyle%></td>
                                                        <td align="right"><font size=3><%=rs("Length")%></font></td>
                                                        <td align="right"><font size=3><%=rs("Width")%></font></td>
                                                        <td align="right"><font size=3><%=rs("Height")%></font></td>
                                                        <td align="right"><font size=3><%=FormatNumber(rs("Volume"))%></font></td>
                                                        <td align="right"><font size=3><%=WeightTmp%></font></td>
                                                        <td align="right"><font size=3><%=MyFormatNumber(rs("TotalWeight"),1)%></font></td>
                                                        
                                                    </tr>
                                    <%
                                                    nLineCounter = nLineCounter + 1
                                            end if
                                            
                                            '25-Apr2005: 註記成"已輸出過"
                                            sql = "update FreightForm set HasDelivered = 1 where SN = " + CStr(rs("SN"))
                                            conn.execute(sql) 
                                            
                                            rs.movenext
                                        wend
                                    
                                        if ((nPageCounter = 1) and nLineCounter >= 20) or ((nPageCounter>1) and nLineCounter >= 35)  then                                        	
                                        	nLineCounter = 0
                                            nPageCounter = nPageCounter + 1
                                            response.write "</table></td></tr></table>"
                                            response.write "<table border=0 cellspacing=1 width=100% height=100% align=center>"
                                            response.write "<tr><td valign=top>"
                                            response.write "<table border=0 cellspacing=1 width=100% align=center>"
                                        end if
                                        
                                        
                                    	'07-Jun2005: 加入備註
                                    	szRemark = request (("Remark_") + CStr(nFreightOwnerCount))  '備註
                                    	if Trim(szRemark) <> "" then
                                    		response.write "<br><br>"                                        
                                        	nLineCounter = nLineCounter + 2	'2行tab
                                    %>
	                                        <table border=0 cellspacing=1 width=100% align=center>
	                                        <tr><br>
	                                    	<td width="10%" valign="top" align="left">備註:</td>
	                                    	<td width="90%" valign="top" align="left">
	                                    	<%
	                                    		response.write "" + szRemark + "<br>"
	                                    		nLineCounter = nLineCounter + CInt(Len(szRemark)/32) + 1
	                                    	%>
	                                    	</td>
	                                   		</tr>
	                                   		</table>
	                                <%
	                                		'寫入資料庫
	                                        dim rs3
	                                        sql = "select * from FreightReportRemark where CLng(OwnerID) = '" + szFreightOwnerID + "'"
	                                        set rs3 = conn.execute(sql)
	                                        
	                                        if not rs3.eof then
	                                        	sql = "update FreightReportRemark set Remark = '" + szRemark + "' where CLng(OwnerID) = '" + szFreightOwnerID + "'"
	                                        	conn.execute(sql)
	                                        else
	                                        	sql = "insert into FreightReportRemark(OwnerID, Remark) values ('" + FormatString(szFreightOwnerID,3) + "','" + szRemark + "')"
	                                        	conn.execute(sql)
	                                        end if                                        
	                                        
	                                        rs3.close
	                                        set rs3 = nothing 
	                                	end if
	                                	
	                                	nFreightOwnerCount = nFreightOwnerCount + 1
	                                	
	                                	if ((nPageCounter = 1) and nLineCounter >= 20) or ((nPageCounter>1) and nLineCounter >= 35)  then                                        	
                                            response.write "</table></td></tr></table>"
                                            response.write "<table border=0 cellspacing=1 width=100% height=100% align=center>"
                                            response.write "<tr><td valign=top>"
                                            response.write "<table border=0 cellspacing=1 width=100% align=center>"
                                        end if                                        
                                        
	                                %>
                                    </table>                                    
                                </td>
                            </tr>
                            
                        </table>
                             
<%
                        end if
                    next
                end if
            next
        end if
    end if
    
    
    rs.close
    conn.close
    
    set rs=nothing
    set conn=nothing
%>  
    
    
</table>

 
</form>
</body>
</html>