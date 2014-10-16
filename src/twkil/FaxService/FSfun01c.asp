<!-- #include file = ../GlobalSet/conn.asp -->
<!-- #include file = ../GlobalSet/ShareFun.asp -->
<!-- #include file = FSShare.asp -->
<!--攬貨報告書列印/傳真-->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>攬貨報告書列印/傳真</title>
	<!-- #include file = ../head_bundle.html -->
	<link href="../GlobalSet/sg.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="FSfun01c.js?t=20140930"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<%
    '===============接收參數===============
    dim szVesselListID, szFreightOwnerID, szFreightOwnerIDTmp, szStatus, szFormID, szCompanyID
    dim szSelectType, szReportType
    dim szVesselLine, szContainer
    dim szHotKey
    dim sql, rs
    
    szVesselListID   = request("VesselListID")     '航次
    szFreightOwnerID = request ("FreightOwnerID")  '攬貨商    
    szFormID = request ("FormID")                  '起始單號
    szCompanyID = request("CompanyID")             '公司名稱
    szStatus = request ("Status")
    szSelectType = request ("SelectType")          '依攬貨商或單號...
    szReportType = request ("ReportType")          '報表格式
    
    szVesselLine = request ("VesselLine")           '航線
    
    'szContainer = request ("Container")             '貨櫃場
       
    szHotKey = request ("HotKey")		'是否使用快速鍵
    
    if szHotKey = "MailAll" or szHotKey = "MailNew" then
    	szFreightOwnerID = "1-999" 		'預設
    	sql = "select MAX(FreightOwner.ID) as MaxId from FreightOwner"
    	set rs = conn.execute(sql)
    	
    	if not rs.eof then
'---------------2009 MAR edit, 最新的攬貨商編號有問題
'    		szFreightOwnerID = "1-" & rs("MaxId")    	
    	end if    	
    end if
    
    szFreightOwnerIDTmp = szFreightOwnerID
    
    '===========Parse 單號/攬貨商===================
    dim szSubStr, szUnParseedID
    
    if szSelectType = 1 then    '依單號
        szSubStr = szFormID        
    else                        '依攬貨商
        szSubStr = szFreightOwnerID
    end if
    
    szUnParseedID = szSubStr
    
    ParseFormID()
        
    '===============查詢資料庫===============  
    
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
        szContainer  = rs("ContainerYard")
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
    
   dim szMailAddress, szMailAddressCC, sql2, rs2
   szMailAddress = ""
   szMailAddressCC = ""
   'szMailAddress = request("MailAddress")
%>

<!------------------------列表----------------------------------->
<form name="form" method="post" OnKeyDown="CheckHotKey()">
<table width="100%" height=""100%"">
<%
    dim nNeededVolume
        
    dim nFoundCount
    '===================詳細尺寸資料==============================
    if szReportType = 2 then            
	
%>
        <tr> 
            <td >
                <table border=0 cellspacing=1 width="95%" height=""100%"" align="center">
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
               <table border=0 cellspacing=0 width="95%" height=""100%"" align="center">
                    <tr align="center">
                        <td align="right" width="7%"><font size=3>倉位</font></td> 
                        <td align="right" width="7%"><font size=3>單號</font></td>
                        <!--<td align="right" width="7%"><font size=3>倉位</font></td>-->
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
                    
                        '20-Apr2005: 把單號format成4位
                        szStartIDTmpFmt = szStartIDTmp(i)
                        szEndIDTmpFmt = szEndIDTmp(i)                        
                        szStartIDTmpFmt = FormatString (szStartIDTmpFmt, 4)
                        szEndIDTmpFmt = FormatString (szEndIDTmpFmt, 4)
                        
                        sql = sql + " and ID >= '" + szStartIDTmpFmt + "' and ID <= '" + szEndIDTmpFmt + "'"
                    end if
                    
                    '20-Jul2005: 以輸入順序顯示 (The priority is higher than PageNo)
                    sql = sql + " order by ID, SN, PageNo"
if szSelectType=2 then
	sql = "select FreightForm.* from FreightForm , FormToOwner where VesselID = '" + szVesselListID + "' and FreightForm.ID = FormToOwner.FormID and FormToOwner.OwnerID ='"+ FormatString(szFreightOwnerID, 3) +"' and FormToOwner.VesselLine='" + szVesselLine + "' order by ID"                
end if	
'response.write "SQL="&sql&"<BR>"
'Sql="select FreightForm.* from FreightForm, FormToOwner where VesselID = '00842' and FreightForm.ID = FormToOwner.FormID and FormToOwner.OwnerID = '107' and FormToOwner.VesselLine='01' order by ID"
'response.write "SQL="&sql&"<BR>"
'response.end
                    set rs = conn.execute(sql)                                      
                    while not rs.eof
%>
        <input type="hidden" name="FreightOwnerID_<%=nFoundCount%>" value="<%=szFreightOwnerID%>">  <!--攬貨商ID-->
        <input type="hidden" name="FreightOwnerName_<%=nFoundCount%>" value="<%=szFreightOwnerName%>">  <!--攬貨商名稱-->

<%                    
                        nFoundCount = nFoundCount + 1
                        
                        fVolume = rs("Volume")
                        
                        Dim nBoard
                        If rs("Board") <> 0 Then
                            nBoard = rs("Board")
                            If IsNumeric(nBoard) Then
                                nBoard = CLng(nBoard)
                            Else
                                nBoard = 1  'just for aborting error
                            End If                            
                        End If
                        
                        dim bChangeBgColor
                        bChangeBgColor = 0
                        
                        if IsIrregularVolume(nBoard, fVolume) = 1 then
			            	bChangeBgColor = 1
			            end if
			                  
			            if IsIrregularLength(rs("IsPL"), rs("Length")) = 1 then
			            	bChangeBgColor = 1
			            end if
			                  
			            if IsIrregularHeight(rs("IsPL"), rs("Height")) = 1 then
			            	bChangeBgColor = 1
			            end if
			            
			            if IsIrregularWidth(rs("IsPL"), rs("Width")) = 1 then
			            	bChangeBgColor = 1
			            end if
                        
                        If bChangeBgColor = 1 Then
                            szBgColor = "#FFFFA0"
                        else
                            szBgColor = "#FFFFFF"   
                        end if
                        
                        '2005-08-15: 查第一個攬貨商的e-Mail
                        if szMailAddress = "" then                        	
	                        sql2 = "select FreightOwner.MailAddr as MailAddress, FreightOwner.MailAddrCC as MailAddressCC from FreightForm, FormToOwner, FreightOwner where "
	                        sql2 = sql2 + " FormToOwner.FormID = FreightForm.ID"
	                        sql2 = sql2 + " and FormToOwner.OwnerID = FreightOwner.ID"
	                        sql2 = sql2 + " and FreightForm.ID = '" + rs("ID") + "'"
	                        sql2 = sql2 + " and FormToOwner.VesselLine = '" + szVesselLine + "'"
	                        
	                        set rs2 = conn.execute(sql2) 
	                        
	                        if not rs2.eof then
	                            szMailAddress = rs2("MailAddress")
	                            szMailAddressCC = rs2("MailAddressCC")
	                        end if
	                        
	                        rs2.close
                            set rs2 = nothing
	                    end if
                        
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
                        
                        '20-Nov2004: 不要每筆資料中間的橫線
                        'if nTotalPiece <> 0 then
                            'response.write "<tr><td colspan=13><hr size=""0""></td></tr>"
                        'end if
            %>                        
                        
                        <tr align="center">
                            <td align="right" style="background-color:<%=szBgColor%>"><font size=3><%=rs("Storehouse")%></font></td> 
                            <td align="right" style="background-color:<%=szBgColor%>"><font size=3><a href="../FreightForm/FFfun02b.asp?VesselListID=<%=szVesselListID%>&ID=<%=rs("ID")%>"><%=rs("ID")%></a></font></td>
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
                                dim szPackageStyle
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
            </td>
        </tr>    
           
        <tr>
            <td><br><hr size="0"><br></td>
        </tr>        
 
<%
    '===================總表格式==============================
    elseif szReportType = 0 or szReportType = 1 then
        nFoundCount = 0
        
        dim nTotalPiece, nTotalVolume, nTotalWeight, j
        nNeededVolume = 0
        
        '************依單號************
        if szSelectType = 1 then        	
%>
            <tr> 
                <td >
                    <table border=0 cellspacing=1 width="100%" height=""100%"" align="center">
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
                 
                   <table border=0 cellspacing=1 width="75%" height=""100%"" align="center">
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
                                '20-Apr2005: 把單號format成4位
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
                                
                                '2005-08-15: 查第一個攬貨商的e-Mail
		                        if szMailAddress = "" then                        	
			                        sql2 = "select FreightOwner.MailAddr as MailAddress, FreightOwner.MailAddrCC as MailAddressCC from FreightForm, FormToOwner, FreightOwner where "
			                        sql2 = sql2 + " FormToOwner.FormID = FreightForm.ID"
			                        sql2 = sql2 + " and FormToOwner.OwnerID = FreightOwner.ID"
			                        sql2 = sql2 + " and FreightForm.ID = '" + rs("ID") + "'"
			                        sql2 = sql2 + " and FormToOwner.VesselLine = '" + szVesselLine + "'"
			                        
			                        set rs2 = conn.execute(sql2) 
			                        
			                        if not rs2.eof then
			                            szMailAddress = rs2("MailAddress")
			                            szMailAddressCC = rs2("MailAddressCC")
			                        end if
			                        
			                        rs2.close
		                            set rs2 = nothing
			                    end if
                    %>
                            <tr align="center"> 
                                <td align="right"></td>
                                <td align="right">
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
                                <td align="right"><%=MyFormatNumber(rs("TotalWeightSum"), 1)%></td>
                            </tr>                
                    <%
                                nTotalPiece = nTotalPiece + rs("TotalPiece")
                                
                                if nNeededVolume <> 0 then
                                    nTotalVolume = nTotalVolume + nNeededVolume
                                else
                                    nTotalVolume = nTotalVolume + rs("TotalVolume")
                                end if
                                
                                nTotalWeight = nTotalWeight + rs("TotalWeightSum")
                            
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
               
            <tr>
                <td><br><hr size="0"><br></td>
            </tr>

    <%
        elseif szSelectType = 2 then    '依攬貨商
            dim nFreightOwnerID, iCurIndex
            dim bNeedSend
            
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
                        sql = sql + " and FreightForm.ID = FormToOwner.FormID and FormToOwner.OwnerID = '" + FormatString(szFreightOwnerID, 3) + "'"
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
            
            %>
                            <tr> 
                                <td >
                                    <table border=0 cellspacing=1 width="100%" height=""100%"" align="center">
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
                                                for j = 0 to 80
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
                                                for j = 0 to 80
                                                    response.write "-"
                                                next
                                            %>
                                            </td>
                                        </tr>
                                   </table>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                   <table border=0 cellspacing=1 width="75%" height=""100%"" align="center">
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
                                        
                                        while not rs.eof
                                            nNeededVolume = rs("NeededForestry")
                                    %>
                                            <tr align="center"> 
                                                <td align="right"></td>
                                                <td align="right">
                                                <%
                                                    if rs("HasDelivered") = TRUE then
                                                        response.write ""
                                                    else
                                                        response.write "新增"
                                                    end if
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
                                                <td align="right"><%=MyFormatNumber(rs("TotalWeightSum"), 1)%></td>
                                            </tr>                
                                    <%
                                            nTotalPiece = nTotalPiece + rs("TotalPiece")
                                            
                                            if nNeededVolume <> 0 then
                                                nTotalVolume = nTotalVolume + nNeededVolume
                                            else
                                                nTotalVolume = nTotalVolume + rs("TotalVolume")
                                            end if
                                            
                                            nTotalWeight = nTotalWeight + rs("TotalWeightSum")
                                        
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
                                            <td align="right"><%=MyFormatNumber(nTotalWeight, 1)%></td>
                                            <td align="right"></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>  
                            <tr>
                                <td><br><br></td>
                            </tr>                             
                            <tr>
                                <td>
                                    <table border=0 cellspacing=1 width="80%" height=""100%"" align="center">
                                    <%
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
                                      %>
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
                                                            	WeightTmp = MyFormatNumber(rs("Weight"), 1)
                                                            end if
                                                            	
                                                        %>
                                                        </font><%=szPackageStyle%></td>
                                                        <td align="right"><font size=3><%=rs("Length")%></font></td>
                                                        <td align="right"><font size=3><%=rs("Width")%></font></td>
                                                        <td align="right"><font size=3><%=rs("Height")%></font></td>
                                                        <td align="right"><font size=3><%=FormatNumber(rs("Volume"))%></font></td>
                                                        <td align="right"><font size=3><%=WeightTmp%></font></td>
                                                        <td align="right"><font size=3><%=MyFormatNumber(rs("TotalWeight"), 1)%></font></td>
                                                        
                                                    </tr>
                                       <%
                                                
                                            end if
                                            
                                            rs.movenext
                                        wend
                                    %>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                            	<%
                            		iCurIndex = iCurIndex + 1
                            		
                            		'查詢是否有記錄的備註
                            		dim szRemark
                            		dim rs3
                                    sql = "select * from FreightReportRemark where CLng(OwnerID) = '" + szFreightOwnerID + "'"
                                    set rs3 = conn.execute(sql)
                                    
                                    if not rs3.eof then
                                    	szRemark = rs3 ("Remark")
                                    else
                                    	szRemark = ""
                                    end if
                                    
                                    rs3.close
	                                set rs3 = nothing 
                            	%>
                            	<td>
                            		<br>
                            		<table border=0 cellspacing=1 width="80%" height=""100%"" align="center">
                                		<tr>
                                			<td align="left">備註：<textarea rows="2" cols="80" name="Remark_<%=nFoundCount%>" onfocus=SelectText(<%=iCurIndex%>) OnKeyDown="ChangeFocus(<%=iCurIndex+1%>)"><%=szRemark%></textarea></td>
                                		</tr>
                                	</table>
                                </td>
                            </tr>     
                            <tr>
                                <td><br><hr size="0"><br></td>
                            </tr>
                        <%
                            '查攬貨商資料
                            dim szPhone, szFaxNo, szContact, szFreightOwnerName, szMailAddr, szMailAddrCC
                            sql = "select * from FreightOwner where CLng(ID) = '" + szFreightOwnerID + "'"  
                            set rs = conn.execute(sql) 
                            
                            if not rs.eof then
                                szPhone = rs("Phone_1")
                                szFaxNo = rs("FaxNo_1")
                                szContact = rs("Contact_1")
                                szFreightOwnerName = rs("Name")
                                szMailAddr = rs("MailAddr")
                                szMailAddrCC = rs("MailAddrCC")
                                szFreightOwnerID = rs("ID")
                            end if
                       %>
                            <tr> 
                                <td > 
                                    <table border=0 cellspacing=1 width="85%" height=""100%"" align="center">
                                        <tr>             
                                            <td colspan=4 align="center"><font size=5>攬 貨 商 資 料</font></td>
                                        </tr>
                                        <tr>
                                            <td colspan=4 align="center">
                                            <%
                                                for j = 0 to 80
                                                    response.write "-"
                                                next
                                            %>
                                            </td>
                                        </tr>
                                        <tr>             
                                            <td width="20%" align="right"><span style="letter-spacing: 8pt">攬貨商：</span></td>
                                            <td width="30%" align="left"><%=szFreightOwnerName%></td>
                                        <%
                                            iCurIndex = iCurIndex + 1
                                        %>    
                                            <td width="20%" align="right"><span style="letter-spacing: 4pt">聯絡人：</span></td>
                                            <td width="30%" align="left">
                                                <input type="text" name="Contact_<%=nFoundCount%>" size="10" maxlength="10" value="<%=szContact%>" onfocus=SelectText(<%=iCurIndex%>) OnKeyDown="ChangeFocus(<%=iCurIndex+1%>)">
                                            </td>
                                        </tr>
                                        <tr>
                                        <%
                                            iCurIndex = iCurIndex + 1
                                        %>               
                                            <td align="right"><span style="letter-spacing: 2pt">聯絡電話：</span></td>
                                            <td align="left">
                                                <input type="text" name="Phone_<%=nFoundCount%>" size="10" maxlength="10" value="<%=szPhone%>" onfocus=SelectText(<%=iCurIndex%>) OnKeyDown="ChangeFocus(<%=iCurIndex+1%>)">
                                            </td>
                                        <%
                                            iCurIndex = iCurIndex + 1
                                        %>                
                                            <td align="right">傳真電話：</td>
                                            <td align="left">
                                                <input type="text" name="FaxNo_<%=nFoundCount%>" size="10" maxlength="10" value="<%=szFaxNo%>" onfocus=SelectText(<%=iCurIndex%>) OnKeyDown="ChangeFocus(<%=iCurIndex+1%>)">
                                            </td>
                                        </tr>
                                        <%
                                            iCurIndex = iCurIndex + 1
                                        %>   
                                        <tr>            
                                            <td align="right"><span style="letter-spacing: 4pt">E-Mail：</span></td>
                                            <td align="left" colspan=3 >
                                                <input type="text" name="MailAddr_<%=nFoundCount%>" size="50" maxlength="50" value="<%=szMailAddr%>" onfocus=SelectText(<%=iCurIndex%>) OnKeyDown="ChangeFocus(<%=iCurIndex+1%>)">
                                            </td>
                                        </tr>
                                        <%
                                            iCurIndex = iCurIndex + 1
                                        %>
                                         <tr>            
                                            <td align="right"><span style="letter-spacing: -1pt">E-Mail副本：</span></td>
                                            <td align="left" colspan=3 >
                                                <input type="text" name="MailAddrCC_<%=nFoundCount%>" size="50" maxlength="50" value="<%=szMailAddrCC%>" onfocus=SelectText(<%=iCurIndex%>) OnKeyDown="ChangeFocus(<%=iCurIndex+3%>)">
                                            </td>
                                        </tr>
                                        <input type="hidden" name="FreightOwnerID_<%=nFoundCount%>" value="<%=szFreightOwnerID%>">  <!--攬貨商ID-->
                                        <input type="hidden" name="FreightOwnerName_<%=nFoundCount%>" value="<%=szFreightOwnerName%>">  <!--攬貨商名稱-->
                                        <%
                                            iCurIndex = iCurIndex + 2
                                        %>
                                    </table>
                                </td>
                            </tr>
                            
                            <tr>
                                <td><br><hr size="0"></td>
                            </tr>
<%
                        end if
                    next
                end if
            next
        end if
    end if
%>  
    <tr>
    <%
        dim szStr
        
        response.write "<tr align=center><td><br>"

        if nFoundCount = 0 then
            
            if szHotKey = "MailAll" or szHotKey = "MailNew" then
                szStr = "<font size=6><br><meta http-equiv=""refresh"" content=""2; url=../FreightForm/FFfun02b.asp?Status=FirstLoad&VesselListID=" & szVesselListID & """><center>查詢不到任何資料!</center>"
                response.write(szStr) 
            else
                response.write "查詢不到任何資料"
                response.write "<tr align=center><td><br>"
                response.write "<input name=""Re_Search"" type=button style=""border-style: outset; border-width: 2"" value=""重新查詢"" name=""ReSearch"" OnKeyDown=""ReSearchByKeyPress()"" OnMouseUp=""OnReSearch()"" onfocusin=""SetFocusStyle(this, true, false)"" onfocusout=""SetFocusStyle(this, false, false)"">"                       
            end if
        else
            if szSelectType = 2 and szReportType <> 2 then    '依攬貨商          
              response.write "<input name=""SendFax"" type=button style=""border-style: outset; border-width: 2"" value=""送出傳真"" OnKeyDown=""FaxByKeyPress()"" OnMouseUp=""Fax()"" onfocusin=""SetFocusStyle(this, true, false)"" onfocusout=""SetFocusStyle(this, false, false)"">　"

            
            elseif szSelectType = 1 or (szSelectType=2 AND szReportType = 2)then    '依單號
            	response.write "主　　　旨：<input type=""text"" name=""Subject"" value=""攬貨報告書"" size=35 onfocus=""SelectText(0)"" OnKeyDown=""ChangeFocus(1)"" onfocusin=""SetFocusStyle(this, true, false)"" onfocusout=""SetFocusStyle(this, false, false)""> 　<br>"
            	response.write "E-Mail　　：<input type=""text"" value=""" & szMailAddress & """ name=""MailAddress"" size=35 onfocus=""SelectText(1)"" OnKeyDown=""ChangeFocus(2)"" onfocusin=""SetFocusStyle(this, true, false)"" onfocusout=""SetFocusStyle(this, false, false)""> 　<br>"
            	response.write "E-Mail副本：<input type=""text"" value=""" & szMailAddressCC & """ name=""MailAddressCC"" size=35 onfocus=""SelectText(2)"" OnKeyDown=""ChangeFocus(3)"" onfocusin=""SetFocusStyle(this, true, false)"" onfocusout=""SetFocusStyle(this, false, false)""> 　<br><br>"
              response.write "<input name=""SendFax"" type=button style=""border-style: outset; border-width: 2"" value=""送出傳真"" OnKeyDown=""FaxByKeyPress()"" OnMouseUp=""Fax()"" onfocusin=""SetFocusStyle(this, true, false)"" onfocusout=""SetFocusStyle(this, false, false)"">　"
            	
            end if
            response.write "<input name=""SendMail"" type=button style=""border-style: outset; border-width: 2"" value=""送出電子郵件"" OnKeyDown=""MailByKeyPress()"" OnMouseUp=""Mail()"" onfocusin=""SetFocusStyle(this, true, false)"" onfocusout=""SetFocusStyle(this, false, false)"">　"            
            response.write "<input name=""SendPrint"" type=button style=""border-style: outset; border-width: 2"" value=""列印"" OnKeyDown=""PrintByKeyPress()"" OnMouseUp=""Print()"" onfocusin=""SetFocusStyle(this, true, false)"" onfocusout=""SetFocusStyle(this, false, false)"">　"
            response.write "<input name=""Re_Search"" type=button style=""border-style: outset; border-width: 2"" value=""重新查詢"" name=""ReSearch"" OnKeyDown=""ReSearchByKeyPress()"" OnMouseUp=""OnReSearch()"" onfocusin=""SetFocusStyle(this, true, false)"" onfocusout=""SetFocusStyle(this, false, false)"">"
        end if       
        
        response.write "</td></tr>"
    %>
    </tr>
    
</table>
<input type="hidden" name="VesselListID" value="<%=szVesselListID%>">  <!--航次-->
<input type="hidden" name="VesselName" value="<%=szVesselName%>">  <!--船名-->
<input type="hidden" name="VesselList" value="<%=szVesselList%>">  <!--航次名稱-->
<input type="hidden" name="VesselDate" value="<%=szVesselDate%>">  <!--結關日-->
<input type="hidden" name="VesselOwner" value="<%=szVesselOwner%>">  <!--船公司-->
<input type="hidden" name="Status" value="<%=szStatus%>">  <!--狀態-->
<input type="hidden" name="FormID" value="<%=szFormID%>">   <!--起始單號-->
<input type="hidden" name="CompanyID" value="<%=szCompanyID%>">  <!--公司ID-->
<input type="hidden" name="SelectType" value="<%=szSelectType%>">   <!--單號/攬貨商-->
<input type="hidden" name="ReportType" value="<%=szReportType%>">  <!--公司ID-->
<input type="hidden" name="VesselLine" value="<%=szVesselLine%>">   <!--航線-->
<input type="hidden" name="Container" value="<%=szContainer%>">  <!--貨櫃場-->
<input type="hidden" name="FoundCount" value="<%=nFoundCount%>">  <!--有資料的攬貨商個數-->
<input type="hidden" name="FreightOwnerID" value="<%=szFreightOwnerIDTmp%>">   <!--攬貨商-->
<input type="hidden" name="UnParseedID" value="<%=szUnParseedID%>">   <!--未parse的ID-->
<input type="hidden" name="HotKey" value="<%=szHotKey%>">   <!--使用HotKey-->

<%=szFreightOwnerIDTmp%>
<%
    rs.close
    conn.close
    
    set rs=nothing
    set conn=nothing
%>  

 
</form>
</body>
</html>