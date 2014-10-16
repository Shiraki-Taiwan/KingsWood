<!-- #include file = ../GlobalSet/conn.asp -->
<!-- #include file = ../GlobalSet/ShareFun.asp -->
<!-- #include file = FFShareFun.asp -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>進倉單資料修改/核定</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="FFfun02b.js"></script>
	<script type="text/javascript" src="FFShareFun.js"></script>
	<!--<script type="text/vbscript" src="FFShareFun.vbs"></script>-->
    <script type="text/javascript">
        function VolumeCalculator2(index) { }
        function WeightCalculator2(index) { }
        function OnPieceFocus(elem) { }
        function SimpleCalculator(elem) { }
    </script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<%
    dim sql, rs
    set rs = nothing
    
    '===========接參數===========
    dim szStatus, szID, szVesselListID, szFind, szPrevFoundID
    
    szVesselListID = request ("VesselListID")            '航次
    
    '查已核定單號時會用到
    dim szVesselID
    szVesselID = szVesselListID
    
    szID = request ("ID")               '單號
    
    szFind = request ("Find")           '要往前或往後一單號?
    
    szStatus = request ("Status")
    
    dim nPageNum
    nPageNum = request ("PageNum")
    
    if nPageNum = "" then
        nPageNum = 1
    else
        nPageNum = CLng(nPageNum)
    end if
	
    '19-Nov2004: if查詢不到某S/O的資料, 返回時要keep最後一次成功的查詢
    szPrevFoundID = request ("PrevFoundID")
    
    dim bFindID, szPrevID
    bFindID = 0
    if szID <> "" then
		szPrevID = szID
    else
		szPrevID = ""
    end if
    
    '找前/後一單號
    sql = "select ID from FreightForm where VesselID = '" + szVesselListID + "' group by ID order by ID"    
    set rs = conn.execute(sql)
    
    dim bGotNumberPart, szStrTmp, szPreStrTmp, szCurID
    
    if szFind = "Prev" then
        while not rs.eof or bFindID <> 1
            
            szID = UCase(szID) & szStrTmp
            
            if not rs.eof then
                szCurID = rs("ID")
            end if
            
            szCurID = UCase(szCurID) & szStrTmp
                        
            
            if szID = szCurID then
                bFindID = 1
                
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
            
            szID = UCase(szID)' & szStrTmp
          
            if not rs.eof then
                szCurID = rs("ID")
            end if
            
            szCurID = UCase(szCurID) '& szStrTmp
            
            if szID = szCurID then
                bFindID = 1
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
    else
        if szStatus = "FirstLoad" then            
            if not rs.eof then
                szID = rs("ID")
            end if
            szPrevFoundID = szID
        elseif szStatus = "UsePrevFoundID" then 
            szID = szPrevFoundID
            nPageNum = 1
        end if
        
    end if

    if szID <> "" then
        IDTmp = szID
        
        nStrLen = Len(szID)
        
        if nStrLen = 1 then
            szID = "000" + CStr(IDTmp)
        elseif nStrLen = 2 then
            szID = "00" + CStr(IDTmp)
        elseif nStrLen = 3 then
            szID = "0" + CStr(IDTmp)
        else
            szID = CStr(IDTmp)
        end if
    end if
    
    '===========查詢欲修改的資料=========== 
    
    dim szIDTmp(50), szStorehouseTmp(50), szPackageStyleIDTmp(50)
    dim szBgColor(50), szColorLength(50), szColorWidth(50), szColorHeight(50), szColorVolume(50)
    dim nPageNoTmp(50), nBoardTmp(50), nPieceTmp(50)
    dim fLengthTmp(50), fWidthTmp(50), fHeightTmp(50), fVolumeTmp(50), fWeightTmp(50), fTotalWeightTmp(50)
    dim bIsPLTmp(50)
    dim nSNTmp(50)
    dim fVolumeForCheckingColor
    
    dim bIsChecked

    sql = "select ID from StoreSum where ID='" + CStr(szID) + "' and VesselListID = '" + CStr(szVesselListID) + "'"
      
    set rs = Conn.execute(sql)
    if not rs.eof then
        bIsChecked = 1
    else
        bIsChecked = 0    
    end if
    
    dim szVesselNoTmp, szVesselNameTmp, VesselDateTmp, szVesselLine
    
    sql = "select * from VesselList where ID = '" + szVesselListID + "'"     
    set rs = conn.execute(sql)
    if not rs.eof then
        szVesselNoTmp       = rs("VesselNo")        '航次 
        szVesselNameTmp     = rs("VesselName")      '船名
        VesselDateTmp       = rs("Date")            '結關日 
        szVesselIDTmp       = rs("ID")              '船碼
        szVesselLine        = rs("VesselLine")      '航線
    end if
    
    
    
    dim nStoreSum_PieceTmp, fStoreSum_WeightTmp, fStoreSum_ForestryTmp, fStoreSum_VolumeTmp            
    nStoreSum_PieceTmp     = 0          '總件數
    fStoreSum_WeightTmp    = 0          '總重量
    fStoreSum_ForestryTmp  = 0          '總才積
    fStoreSum_VolumeTmp    = 0          '總體積
      
    dim nPredictPiece, fPredictVolume, fPredictForestry, fNeededForestry
      
    '19-Dec2004: For 分頁功能 ===========================
    
    dim nMaxDataCounter, nTotalNumInDB, nTotalPage
    nMaxDataCounter = 50     '每頁最多的筆數    
   
    
    '計算全部筆數
    sql = "select count(ID) as TotalDataNum from FreightForm where ID = '" + szID + "'" 
    sql = sql + " and VesselID = '" + szVesselListID + "'"    
    set rs = conn.execute(sql)
    
    if not rs.eof then
        nTotalNumInDB = rs("TotalDataNum")
    else
        nTotalNumInDB = 0
    end if
    
    
    '計算頁次
    if (nTotalNumInDB mod nMaxDataCounter) = 0 then
        nTotalPage = (nTotalNumInDB \ nMaxDataCounter)
    else
        nTotalPage = (nTotalNumInDB \ nMaxDataCounter) + 1
    end if
    
    '計算該略過的筆數=>其下一筆資料即為第一個該顯示的資料
    dim nSkipDataNum
    nSkipDataNum = nMaxDataCounter * (nPageNum-1)
    
    '=====================================================
    
    '查細項資料
    dim nDataCounter   
    nDataCounter = 0
    
    set rs = nothing
    
    '20-Jul2005: 以輸入順序顯示 (The priority is higher than PageNo)
    '04-Jan2005: 以輸入順序顯示
    sql = "select * from FreightForm where ID = '" + szID + "'" 
    sql = sql + " and VesselID = '" + szVesselListID + "' order by SN, PageNo"
    
    set rs = conn.execute(sql)
    
    dim nSkipDataNumTmp
    nSkipDataNumTmp = nSkipDataNum
    
    dim iDataIndex, bSkipAnyPiece
    bSkipAnyPiece = false
    
    
    while not rs.eof
        '只有指定頁次的資料需要顯示
        if nSkipDataNumTmp <= 0 and nDataCounter < nMaxDataCounter then
            nSNTmp(nDataCounter)              = rs("SN")                'Primary key: Serial no
            nPageNoTmp(nDataCounter)          = rs("PageNo")            '頁次
            szIDTmp(nDataCounter)             = rs("ID")                '單號
            szStorehouseTmp(nDataCounter)     = rs("Storehouse")        '倉位
            nBoardTmp(nDataCounter)           = rs("Board")             '板數    
            bIsPLTmp(nDataCounter)            = rs("IsPL")              '堆量 ( P/L )
            nPieceTmp(nDataCounter)           = rs("Piece")             '件數
            szPackageStyleIDTmp(nDataCounter) = rs("PackageStyleID")    '包裝    
            fLengthTmp(nDataCounter)          = rs("Length")            '長
            fWidthTmp(nDataCounter)           = rs("Width")             '寬
            fHeightTmp(nDataCounter)          = rs("Height")            '高
            fVolumeTmp(nDataCounter)          = rs("Volume")            '體積
            fWeightTmp(nDataCounter)          = rs("Weight")            '單重
            fTotalWeightTmp(nDataCounter)     = rs("TotalWeight")       '總重 
            
            '05-Jan2005:若件數為0, 則顯示空白
            if nPieceTmp(nDataCounter) = 0 then
                nPieceTmp(nDataCounter) = ""
                
                bSkipAnyPiece = true
            end if
            
            '[2007.03.20]:如果重量為0, 則顯示空白
            if fWeightTmp(nDataCounter) = 0 then
                fWeightTmp(nDataCounter) = ""
            end if
            
            if fTotalWeightTmp(nDataCounter) = 0 then
                fTotalWeightTmp(nDataCounter) = ""
            else
            	fTotalWeightTmp(nDataCounter) = MyFormatnumber(fTotalWeightTmp(nDataCounter), 1)
            end if
            
            fVolumeForCheckingColor = fVolumeTmp(nDataCounter)
            '
            
            '以下檢查是否為"非一般"的尺寸
            dim ColorPink, ColorYellow, ColorWhite            
            ColorYellow = "#FFFFA0"
		    ColorWhite = "#ffffff"
		    ColorPink  = "#FFC8FF"
    
            szBgColor(nDataCounter) = ColorYellow
            szColorLength(nDataCounter) = ColorYellow
            szColorWidth(nDataCounter) = ColorYellow
            szColorHeight(nDataCounter) = ColorYellow
            szColorVolume(nDataCounter) = ColorYellow
            
            dim bChangeBgColor
            bChangeBgColor = 0
            
            if IsIrregularVolume(nBoardTmp(nDataCounter), fVolumeForCheckingColor) = 1 then
            	bChangeBgColor = 1
            	szColorVolume(nDataCounter) = ColorPink
            end if
                  
            if IsIrregularLength(bIsPLTmp(nDataCounter), fLengthTmp(nDataCounter)) = 1 then
            	bChangeBgColor = 1
            	szColorLength(nDataCounter) = ColorPink
            end if
                  
            if IsIrregularHeight(bIsPLTmp(nDataCounter), fHeightTmp(nDataCounter)) = 1 then
            	bChangeBgColor = 1
            	szColorHeight(nDataCounter) = ColorPink
            end if
            
            if IsIrregularWidth(bIsPLTmp(nDataCounter), fWidthTmp(nDataCounter)) = 1 then
            	bChangeBgColor = 1
            	szColorWidth(nDataCounter) = ColorPink
            end if
                    
            If bChangeBgColor = 0 then
                szBgColor(nDataCounter) = ColorWhite
                szColorLength(nDataCounter) = ColorWhite
                szColorWidth(nDataCounter) = ColorWhite
                szColorHeight(nDataCounter) = ColorWhite
                szColorVolume(nDataCounter) = ColorWhite
            end if
            
            if nBoardTmp(nDataCounter) <= 1 then
                nBoardTmp(nDataCounter) = ""
            end if  
            
            nDataCounter = nDataCounter + 1
            
        end if
        
        nPredictPiece                     = rs("PredictPiece")      '預測件數
        fPredictVolume                    = rs("PredictVolume")     '預測體積
        fPredictForestry                  = rs("PredictForestry")   '預測才積
        fNeededForestry                   = rs("NeededForestry")    '需求才積
        
        '計算總量
        nStoreSum_PieceTmp = nStoreSum_PieceTmp + rs("Piece")
        fStoreSum_WeightTmp = fStoreSum_WeightTmp + rs("TotalWeight")                
        fStoreSum_VolumeTmp = fStoreSum_VolumeTmp + rs("Volume")        
        
        '06-Jun2005
        if rs("Piece") = 0 then
            bSkipAnyPiece = true
        end if
        
        iDataIndex = iDataIndex + 1
        nSkipDataNumTmp = nSkipDataNumTmp - 1
        rs.movenext            
    wend
    
    '18-Nov2004: 若沒找到資料,則顯示沒有S/O資料 
    '19-Nov2004: 若有找到資料,則將此S/O號碼記起來,用以keep最後一次成功的查詢
    if szStatus = "Search" then
        if nDataCounter = 0 then 
            response.redirect "FFfun02f.asp?VesselListID=" + szVesselListID + "&PrevFoundID=" + szPrevFoundID
        else
            szPrevFoundID = szID
        end if
    '07-Apr2005
    elseif szStatus = "FirstLoad" then
    	'這個else是避免當Firstload and 查不到任何資料時,會有沒回應的問題
    else                'if szStatus = "DeleteFinished" then
        '21-Nov2004: if刪除了全部的資料,則回到第一筆 
        '05-Mar2005: 只要查不到任何資料,就回到第一筆
        if nDataCounter = 0 then 
            response.redirect "FFfun02b.asp?Status=FirstLoad&VesselListID=" + szVesselListID + "&PrevFoundID=" + szPrevFoundID
        end if
    end if
    
    if nPredictPiece = 0 then
        nPredictPiece = ""
    end if
    
    if fPredictVolume = 0 then
        fPredictVolume = ""
    else
        fPredictVolume = FormatNumber(fPredictVolume, 2)
    end if
    
    if fPredictForestry = 0 then
        fPredictForestry = ""
    else
        fPredictForestry = FormatNumber (fPredictForestry, 2)
    end if
    
    if fNeededForestry = 0 then
        fNeededForestry = ""
    else
        fNeededForestry = Formatnumber (fNeededForestry, 2)
    end if
    
    
    '計算總才積
    fStoreSum_ForestryTmp = FormatNumber(fStoreSum_VolumeTmp * 35.3445, 2)
    'fStoreSum_ForestryTmp = fStoreSum_VolumeTmp * 35.3445
    '檢查資料是否超出範圍
    if nDataCounter > 50 then
        nDataCounter = 50
    end if
    
    '計算該船的總體積
    dim nTotalVolumeOnVessel
    sql = "select Sum(Volume) as TotalVolume from FreightForm where VesselID = '" + szVesselListID + "'"
    set rs = conn.execute(sql)
    if not rs.eof then
        nTotalVolumeOnVessel = rs("TotalVolume")
    end if
    
    
    '計算PackageStyle的總數
    dim nPackageStyleCount
    sql="select Count(ID) as TotalCount from PackageStyle"
    
    set rs=conn.execute(sql)
    if not rs.eof then
        nPackageStyleCount = rs("TotalCount")
    end if
    
    '查攬貨商
    dim szFreightOwnerID, szFreightOwnerName
    set rs = nothing
    sql = "select FreightOwner.ID as ID, FreightOwner.Name as Name from FreightOwner, FormToOwner where FormToOwner.OwnerID = FreightOwner.ID"
    sql = sql + " and FormToOwner.FormID = '" + szID + "' and FormToOwner.VesselLine ='" + szVesselLine + "'"
    
    set rs=conn.execute(sql)
    if not rs.eof then
        szFreightOwnerID = rs("ID")
        szFreightOwnerName = rs("Name")
    end if
%>

<form name="form" method="post" action="FFfun02c.asp" onsubmit=" javascript: return checkform();"  OnKeyDown="CheckHotKey(); CheckSaveHotKey(); CheckPrivateHotKey(); return CheckForbiddenKey();">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
<tr> 
    <td> 
        <table cellspacing=0 cellpadding=0 width="100%" border=0>
            <tbody> 
                <tr bgcolor=#3366cc> 
                    <td width="1"><img src="../image/coin2ltb.gif" width="20" height="26" /></td>
					<td style="font-size: 14pt; color: #0000ff; text-align: center; width: 15%;">
						<a style="color: #fff;" href="#">
							計算機(=)
						</a>
					</td>
					<td style="font-size: 14pt; color: #0000ff; text-align: center; width: 15%;">
						<a style="color: #fff;" href="#" onclick="javascript:(function(){ if (document.form.IsChecked.value == 0) document.form.Piece_0.focus(); else alert('請取消核定再修改!'); })();">
							到件數欄(Esc)
						</a>
					</td>
					<td style="font-size: 14pt; color: #0000ff; text-align: center; width: 40%;">
						<span style="color: #fff;">進倉單資料查詢/修改/核定</span>
					</td>
					<td style="font-size: 14pt; color: #0000ff; text-align: center; width: 15%;">
						<a style="color: #fff;" href="../FaxService/FSfun01c.asp?ReportType=0&SelectType=2&HotKey=MailNew&VesselLine=<%=szVesselLine %>&VesselListID=<%=szVesselListID %>">
							e-mail新增(Ctrl+S)
						</a>
					</td>
					<td style="font-size: 14pt; color: #0000ff; text-align: center; width: 15%;">
						<a style="color: #fff;" href="../FaxService/FSfun01c.asp?ReportType=0&SelectType=2&HotKey=MailAll&VesselLine=<%=szVesselLine %>&VesselListID=<%=szVesselListID %>">
							e-mail全部(Ctrl+A)
						</a>
					</td>
                    <td width="1"><img src="../image/coin2rtb.gif" width="20" height="26" /></td>
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
                                <tr bgcolor=#C9E0F8 height="2">
                                    <td align="left"></td> 
                                    <td align="left">船名:<%=szVesselNameTmp%></td>
                                    <td align="left" colspan=2>航次:<%=szVesselNoTmp%></td>
                                    <td align="left" colspan=2>結關日:<%=VesselDateTmp%></td>
                                    <td align="left"></td> 
                                    <td align="left"></td>                
                                </tr>
                                <tr bgcolor=#C9E0F8 height="2">
                                    <td align="left" colspan=8><hr></td>
                                </tr>
                                <tr align="center" bgcolor=#C9E0F8 height="2"> 
                                    <td align="left" width="2%"></td>
                                    <td align="left" width="24%">單號S/O:
                                        <input type="text" name="ID" size="4" maxlength="10" value="<%=szID%>" onfocus=SelectText(0) OnKeyDown="SearchForm()" OnKeyUp="CheckIDLen()" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                    </td>  
                                    <td align="left" width="8%">件　　數:</td>
                                    <td align="left"  width="9%" id="td_TotalPiece"><font size="6"> 
                                    	<%
                                            if not bSkipAnyPiece then
                                                response.write nStoreSum_PieceTmp
                                            end if
                                        %>
                                    </td>
                            <%
                                if fNeededForestry = "" then
                            %>
                                    <td align="left" width="8%">體　　積:</td>
                                    <td align="left" width="9%" id="td_TotalVolume"><font size="6"> 
                                        <%
                                            if not bSkipAnyPiece then
                                                response.write FormatNumber(fStoreSum_VolumeTmp, 2)
                                            end if
                                        %>
                                    </td>                                    
                            <%
                                else
                            %>
                                    <td align="left" width="8%">體　　積:</td>
                                    <td align="left" width="9%" id="td_TotalVolume"><font size="6" color="#FF0000"> <%=FormatNumber(fNeededForestry, 2)%></font></td>
                            <%
                                end if
                            %>
                                    <td align="left" width="8%">核　　定:</td>
                                    <td align="left" width="15%"><font size="6">
                                    <%
                                        if bIsChecked = 0 then
                                            response.write ""
                                        else
                                            response.write "YES"
                                        end if
                                    %></font>
                                    </td>
                                </tr>
                                <tr align="center" bgcolor=#C9E0F8 height="2">
                                    <td align="left" ></td>
                                    <td align="left"><span style="letter-spacing: 7pt">攬貨商</span>: <%=szFreightOwnerName%></td>  
                                    <td align="left" colspan=2>預測件數:
                                        <input type="text" name="PredictPiece" size="4" maxlength="6" value="<%=nPredictPiece%>" onfocus=SelectText(1) OnKeyDown="ChangeFocus(2)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                    </td> 
                                    <td align="left" colspan=2>需求才積:
                                        <input type="text" name="NeededForestry" size="4" maxlength="6" value="<%=fNeededForestry%>"" onblur="UpdateTotalPieceAndVolume(); UpdateNeededForestryLoseFocus()" onfocus="SelectText(2); UpdateNeededForestryHasFocus()" OnKeyDown="ChangeFocus(3)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                    </td>
                                    
                                    <td align="left" colspan=2></td>
                                </tr>
                                <tr align="center" bgcolor=#C9E0F8 height="2"> 
                                    <td align="left" ></td>
                                    <td align="left">重　　量: <%=MyFormatNumber(fStoreSum_WeightTmp, 1)%></td>                                     
                                    <td align="left" colspan=2>預測才積:
                                        <input type="text" name="PredictForestry" size="4" maxlength="10" value="<%=fPredictForestry%>" onfocus=SelectText(3) OnKeyDown="ChangeFocus(4)" onblur="Predict_Volume()" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                    </td>
                                    <td align="left" colspan=2>預測體積:
                                <%
                                    if bIsChecked = 0 then  '未核定
                                %>
                                        <input type="text" name="PredictVolume" size="4" maxlength="10" value="<%=fPredictVolume%>" onfocus="SelectText(4); UpdatePredictVolumeHasFocus()" OnKeyDown="ChangeFocus(8)" onblur="Predict_Forestry(); UpdatePredictVolumeLoseFocus()" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                <%
                                    else
                                %>
                                        <input type="text" name="PredictVolume" size="4" maxlength="10" value="<%=fPredictVolume%>" onfocus="SelectText(4); UpdatePredictVolumeHasFocus()" OnKeyDown="OnPredictVolume()" onblur="Predict_Forestry(); UpdatePredictVolumeLoseFocus()" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                <%
                                    end if
                                %>    
                                    </td>
                                    <td align="left" colspan=2></td>
                                </tr>
                                    
                         </tbody>
                        </table>
                        <table cellspacing=0 cellpadding=0 width="100%" bgcolor=#ebebeb border=0 height="38">
                            <tbody>        
                                <tr align="center" bgcolor=#C9E0F8 height="1">
                                    <td width="3%"></td>
                                    <td width="8%"></td>
                                    <td width="8%"></td>
                                    <td width="8%"></td>
                                    <td width="6%"></td>
                                    <td width="5%"></td>
                                    <td width="8%"></td>
                                    <td width="6%"></td>
                                    <td width="8%"></td>
                                    <td width="7%"></td>
                                    <td width="8%"></td>
                                    <td width="8%"></td>
                                    <td width="8%"></td>
                                    <td width="8%"></td> 
                                    <td width="1%"></td>                     
                                </tr> 
                                
                                
                                <tr> 
                                    <td align="center" bgcolor=#C9E0F8 height="2" colspan=15><hr></td>                            
                                </tr> 
                                
                                <tr align="center" bgcolor=#C9E0F8 height="2"> 
                                    <td></td> 
                                    <td><font size="3">頁 次</font></td>
                                    <td><font size="3">單 號</font></td>
                                    <td><font size="3">倉 位</font></td>
                                    <td><font size="3">板 數</font></td>
                                    <td><font size="3">堆 量</font></td>
                                    <td><font size="3">件 數</font></td>
                                    <td><font size="3">包 裝</font></td>
                                    <td><font size="3">長</font></td>
                                    <td><font size="3">寬</font></td>  
                                    <td><font size="3">高</font></td>
                                    <td><font size="3">體 積</font></td>
                                    <td><font size="3">單 重</font></td> 
                                    <td><font size="3">總 重</font></td> 
                                    <td></td>                       
                                </tr> 
                         
                        <%
                            dim nLineCnt, nPreTextCount, nCurIndex
                            nLineCnt = 0
                            nPreTextCount = 5
                            nCurIndex = 0
                            
                             
                            '查詢包裝
                            sql="select * from PackageStyle order by ID"
                            set rs=conn.execute(sql)
                            
                            
                            for i = 0 to nDataCounter-1
                                if bIsChecked = 0  then 
                        %> 
                                    <%
                                        nCurIndex = nLineCnt + nPreTextCount
                                    %>
                                    <input type="hidden" name="SN_<%=i%>" value=<%=nSNTmp(i)%>>         <!--Serial number-->
                                    <tr align="center" bgcolor=#C9E0F8 height="2">
                                        <%
                                            nCurIndex = nCurIndex + 1 
                                        %>
                                        <td><input type="checkbox" name="Del_<%=i%>" value="1" style="border-color: #C9E0F8" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                        <%
                                            nCurIndex = nCurIndex + 1
                                        %>
                                        <td><input type="text" style="background-color: <%=szBgColor(i)%>; direction: rtl" name="PageNo_<%=i%>"  size="3" maxlength="3" value="<%=nPageNoTmp(i)%>" onfocus="SelectText(<%=nCurIndex%>)" OnKeyDown="OnDirectionKey_LField(<%=nCurIndex%>, <%=nDataCounter%>, <%=i%>);ChangeFocus(<%=nCurIndex+1%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>  
                                        <%
                                            nCurIndex = nCurIndex + 1
                                        %>
                                        <td><input type="text" style="background-color: <%=szBgColor(i)%>" name="ID_<%=i%>" size="4" maxlength="5" value="<%=szIDTmp(i)%>" onblur="Repeat(<%=nCurIndex%>)" onfocus="SelectText(<%=nCurIndex%>)" OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=nDataCounter%>, <%=i%>); ChangeFocus(<%=nCurIndex+3%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>  
                                        <%
                                            nCurIndex = nCurIndex + 1
                                        %>
                                        <td><input type="text" style="background-color: <%=szBgColor(i)%>" name="Storehouse_<%=i%>" size="3" maxlength="4" value="<%=szStorehouseTmp(i)%>" onfocus=SelectText(<%=nCurIndex%>) OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=nDataCounter%>, <%=i%>); ChangeFocus(<%=nCurIndex+2%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                        <%
                                            nCurIndex = nCurIndex + 1
                                        %>
                                        <td ><input type="text" style="background-color: <%=szBgColor(i)%>; direction: rtl" name="Board_<%=i%>" size="1" maxlength="2" value="<%=nBoardTmp(i)%>" onfocus=SelectText(<%=nCurIndex%>) OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=nDataCounter%>, <%=i%>); ChangeFocus(<%=nCurIndex+1%>)" onblur="VolumeCalculator2(<%=i%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                        <%
                                            nCurIndex = nCurIndex + 1
                                        %>
                                        <td >
                                            <select size="1" style="background-color: <%=szBgColor(i)%>" name="IsPL_<%=i%>" onmousewheel="javascript: return false;" OnKeyDown="OnIsPL(<%=i%>, <%=nCurIndex%>, <%=nCurIndex+1%>); OnDirectionKey_MField_OnlyRL(<%=nCurIndex%>, <%=nDataCounter%>, <%=i%>)" onblur="VolumeCalculator2(<%=i%>)">
                                            <%
                                                if (bIsPLTmp(i) = 0) then 
                                                    response.write "<option selected value=0>" + "" + "</option>"
                                                    response.write "<option value=1>" + "是" + "</option>"                                                    
                                                else
                                                    response.write "<option value=0>" + "" + "</option>"
                                                    response.write "<option selected value=1>" + "是" + "</option>"                                                    
                                                end if     
                            	  	        %>                                	  	
                                      	    </select>
                                        </td>
                                        <%
                                            nCurIndex = nCurIndex + 1
                                        %>
                                        <td><input type="text" style="background-color: <%=szBgColor(i)%>" dir="rtl" name="Piece_<%=i%>" size="3" maxlength="5" value="<%=nPieceTmp(i)%>" onblur="SimpleCalculator(<%=nCurIndex%>); VolumeCalculator2(<%=i%>)" onfocus="OnPieceFocus(<%=nCurIndex%>); SelectText(<%=nCurIndex%>)" OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=nDataCounter%>, <%=i%>); ChangeFocus(<%=nCurIndex+2%>)"  onkeyup="CallInputBox(<%=nCurIndex%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                        <%
                                            nCurIndex = nCurIndex + 1
                                        %>
                                        <td>
                                            <select size="1" style="background-color: <%=szBgColor(i)%>" name="PackageStyleID_<%=i%>" onmousewheel="javascript: return false;" OnKeyDown="OnDirectionKey_MField_OnlyRL(<%=nCurIndex%>, <%=nDataCounter%>, <%=i%>); OnPackageStyleID(<%=nPackageStyleCount%>, <%=i%>, <%=nCurIndex%>, <%=nCurIndex+1%>)">             
                                	  	        <option value="0"></option>
                                        	  	<%            	  	    
                                        	 		if nPackageStyleCount <> 0 then
                                        	 		    rs.movefirst
                                        	 		end if
                                        	 		
                                        	 		while not rs.eof
                                        	 		    if rs("ID") = szPackageStyleIDTmp(i) then
                                        	 		        response.write "<option selected value=" & rs("ID") & ">" & rs("Name") & "</option>" 
                                        	 		    else           		 	
                                                            response.write "<option value=" & rs("ID") & ">" & rs("Name") & "</option>" 
                                                        end if
                                        	 			rs.movenext
                                        	 		wend
                                        	 		
                                        	 	%> 	    
                                      	    </select>
                                        </td>
                                        <%
                                            nCurIndex = nCurIndex + 1
                                        %>
                                        <td><input type="text" style="background-color: <%=szColorLength(i)%>; direction: rtl" name="Length_<%=i%>" size="3" maxlength="4" value="<%=fLengthTmp(i)%>" onfocus=SelectText(<%=nCurIndex%>) OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=nDataCounter%>, <%=i%>); ChangeFocus(<%=nCurIndex+1%>)" onblur="VolumeCalculator2(<%=i%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                        <%
                                            nCurIndex = nCurIndex + 1
                                        %>
                                        <td><input type="text" style="background-color: <%=szColorWidth(i)%>; direction: rtl" name="Width_<%=i%>" size="3" maxlength="4" value="<%=fWidthTmp(i)%>"  onfocus=SelectText(<%=nCurIndex%>) OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=nDataCounter%>, <%=i%>); ChangeFocus(<%=nCurIndex+1%>)" onblur="VolumeCalculator2(<%=i%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                        <%
                                            nCurIndex = nCurIndex + 1
                                        %>
                                        <td><input type="text" style="background-color: <%=szColorHeight(i)%>; direction: rtl" name="Height_<%=i%>" size="3" maxlength="4" value="<%=fHeightTmp(i)%>" onfocus=SelectText(<%=nCurIndex%>) OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=nDataCounter%>, <%=i%>); ChangeFocus(<%=nCurIndex+2%>)" onblur="VolumeCalculator2(<%=i%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                        <%
                                            nCurIndex = nCurIndex + 1
                                        %>
                                        <td><input type="text" style="background-color: <%=szColorVolume(i)%>; direction: rtl" name="Volume_<%=i%>" size="5" maxlength="6" value="<%=FormatNumber(fVolumeTmp(i), 2)%>" onfocus=SelectText(<%=nCurIndex%>) OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=nDataCounter%>, <%=i%>); ChangeFocus(<%=nCurIndex+6%>)" onblur="VolumeCalculator2(<%=i%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                        <%
                                            nCurIndex = nCurIndex + 1
                                        %>
                                        <td><input type="text" style="background-color: <%=szBgColor(i)%>; direction: rtl" name="Weight_<%=i%>"  size="3" value="<%=fWeightTmp(i)%>" onfocus="SelectText(<%=nCurIndex%>);UpdateWeightHasFocus()" OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=nDataCounter%>, <%=i%>); ChangeFocus(<%=nCurIndex+1%>)" onblur="WeightCalculator2(<%=i%>);UpdateWeightLoseFocus()" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                        <%
                                            nCurIndex = nCurIndex + 1
                                        %>
                                        <td><input type="text" style="background-color: <%=szBgColor(i)%>; direction: rtl" name="TotalWeight_<%=i%>" size="4" value="<%=fTotalWeightTmp(i)%>" onfocus="SelectText(<%=nCurIndex%>);UpdateTotalWeightHasFocus()" OnKeyDown="OnDirectionKey_RField(<%=nCurIndex%>, <%=nDataCounter%>, <%=i%>); ChangeFocus(<%=nCurIndex+4%>)" onblur=UpdateTotalWeightLoseFocus() onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>                 
                                        <td></td>
                                    </tr>  
                                
                        <%      
                                else
                        %>
                                    <input type="hidden" name="SN_<%=i%>" value=<%=nSNTmp(i)%>>         <!--Serial number-->
                                    <tr align="center" bgcolor=#C9E0F8 height="2">
                                        <%
                                            nCurIndex = nCurIndex + 1 
                                        %>
                                        
                                        <td><input type="checkbox" disabled name="Del_<%=i%>" value="ON" style="border-color: #C9E0F8"></td>
                                        <td><input readonly type="text" style="background-color: <%=szBgColor(i)%>; direction: rtl" name="PageNo_<%=i%>"          size="3" value="<%=nPageNoTmp(i)%>" onfocus="SelectText(<%=nLineCnt%>+1+<%=nPreTextCount%>)" OnKeyDown="ChangeFocus(<%=nLineCnt%>+2+<%=nPreTextCount%>)"></td>  
                                        <td><input readonly type="text" style="background-color: <%=szBgColor(i)%>" name="ID_<%=i%>"              size="4" value="<%=szIDTmp(i)%>" onfocus="SelectText(<%=nLineCnt%>+2+<%=nPreTextCount%>)" OnKeyDown="ChangeFocus(<%=nLineCnt%>+5+<%=nPreTextCount%>)"></td>  
                                        <td align="right"><input readonly type="text" style="background-color: <%=szBgColor(i)%>" name="Storehouse_<%=i%>"      size="3" value="<%=szStorehouseTmp(i)%>" onfocus=SelectText(<%=nLineCnt%>+3+<%=nPreTextCount%>)></td>
                                        <td><input readonly type="text" style="background-color: <%=szBgColor(i)%>; direction: rtl" name="Board_<%=i%>"           size="1" value="<%=nBoardTmp(i)%>" onfocus=SelectText(<%=nLineCnt%>+4+<%=nPreTextCount%>)></td>
                                        <td>
                                            <select disabled size="1" style="background-color: <%=szBgColor(i)%>" name="IsPL_<%=i%>" OnKeyDown="ChangeFocus(<%=nLineCnt%>+6+<%=nPreTextCount%>)">
                                            <%
                                                if (bIsPLTmp(i) = 0) then 
                                                    response.write "<option value=1>" + "是" + "</option>"
                                                    response.write "<option selected value=0>" + "" + "</option>"
                                                else
                                                    response.write "<option selected value=1>" + "是" + "</option>"
                                                    response.write "<option value=0>" + "" + "</option>"
                                                end if     
                            	  	        %>                                	  	
                                      	    </select>
                                        </td>
                                        <td><input readonly type="text" style="background-color: <%=szBgColor(i)%>; direction: rtl" name="Piece_<%=i%>" size="3" value="<%=nPieceTmp(i)%>" ></td>
                                        <td>
                                            <select disabled size="1" style="background-color: <%=szBgColor(i)%>" name="PackageStyleID_<%=i%>" OnKeyDown="ChangeFocus(<%=nLineCnt%>+8+<%=nPreTextCount%>)">             
                                	  	        <option value="0"></option>
                                        	  	<%   
                                        	  	    if nPackageStyleCount <> 0 then
                                        	 		    rs.movefirst
                                        	 		end if
                                        	 		         	  	    
                                        	 		while not rs.eof
                                        	 		    if rs("ID") = szPackageStyleIDTmp(i) then
                                        	 		        response.write "<option selected value=" & rs("ID") & ">" & rs("ID") & "-" & rs("Name") & "</option>" 
                                        	 		    else           		 	
                                                            response.write "<option value=" & rs("ID") & ">" & rs("ID") & "-" & rs("Name") & "</option>" 
                                                        end if
                                        	 			rs.movenext
                                        	 		wend
                                        	 		
                                        	 	%> 	    
                                      	    </select>
                                        </td>
                                        <td><input readonly type="text" style="background-color: <%=szColorLength(i)%>; direction: rtl" name="Length_<%=i%>"          size="3" value="<%=fLengthTmp(i)%>"></td>
                                        <td><input readonly type="text" style="background-color: <%=szColorWidth(i)%>; direction: rtl" name="Width_<%=i%>"           size="3" value="<%=fWidthTmp(i)%>"></td>
                                        <td><input readonly type="text" style="background-color: <%=szColorHeight(i)%>; direction: rtl" name="Height_<%=i%>"          size="3" value="<%=fHeightTmp(i)%>" ></td>
                                        <td><input readonly type="text" style="background-color: <%=szColorVolume(i)%>; direction: rtl" name="Volume_<%=i%>"          size="5" value="<%=FormatNumber(fVolumeTmp(i), 2)%>" ></td>
                                        <td><input readonly type="text" style="background-color: <%=szBgColor(i)%>; direction: rtl" name="Weight_<%=i%>"          size="3" value="<%=fWeightTmp(i)%>" ></td>
                                        <td><input readonly type="text" style="background-color: <%=szBgColor(i)%>; direction: rtl" name="TotalWeight_<%=i%>"     size="4" value="<%=fTotalWeightTmp(i)%>"></td>                 
                                        <td></td>
                                    </tr> 
                        <%
                                end if                          
                                nLineCnt = nLineCnt + 15
                            next
                        %>	                            
	                            <tr><td align=center bgcolor=#C9E0F8 height="2"  colspan=15>
	                            	<%
                                        '19-Dec2004: 加入分頁
                                        if nTotalPage > 1 then
                                            response.write "頁次：" & nPageNum & "/" & nTotalPage & "　" 
                                            
                                            'if nPageNum > 2 then
                                            '    response.write "　<a href=FFfun02b.asp?ID=" + szID + "&VesselListID=" + szVesselListID + "&PageNum=1" + ">首頁</a>"
                                            'else
                                            '    response.write "　首頁"
                                            'end if 
                                            
                                            if nPageNum > 1 then
                                                response.write "　<a href=FFfun02b.asp?ID=" + szID + "&VesselListID=" + szVesselListID + "&PageNum=" + CStr(nPageNum - 1) + ">上一頁(Page Up)</a>"
                                            else
                                                response.write "　上一頁(Page Up)"
                                            end if
                                            
                                            if nPageNum < nTotalPage then
                                                response.write "　<a href=FFfun02b.asp?ID=" + szID + "&VesselListID=" + szVesselListID + "&PageNum=" + CStr(nPageNum+1) + ">下一頁(Page Down)</a>"
                                            else
                                                response.write "　下一頁(Page Down)"
                                            end if                           
                                            
                                            'if nPageNum + 1 < nTotalPage then
                                            '    response.write "　<a href=FFfun02b.asp?ID=" + szID + "&VesselListID=" + szVesselListID + "&PageNum=" + CStr(nTotalPage) + ">末頁</a>"
                                            'else
                                            '    response.write "　末頁"
                                            'end if
                                        else
                                            response.write "</td>"
                                        end if 
                                    %> 
                                    </td>
	                            </tr>
	                            <tr><td align="center" bgcolor=#C9E0F8 height="2" colspan=15><br></td></tr>    
                            </tbody> 
                        </table>
                    </td>
                </tr>
            </tbody> 
        </table>
    </td>
</tr>

<%    
    '查核定過的單號
    dim TextString
    call ShowText(conn,TextString)	
    
%>

<script lanuage="javascript">
<%= TextString %>
</script>

<tr> 
    <td> 
        <table cellspacing=0 cellpadding=0 width="100%" border=0>
            <tbody> 
               <tr bgcolor=#3366cc > 
                    <td colspan="9" width="100%" valign=center align=middle bgcolor=#3366cc> 
                    <input type="hidden" name="">         <!--*只為了加一個欄位*-->
                    <input type="hidden" name="">         <!--*只為了加一個欄位*-->
                    <input type="hidden" name="">         <!--*只為了加一個欄位*-->
                        <div align="center">
                        
                    <%
                        if bIsChecked = 0 and nDataCounter<>0 then
                    %>
                            <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="核定" name="CheckIn" OnKeyDown="CheckInByKeyPress()" OnMouseUp="OnCheckIn()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)"> 
                            <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="儲存" name="Save" OnKeyDown="SaveByKeyPress()" OnMouseUp="OnSave()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                            <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="移動" name="Move" OnKeyDown="MoveByKeyPress()" OnMouseUp="OnMove()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                            <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="刪除" name="Delete" OnKeyDown="DeleteByKeyPress()" OnMouseUp="OnDelete()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                            <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="回上一頁" name="ReSearch" OnKeyDown="ReSearchByKeyPress()" OnMouseUp="OnReSearch()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                    <%
                        elseif bIsChecked = 1 and nDataCounter<>0 then
                    %>
                            <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="取消核定" name="CheckIn" OnKeyDown="DelCheckInByKeyPress()" OnMouseUp="OnDelCheckIn()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                            <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="已核定，回上一頁" name="ReSearch" OnKeyDown="ReSearchByKeyPress()" OnMouseUp="OnReSearch()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">                            
                    <%
                        else
                    %>
                            <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="回上一頁" name="ReSearch" OnKeyDown="ReSearchByKeyPress()" OnMouseUp="OnReSearch()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                    <%
                        end if
                    %>
                            
                        </div>
                    </td>
                </tr>
                <tr> 
                    <td width=1><img height=38 src="../image/box1.gif" width=20></td>
                    <td width=1><img height=38 src="../image/box2.gif" width=9></td>
                    <td valign=center align=middle width=1 background="../image/box3.gif">&nbsp; </td>
                    <td width=1><img height=38 src="../image/box4.gif" width=27></td>
                    <td style="text-align: center; width: 25%; background-color: #3366cc">
						<span style="color: #fff;">儲存(F2)</span>
                    </td>
					<td style="text-align: center; width: 25%; background-color: #3366cc">
						<a href="../FreightForm/FFfun01a.asp?Status=Add&VesselID=<%=szVesselListID%>&PrevFoundID=<%=szPrevFoundID%>" style="color: #fff;">
							<span>倉單輸入(F8)</span>
						</a>
					</td>
					<td style="text-align: center; width: 25%; background-color: #3366cc">
						<a href="../FaxService/FSfun03c.asp?ReportType=1&VesselLine=<%=szVesselLine%>&VesselListID=<%=szVesselListID%>" style="color: #fff;">
							<span>總表查詢(Ctrl+Z)</span>
						</a>
					</td>
					<td style="text-align: center; width: 25%; background-color: #3366cc">
						<a href="../FaxService/FSfun03c.asp?ReportType=2&VesselLine=<%=szVesselLine%>&VesselListID=<%=szVesselListID%>" style="color: #fff;">
							<span>尺寸資料查詢(Ctrl+X)</span>
						</a>
                    </td>
                    <td width=1><img height=38 src="../image/box5.gif" width=20></td>
                </tr>
            </tbody>
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
<input type="hidden" name="Status" value=<%=szStatus%> />         <!--狀態-->
<input type="hidden" name="VesselID" value=<%=szVesselIDTmp%> />  <!--船碼-->
<input type="hidden" name="DataCounter" value=<%=nDataCounter%> />  <!--筆數-->
<input type="hidden" name="StoreSum_Piece" value=<%=nStoreSum_PieceTmp%> />         <!--總件數-->
<input type="hidden" name="StoreSum_Weight" value=<%=fStoreSum_WeightTmp%> />       <!--總重量-->
<input type="hidden" name="StoreSum_Forestry" value=<%=fStoreSum_ForestryTmp%> />   <!--總才積-->
<input type="hidden" name="StoreSum_Volume" value=<%=fStoreSum_VolumeTmp%> />       <!--總體積-->
<input type="hidden" name="VesselListID" value=<%=szVesselListID%> />  <!--航次-->
<input type="hidden" name="VesselLine" value=<%=szVesselLine%> />  <!--航線-->
<input type="hidden" name="IsChecked" value=<%=bIsChecked%> />  <!--核定與否-->
<input type="hidden" name="PrevFoundID" value=<%=szPrevFoundID%> />  <!---->
<input type="hidden" name="PageNum" value=<%=nPageNum%> />  <!---->
<input type="hidden" name="TotalPageNum" value=<%=nTotalPage%> />  <!---->


</form>
</body>
</html>
