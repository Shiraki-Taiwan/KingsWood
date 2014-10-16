<!-- #include file = ../GlobalSet/conn.asp -->
<!-- #include file = ../GlobalSet/ShareFun.asp -->
<!--進倉單資料設定:新增-->

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@CMT") < 0 then
		response.redirect "../Login/Login.asp"
	end if


    '===============接收參數===============
    dim szStatus
    szStatus = request("Status")    '編輯Status
    
    dim nGotoFFSearch
    nGotoFFSearch = request("GotoFFSearch")
    
    dim szVesselID
    dim szStorehouse(50), szPackageStyleID(50), szID(50)
    dim nPageNo(50), nBoard(50), nPiece(50)
    dim fLength(50), fWidth(50), fHeight(50), fVolume(50), fWeight(50), fTotalWeight(50)
    dim bIsPL(50)
    dim nSN(50)
    dim IDTmp
    
    szVesselID              = request("VesselID")           '船碼
    
    dim i, nDataCounter
    
    dim nTotalLineCounter, nStrLen
    nTotalLineCounter = 50
    
    nDataCounter = 0
    
    for i = 0 to nTotalLineCounter
        nSN(i)              = request("SN_" + CStr(i))               'Serial numberX
        szID(i)             = request("ID_" + CStr(i))                '倉單
        nPageNo(i)          = request("PageNo_" + CStr(i))            '頁次
        szStorehouse(i)     = request("Storehouse_" + CStr(i))        '倉位
        nBoard(i)           = request("Board_" + CStr(i))             '板數    
        bIsPL(i)            = request("IsPL_" + CStr(i))              '形 式 ( P/L )
        nPiece(i)           = request("Piece_" + CStr(i))             '件數
        szPackageStyleID(i) = request("PackageStyleID_" + CStr(i))    '包裝    
        fLength(i)          = request("Length_" + CStr(i))            '長
        fWidth(i)           = request("Width_" + CStr(i))             '寬
        fHeight(i)          = request("Height_" + CStr(i))            '高
        '17-Jun2005: 體積改成用計算的, 用以修正改了長寬高等資訊, 卻沒按enter來fire "onblur" event
        '所造成的體積沒重新計算問題
        'fVolume(i)          = request("Volume_" + CStr(i))            '體積
        fWeight(i)          = request("Weight_" + CStr(i))            '單重
        fTotalWeight(i)     = request("TotalWeight_" + CStr(i))       '總重
        
        
        if szID(i) <> "" then
            IDTmp = szID(i)
            nStrLen = Len(szID(i))
            
            if nStrLen = 1 then
                szID(i) = "000" + CStr(IDTmp)
            elseif nStrLen = 2 then
                szID(i) = "00" + CStr(IDTmp)
            elseif nStrLen = 3 then
                szID(i) = "0" + CStr(IDTmp)
            else
                szID(i) = CStr(IDTmp)
            end if
                 
        end if
        
        if Trim(nBoard(i)) = "" then
            nBoard(i) = 0
        end if
        
        if Trim(nPiece(i)) = "" then
            nPiece(i) = 0
        end if
        
        if Trim(fLength(i)) = "" then
            fLength(i) = 0
        end if
        
        if Trim(fWidth(i)) = "" then
            fWidth(i) = 0
        end if
        
        if Trim(fHeight(i)) = "" then
            fHeight(i) = 0
        end if
        
        if Trim(fVolume(i)) = "" then
            fVolume(i) = 0
        end if
        
        if Trim(fWeight(i)) = "" then
            fWeight(i) = 0
        end if
        
        if Trim(fTotalWeight(i)) = "" then
            fTotalWeight(i) = 0
        end if
        
        '17-Jun2005: 體積改成用計算的
	    fVolume(i) = VolumeCalculator(nBoard(i), bIsPL(i), nPiece(i), fLength(i), fWidth(i), fHeight(i))
	    
    next
    
    '===============寫入資料庫===============  
    dim sql, rs, rs1 
    set rs = nothing 
    set rs1 = nothing
  
    '=========新增=========
    if szStatus = "Add" then 
    
        '讀出倉單筆數
        dim nFormCount
        nFormCount = 0
        sql = "select DataCount from VesselList where ID = '" + szVesselID + "'"
        set rs = conn.execute(sql)
        if not rs.eof then
            nFormCount = rs("DataCount")
        end if
        
        '寫入細項資料
        
        dim HasDeliveredTmp
        HasDeliveredTmp = 0
        
        '逐筆加入資料庫
        for i = 0 to nTotalLineCounter
            if nPiece(i) <> 0 or fLength(i) <> 0 or fWidth(i) <> 0 or fHeight(i) <> 0 then
                sql = "select HasDelivered from FreightForm where ID='" + szID(i) + "' and VesselID='" + szVesselID + "'"
                set rs1 = conn.execute(sql)
                while not rs1.eof
                    if rs1("HasDelivered") = TRUE then
                        HasDeliveredTmp = 1
                    end if
                    
                    rs1.movenext
                    bFound = 1
                wend
                
                rs1.close
                
                sql = "insert into FreightForm(ID, PageNo, Storehouse, Board, IsPL, Piece, PackageStyleID, Length, " &_
                      "Width, Height, Volume, Weight, TotalWeight, VesselID, HasDelivered)" &_
                      " values('" + szID(i) + "','" + CStr(nPageNo(i)) + "','" + szStorehouse(i) + "'," + CStr(nBoard(i)) + "," &_ 
                      CStr(bIsPL(i)) + "," + CStr(nPiece(i)) + ",'" + szPackageStyleID(i) + "'," + CStr(fLength(i)) + "," &_
                      CStr(fWidth(i)) + "," + CStr(fHeight(i)) + "," + CStr(fVolume(i)) + "," + CStr(fWeight(i)) + "," &_
                      CStr(fTotalWeight(i)) + ",'" + szVesselID + "'," + CStr(HasDeliveredTmp) + ")"
                 
                conn.execute(sql) 
                
                '寫入筆數 (累加)
                nFormCount = nFormCount + 1
                sql = "update VesselList set DataCount = " + CStr(nFormCount) + " where ID = '" + szVesselID + "'"
                
                conn.execute(sql)
                
            end if           
        next
        
        rs.close        
        conn.close
        set rs=nothing
        set rs1=nothing
        set conn=nothing
        
        if nGotoFFSearch = 1 then
            '到倉單查詢
            response.redirect "FFfun02b.asp?Status=FirstLoad&VesselListID=" + szVesselID
        else
            response.redirect "FFfun01a.asp?Status=Add&VesselID=" + szVesselID
        end if
                    
        
    '=========修改=========
    elseif szStatus = "Modify" then
    
        '修改細項資料
        for i = 0 to nTotalLineCounter
            if nSN(i) <> "" then
              	sql = "update FreightForm set Storehouse = '" + szStorehouse(i) &_
              	      "', Board = " + CStr(nBoard(i)) + ", IsPL = " + CStr(bIsPL(i)) + ", Piece = " + CStr(nPiece(i)) &_
              	      ", PackageStyleID = '" + szPackageStyleID(i) + "', Length = " + CStr(fLength(i)) &_
              	      ", Width = " + CStr(fWidth(i)) + ", Height = " + CStr(fHeight(i)) + ", Volume = " + CStr(fVolume(i)) &_
              	      ", Weight = " + CStr(fWeight(i)) + ", TotalWeight = " + CStr(fTotalWeight(i)) &_
              	      ", VesselID = '" + szVesselID + "',ID = '"+ szID(i) +"' ,PageNo = '" + CStr(nPageNo(i)) + "'" &_
              	      " where SN = " + nSN(i)
              
                conn.execute(sql)
            end if
        next
      	
      	conn.close
        set conn=nothing
        response.redirect "../Voyage/VOfun01a.asp"
    
      
    '=========刪除=========
    elseif szStatus = "Delete" then
        sql = "delete from FreightForm where ID ='" + szID(0) + "'"
      	conn.execute(sql)
      	
      	conn.close
        set conn=nothing
        response.redirect "FFfun01a.asp"
    end if
%>