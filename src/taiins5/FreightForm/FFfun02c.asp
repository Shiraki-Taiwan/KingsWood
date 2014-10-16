<!-- #include file = ../GlobalSet/conn.asp -->
<!-- #include file = ../GlobalSet/ShareFun.asp -->
<!--進倉單資料核定-->

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@TAIINS5") < 0 then
		response.redirect "../Login/Login.asp"
	end if



    '===============接收參數===============
    dim szStatus
    szStatus = request("Status")    '編輯Status
    
    dim szVesselListID
    
    szVesselListID = request ("VesselListID")            '航次
    
    dim nStoreSum_Piece, fStoreSum_Weight, fStoreSum_Forestry, fStoreSum_Volume
    dim szVesselID, szFreightOwnerID
    
    szVesselID              = request("VesselID")                 '船碼
    
    nStoreSum_Piece         = request("StoreSum_Piece")           '總件數 
    fStoreSum_Weight        = CDbl(request("StoreSum_Weight"))    '總重量 
    fStoreSum_Forestry      = CDbl(request("StoreSum_Forestry"))  '總才積 
    fStoreSum_Volume        = CDbl(request("StoreSum_Volume"))    '總體積 
    
    if nStoreSum_Piece = "" then
        nStoreSum_Piece = 0
    end if  

    if fStoreSum_Weight = "" then
        fStoreSum_Weight = 0
    end if
    
    if fStoreSum_Forestry = "" then
        fStoreSum_Forestry = 0
    end if
    
    if fStoreSum_Volume = "" then
        fStoreSum_Volume = 0
    end if  
   
    dim szFormID
    szFormID = request("ID")
  
    
    dim nPredictPiece, fPredictVolume, fPredictForestry, fNeededForestry, fPredictWeight
    nPredictPiece = request("PredictPiece")        '預測件數
    fPredictVolume = request("PredictVolume")      '預測體積
    fPredictWeight = request("PredictWeight")      '預測重量
    fPredictForestry = request("PredictForestry")  '預測才積
    fNeededForestry = request("NeededForestry")    '需求才積
    
    if nPredictPiece = "" then
        nPredictPiece = 0
    end if
    
    if fPredictVolume = "" then
        fPredictVolume = 0
    end if
    
    if fPredictForestry = "" then
        fPredictForestry = 0
    end if
    
    if fPredictWeight = "" then
        fPredictWeight = 0
    end if
    
    if fNeededForestry = "" then
        fNeededForestry = 0
    end if
    
    '細項資料
    dim szStorehouse(50), szPackageStyleID(50), szID(50)
    dim nPageNo(50), nBoard(50), nPiece(50)
    dim fLength(50), fWidth(50), fHeight(50), fVolume(50), fWeight(50), fTotalWeight(50)
    dim bIsPL(50), bDel(50)
    dim nSN(50)
    dim fWeightTmp, fTotalWeightTmp
    
    dim i, nDataCounter
    
    nDataCounter = request("DataCounter")
    
    for i = 0 to nDataCounter-1
        nSN(i)              = request("SN_" + CStr(i))               'Serial numberX
        szID(i)             = ucase(request("ID_" + CStr(i)))                '倉單
        
        '05-Mar2005: 若輸入的單號不滿4位, 以0補滿, ex, 999 ==> 0999    
        dim IDTmp
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
        
        nPageNo(i)          = request("PageNo_" + CStr(i))            '頁次
        szStorehouse(i)     = request("Storehouse_" + CStr(i))        '倉位
        nBoard(i)           = request("Board_" + CStr(i))             '板數    
        bIsPL(i)            = request("IsPL_" + CStr(i))              '堆量 ( P/L )
        nPiece(i)           = request("Piece_" + CStr(i))             '件數
        szPackageStyleID(i) = request("PackageStyleID_" + CStr(i))    '包裝    
        fLength(i)          = request("Length_" + CStr(i))            '長
        fWidth(i)           = request("Width_" + CStr(i))             '寬
        fHeight(i)          = request("Height_" + CStr(i))            '高
        '17-Jun2005: 體積改成用計算的, 用以修正改了長寬高等資訊, 卻沒按enter來fire "onblur" event
        '所造成的體積沒重新計算問題
        'fVolume(i)          = CDbl(request("Volume_" + CStr(i)))      '體積
        
        fWeightTmp = request("Weight_" + CStr(i))
        if fWeightTmp = "" then
        	fWeight(i) = 0      		'單重
        else
        	fWeight(i) = CDbl(fWeightTmp)      '單重
        end if
        
        fTotalWeightTmp = request("TotalWeight_" + CStr(i))
        if fTotalWeightTmp = "" then
        	fTotalWeight(i) = 0 '總重
        else
        	fTotalWeight(i) = CDbl(fTotalWeightTmp) '總重
        end if
        
        bDel(i)             = request("Del_" + CStr(i))               '是否打勾
        
        if nPageNo(i) = "" then
            nPageNo(i) = 0
        end if
        
        if nBoard(i) = "" then
            nBoard(i) = 0
        end if
        
        if nPiece(i) = "" then
            nPiece(i) = 0
        end if
        
        if fLength(i) = "" then
            fLength(i) = 0
        end if
        
        if fWidth(i) = "" then
            fWidth(i) = 0
        end if
        
        if fHeight(i) = "" then
            fHeight(i) = 0
        end if
        
        if fVolume(i) = "" then
            fVolume(i) = 0
        end if
        
        if fWeight(i) = "" then
            fWeight(i) = 0
        end if
        
        if fTotalWeight(i) = "" then
            fTotalWeight(i) = 0
        end if
        
        '17-Jun2005: 體積改成用計算的
	    fVolume(i) = VolumeCalculator(nBoard(i), bIsPL(i), nPiece(i), fLength(i), fWidth(i), fHeight(i))
    next
    
    '===============寫入資料庫===============  
    dim sql, rs 
    
    '=========修改=========
    if szStatus = "Modify" then
    
        '修改細項資料
        for i = 0 to nDataCounter-1
            if nPageNo(i) <> 0 and nSN(i) <> "" then
              	sql = "update FreightForm set ID = '" + szID(i) + "', Storehouse = '" + szStorehouse(i) &_
              	      "', Board = " + CStr(nBoard(i)) + ", IsPL = " + CStr(bIsPL(i)) + ", Piece = " + CStr(nPiece(i)) &_
              	      ", PackageStyleID = '" + szPackageStyleID(i) + "', Length = " + CStr(fLength(i)) &_
              	      ", Width = " + CStr(fWidth(i)) + ", Height = " + CStr(fHeight(i)) + ", Volume = " + CStr(fVolume(i)) &_
              	      ", Weight = " + CStr(fWeight(i)) + ", TotalWeight = " + CStr(fTotalWeight(i)) &_
              	      ", PageNo = '" + CStr(nPageNo(i)) + "'"
              	      
              	sql = sql + ", PredictPiece = " + CStr(nPredictPiece)
              	sql = sql + ", PredictVolume = " + CStr(fPredictVolume)
              	sql = sql + ", PredictForestry =" + CStr(fPredictForestry)
                'sql = sql + ", PredictWeight =" + CStr(fPredictWeight)
              	sql = sql + ", NeededForestry = " + CStr(fNeededForestry)
              	
              	sql = sql + " where SN = " + nSN(i)
               
                conn.execute(sql)
            end if
        next
      	
      	conn.close 
        set conn=nothing
        response.redirect "FFfun02b.asp?VesselListID=" + szVesselListID + "&ID=" + szFormID
    
    '=========核定=========
    elseif szStatus = "CheckIn" then
        '查攬貨商資料
        sql = "select OwnerID from FormToOwner where FormID = '" + szFormID + "'"
        set rs = conn.execute(sql)
        if not rs.eof then
            szFreightOwnerID = rs("OwnerID")
        else
            szFreightOwnerID = ""
            'set rs=nothing
            'set conn=nothing
            'response.write("<meta http-equiv=""refresh"" content=""3; url=FFfun02a.asp""><center>請先設定相對應的""攬貨商-單號對照資料""</center>")
            'response.end
        end if
        
        
        '寫入總資料
        sql = "select ID from StoreSum where ID='" + szFormID + "' and VesselListID = '" + szVesselListID + "'"           
        set rs = conn.execute(sql)
        
        if rs.eof then
            sql = "insert into StoreSum (ID, Piece, Forestry, Weight, Volume, VesselListID, FreightOwnerID, IsChecked) " &_
                  "values ('" + szFormID + "'," + CStr(nStoreSum_Piece) + "," + CStr(fStoreSum_Forestry) + "," &_
                  CStr(fStoreSum_Weight) + "," + CStr(fStoreSum_Volume) + ",'" + szVesselListID + "','" + szFreightOwnerID + "', TRUE )"
            
            conn.execute(sql)
            
            '更新FreightForm
            'for i = 0 to nDataCounter-1
                'sql = "update FreightForm set IsChecked = TRUE where SN = " + nSN(i)
                sql = "update FreightForm set IsChecked = TRUE, HasDelivered = FALSE where ID='" + szFormID + "' and VesselID = '" + szVesselListID + "'"
                conn.execute(sql)
            'next
            
            rs.close
            conn.close 
            set rs=nothing
            set conn=nothing
            response.redirect "FFfun02b.asp?VesselListID=" + szVesselListID + "&ID=" + szFormID
        
        else
            '編號重複, 回上一頁
            rs.close
            conn.close 
            set rs=nothing
            set conn=nothing
            response.write "<font size=6><br><meta http-equiv=""refresh"" content=""1; url=FFfun02b.asp?VesselListID=" + szVesselListID + "&ID=" + szFormID + """><center>此單已核定過</center>" 
        end if
    
    '=========取消核定=========
    elseif szStatus = "DelCheckIn" then 
        sql = "delete * from StoreSum where ID='" + szFormID + "' and VesselListID = '" + szVesselListID + "'"
        conn.execute(sql)
        
        '更新FreightForm
        'for i = 0 to nDataCounter-1
            'sql = "update FreightForm set IsChecked = FALSE where SN = " + nSN(i)
            sql = "update FreightForm set IsChecked = FALSE, HasDelivered = FALSE where ID='" + szFormID + "' and VesselID = '" + szVesselListID + "'"
            conn.execute(sql)
        'next
            
        
        conn.close 
        set conn=nothing
        response.redirect "FFfun02b.asp?VesselListID=" + szVesselListID + "&ID=" + szFormID
        
    '=========刪除=========
    elseif szStatus = "Delete" then 
        '讀出倉單筆數
        dim nFormCount
        nFormCount = 0
        sql = "select DataCount from VesselList where ID = '" + szVesselID + "'"
        set rs = conn.execute(sql)
        if not rs.eof then
            nFormCount = rs("DataCount")
        end if
        
        '刪除細項資料
        for i = 0 to nDataCounter-1
            if bDel(i) = "1" then
              	sql = "delete from FreightForm where SN = " + nSN(i)              
                conn.execute(sql)
                
                nFormCount = nFormCount - 1
            end if
        next
        
        
        '寫入筆數 (累加)        
        sql = "update VesselList set DataCount = " + CStr(nFormCount) + " where ID = '" + szVesselID + "'"
        conn.execute(sql)
        
        rs.close
        conn.close 
        set rs=nothing
        set conn=nothing
        response.redirect "FFfun02b.asp?Status=DeleteFinished&VesselListID=" + szVesselListID + "&ID=" + szFormID
  
    end if
    
  
%>