<!-- #include file = ../GlobalSet/conn.asp -->
<!--移動倉單-->

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@DEFAULT") < 0 then
		response.redirect "../Login/Login.asp"
	end if


    '===============接收參數===============
    dim szStatus
    szStatus = request("Status")    '編輯Status
    
    dim szFormID, szOriginalVesselListID
    szFormID = request ("FormID")   '單號
    szOriginalVesselListID = request ("OriginalVesselListID")   '原航次
    
    dim szVesselListID    
    szVesselListID = request ("VesselID")            '新航次
    
    dim nDataCounter, nSN(50), i
    
    nDataCounter = request ("DataCounter")
    if nDataCounter = "" then
        nDataCounter = 0
    end if
    
    
    '檢查新航次是否已核定該倉單
    sql = "select ID from StoreSum where ID='" + szFormID + "' and VesselListID = '" + szVesselListID + "'" 
    set rs = conn.execute(sql)
    
    if rs.eof then
    
        '讀出倉單筆數
        dim nOldFormCount, nNewFormCount
        nOldFormCount = 0
        nNewFormCount = 0
        
        '原航次的倉單筆數
        sql = "select DataCount from VesselList where ID = '" + szOriginalVesselListID + "'"
        set rs = conn.execute(sql)
        if not rs.eof then
            nOldFormCount = rs("DataCount")
        end if
        
        '新航次的倉單筆數
        sql = "select DataCount from VesselList where ID = '" + szVesselListID + "'"
        set rs = conn.execute(sql)
        if not rs.eof then
            nNewFormCount = rs("DataCount")
        end if
        
        '移動倉單
        for i = 0 to nDataCounter-1
            nSN(i) = request("SN_" + CStr(i))
            sql = "update FreightForm set VesselID = '" + szVesselListID + "' where SN = " + nSN(i)
               
            conn.execute(sql)
            
            nOldFormCount = nOldFormCount - 1
            nNewFormCount = nNewFormCount + 1
        next
  
        '寫入原航次的倉單筆數 (累加)        
        sql = "update VesselList set DataCount = " + CStr(nOldFormCount) + " where ID = '" + szOriginalVesselListID + "'"
        conn.execute(sql)
        
        '寫入新航次的倉單筆數 (累加)        
        sql = "update VesselList set DataCount = " + CStr(nNewFormCount) + " where ID = '" + szVesselListID + "'"
        conn.execute(sql)
  
        rs.close
        conn.close  
        set rs=nothing
        set conn=nothing
        response.redirect "FFfun02b.asp?Status=FirstLoad&VesselListID=" + szOriginalVesselListID
    else
        '已核定, 回上一頁
        
        rs.close
        conn.close  
        set rs=nothing
        set conn=nothing
        response.write("<font size=6><br><meta http-equiv=""refresh"" content=""2; url=FFfun02b.asp?Status=FirstLoad&VesselListID=" + szOriginalVesselListID +"""><center>無法移動,此單號在新航次中已核定過</center>")
    end if
%>