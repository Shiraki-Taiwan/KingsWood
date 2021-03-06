<!-- #include file = ../GlobalSet/conn.asp -->
<!--航次資料刪除-->

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@KEEINS8") < 0 then
		response.redirect "../Login/Login.asp"
	end if



    '===============接收參數===============
    dim szStatus
    szStatus = request("Status")    '編輯Status
    
    dim szID, szVesselName, szVesselNo, szCheckInID, nYear, nMonth, nDay, szOwner, szVesselLine
      
    szID = request("ID")        '編號 
    
    szVesselName    = request("VesselName")       '船名
    szOwner         = request("Owner")            '船公司
    szVesselNo      = request("VesselNo")         '航次
    szCheckInID     = request("CheckInID")        '船掛
    nYear           = request("Year")             '結關年
    nMonth          = request("Month")            '結關月
    nDay            = request("Day")              '結關日
    
    '21-Nov2004: 加入航線
    szVesselLine    = request("VesselLine")       '航線
    
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
    szDate = CStr(int(nYear)) + "/" + CStr(int(nMonth)) + "/" + CStr(int(nDay))     
    
    '===============寫入資料庫===============  
    dim sql, rs 
    set rs = nothing 
  
    '=========新增=========
    if szStatus = "Add" then 
        '檢查編號是否重複  
        sql = "select ID from VesselList where VesselName='" + szVesselName + "'"
        sql = sql + " and Owner='" + szOwner + "' and VesselNo='" + szVesselNo + "' and CheckInID='" + szCheckInID + "'"
        
        '07-Apr2005: 若沒有日期, 就不要加入查詢指令, 否則會格式不合
        if szDate <> "0/0/0" then
	       	sql = sql + " and Date=#" + szDate + "#"
        end if

		sql = sql + " and VesselLine='" + szVesselLine+ "'"
        
        set rs = conn.execute(sql)    
        
        if rs.eof then
            
            if szDate = "0/0/0" then
                sql = "insert into VesselList (ID, VesselName, Owner, VesselNo, CheckInID, VesselLine)"            
                sql = sql + " values('" & szID & "','" & szVesselName & "','" & szOwner & "','" & szVesselNo & "','" & szCheckInID & "','" & szVesselLine & "')"
            else
                sql = "insert into VesselList" &_
              " values('" & szID & "','" & szVesselName & "','" & szOwner & "','" & szVesselNo & "','" & szCheckInID & "','" & szDate & "', 0,'', '" & szVesselLine & "')"
            end if
            
            
            conn.execute(sql)
            
            rs.close
            conn.close 
            set rs=nothing
            set conn=nothing
            response.redirect "VOfun04a.asp"
        else
            '編號重複, 回上一頁
            
            rs.close
            conn.close 
            set rs=nothing
            set conn=nothing
            response.write("<font size=6><br><meta http-equiv=""refresh"" content=""2; url=VOfun04a.asp""><center>資料重複</center>")    
        end if
        
    '=========修改=========
    elseif szStatus = "Modify" then
        set rs = nothing 
      	'sql = "update VesselList set VesselName='"+ szVesselName + "', Owner = '" + szOwner + "', VesselNo='" + szVesselNo + "', CheckInID='" + szCheckInID + "' where ID = '"+ szID +"'"
      	
      	
      	'因為目前不知如何update日期，所以暫時先刪除，再新增修改過的資料
      	dim nDataCount
      	sql = "select DataCount from VesselList where ID ='" + szID + "'"
      	set rs = conn.execute(sql)
      	if not rs.eof then
      	    nDataCount = rs("DataCount")
      	else
      	    nDataCount = 0
      	end if
      	
      	sql = "delete from VesselList where ID ='" + szID + "'"
      	conn.execute(sql)
      	
      	if szDate = "0/0/0" then
            sql = "insert into VesselList (ID, VesselName, Owner, VesselNo, CheckInID, VesselLine, DataCount)"            
            sql = sql + " values('" & szID & "','" & szVesselName & "','" & szOwner & "','" & szVesselNo & "','" & szCheckInID & "','" + szVesselLine + "'," + CStr(nDataCount) + ")"
        else
            sql = "insert into VesselList" &_
          " values('" & szID & "','" & szVesselName & "','" & szOwner & "','" & szVesselNo & "','" & szCheckInID & "','" & szDate & "'," & CStr(nDataCount) &", '', '" + szVesselLine + "')"
        end if
         
        conn.execute(sql)
      	
      	rs.close
        conn.close 
      	set rs=nothing
        set conn=nothing
        response.redirect "VOfun04a.asp"
        
    '=========刪除=========
    elseif szStatus = "Delete" then
        sql = "delete from VesselList where ID ='" + szID + "'"
      	conn.execute(sql)
      	
      	conn.close 
        set conn=nothing
        response.redirect "VOfun04a.asp"
        
    '=========結關=========
    elseif szStatus = "Close" then
        dim sz
        sql = "select * from VesselList where ID ='" + szID + "'"
        set rs = conn.execute(sql)
        
        if not rs.eof then
            szOwner = rs("Owner")
            szCheckInID = rs("CheckInID")
        else
            rs.close
            conn.close 
            set rs=nothing
            set conn=nothing
            response.redirect "VOfun04a.asp"       
        end if
        
        Set objFileSystem = CreateObject ("Scripting.FileSystemObject")
        
        dim szFilePath, szFileName
        szFilePath = Request.ServerVariables ("PATH_TRANSLATED")
        
        nPos = InStrRev (szFilePath, "\")
        
        szFilePath = Mid (szFilePath, 1, nPos)
        
        szFilePath = szFilePath + "TextFiles\"
        
        '21-Feb2005: 開新資料夾
        szFilePath = szFilePath & szID
        
        if not (objFileSystem.FolderExists(szFilePath)) then
            objFileSystem.CreateFolder(szFilePath)
        end if
        
        szFilePath = szFilePath + "\"
        
        '==========================================================
        '寫txt檔: 第一種格式
        
        szFileName = szFilePath & szID & ".txt"
        
        if (objFileSystem.FileExists(szFileName)) then
            objFileSystem.DeleteFile(szFileName)
        end if
        
        Set objWritedTextFile = objFileSystem.CreateTextFile(szFileName, -1, 0)
        
        dim szStrTmp, nStrLen
        
        dim i, szPreStr
        
        '接船公司: 5位
        nStrLen = Len(szOwner)
        szPreStr = ""
        for i = nStrLen+1 to 5
            szPreStr = szPreStr & "0"
        next
        
        if nStrLen > 5 then
            szOwner = Mid(szOwner, 1, 5)
        else
            szOwner = szPreStr & szOwner
        end if
        
        '接船掛: 6位
        szPreStr = ""
        nStrLen = Len(szCheckInID)
        for i = nStrLen+1 to 6
            szPreStr = szPreStr & "0"
        next
        
        if nStrLen > 6 then
            szCheckInID = Mid(szCheckInID, 1, 6)
        else
            szCheckInID = szPreStr & szCheckInID
        end if
        
        
        '04-Jan2005: 只列出已核對的部分
        sql = "select ID, IsChecked, sum(Piece) as TotalPiece, sum(Volume) as TotalVolume, sum(Weight) as TotalWeight"
        sql = sql + " from FreightForm where VesselID = '" + szID + "' and IsChecked = true"
        
        sql = sql + " group by ID, IsChecked order by ID"
        
        set rs = conn.execute(sql)        
        
        '05-Mar2005: Patch for checking the 2nd rs.eof
        dim bRsEof        
        if not rs.eof then
            bRsEof = 0
        else
            bRsEof = 1  
        end if
        
        dim fForestry, szForestryPre, szForestryTail, nVolume, nPiece, nWeight
        while not rs.eof
            nVolume = rs("TotalVolume")
            nPiece = rs("TotalPiece")
            nWeight = rs("TotalWeight")
            
            
            '起始碼 + 5位船公司 + 6位船掛
            szStrTmp = "1" & szOwner & szCheckInID
            
            '貨號: 4位
            nStrLen = Len(rs("ID"))
            szPreStr = ""
            for i = nStrLen+1 to 4
                szPreStr = szPreStr & "0"
            next
        
            if nStrLen > 6 then
                szStrTmp = szStrTmp & Mid(CStr(rs("ID")), 1, 4)
            else
                szStrTmp = szStrTmp & szPreStr & CStr(rs("ID"))
            end if
            
            '總件數: 5位
            nStrLen = Len(nPiece)
            szPreStr = ""
            for i = nStrLen+1 to 5
                szPreStr = szPreStr & "0"
            next
        
            if nStrLen > 5 then
                szStrTmp = szStrTmp & Mid(CStr(nPiece), 1, 4)
            else
                szStrTmp = szStrTmp & szPreStr & CStr(nPiece)
            end if
            
            '總才積: 4.4位
            '05-Mar2005:改用立方公尺為單位, 不再換算成才積
            'fForestry = FormatNumber(rs("TotalVolume") * 35.3445, 2)
            fForestry = FormatNumber(rs("TotalVolume"), 2)
            '拿掉format後產生的逗號
            fForestry = Replace(fForestry, ",", "")
            
            nPos = InStrRev (CStr(fForestry), ".")            
            szForestryPre = Mid(CStr(fForestry), 1, nPos-1)
            
            '小數點前4位
            nStrLen = Len(szForestryPre)
            szPreStr = ""
            
            for i = nStrLen+1 to 4
                szPreStr = szPreStr & "0"
            next
            
            if nStrLen > 4 then
                szStrTmp = szStrTmp & Mid(CStr(szForestryPre), 1, 4)
            else
                szStrTmp = szStrTmp & szPreStr & CStr(szForestryPre)
            end if
            
            '小數點後4位
            szForestryTail = Mid(CStr(fForestry), nPos+1)
            nStrLen = Len(szForestryTail)
            szPreStr = ""
            
            for i = nStrLen+1 to 4
                szPreStr = szPreStr & "0"
            next
            
            if nStrLen > 4 then
                szStrTmp = szStrTmp & Mid(CStr(szForestryTail), 1, 4)
            else
                szStrTmp = szStrTmp & CStr(szForestryTail) & szPreStr
            end if
            
            
            '總重量: 5位
            nStrLen = Len(nWeight)
            szPreStr = ""
            for i = nStrLen+1 to 5
                szPreStr = szPreStr & "0"
            next
        
            if nStrLen > 5 then
                szStrTmp = szStrTmp & Mid(CStr(nWeight), 1, 4)
            else
                szStrTmp = szStrTmp & szPreStr & CStr(nWeight)
            end if
            
            
            '結束碼
            szStrTmp = szStrTmp & "0"
        
            objWritedTextFile.WriteLine szStrTmp
            rs.movenext
        wend
        
        objWritedTextFile.Close
        
        
        '==========================================================
        '21-Feb2005: 輸出第二種格式的文字檔
        'Format: 1,S/O號碼,件數,體積,0
        dim szFormID
        
        szFileName = szFilePath & "ZACT.txt"
        
        if (objFileSystem.FileExists(szFileName)) then
            objFileSystem.DeleteFile(szFileName)
        end if
        
        Set objWritedTextFile = objFileSystem.CreateTextFile(szFileName, -1, 0)
        
        if bRsEof = 0 then
            rs.movefirst
        end if
        
        while not rs.eof
        	nVolume = FormatNumber(rs("TotalVolume"), 2)
            nPiece = rs("TotalPiece")
            szFormID = rs("ID")
                        
            szStrTmp = "1," & szFormID & "," & nPiece & "," & nVolume & ",0"
            objWritedTextFile.WriteLine szStrTmp
            
        	rs.movenext
        wend
        
        objWritedTextFile.Close
        
        '========================================================== 
        '25-May2005: 輸出第三種格式的文字檔
        '僅列出已核對的部分
        szFileName = szFilePath & "DAI_1.txt"
        if (objFileSystem.FileExists(szFileName)) then
            objFileSystem.DeleteFile(szFileName)
        end if
        
        Set objWritedTextFile = objFileSystem.CreateTextFile(szFileName, -1, 0)
        
        if bRsEof = 0 then
            rs.movefirst
        end if
        
        
        while not rs.eof
        	szStrTmp = RIGHT (Space(5) & Trim(rs("ID")),5)
            szStrTmp = szStrTmp & RIGHT (Space(6) & Trim(rs("TotalPiece")),6)
        	szStrTmp = szStrTmp & RIGHT (Space(8) & Trim(FormatNumber(rs("TotalVolume"), 2)),8)
        	szStrTmp = szStrTmp & RIGHT (Space(8) & Trim(FormatNumber(rs("TotalWeight"), 2)),8)
        	szStrTmp = szStrTmp & RIGHT (Space(8) & Trim(szVesselNo),8)
        	szStrTmp = szStrTmp & RIGHT (Space(8) & Trim(szVesselName),8)
        	
        	szDate = CStr(int(nYear)) & "." & CStr(RIGHT("0" & int(nMonth), 2)) & "." & CStr(RIGHT("0" & int(nDay), 2)) 
        	szStrTmp = szStrTmp & RIGHT (Space(15) & Trim(szDate),15)
        	
            objWritedTextFile.WriteLine szStrTmp
            
        	rs.movenext
        wend
        
        objWritedTextFile.Close
        
        '========================================================== 
        '25-May2005: 輸出第四種格式的文字檔
        '全部資料(已核對+未核對)
        szFileName = szFilePath & "DAI_2.txt"
        if (objFileSystem.FileExists(szFileName)) then
            objFileSystem.DeleteFile(szFileName)
        end if
        
        Set objWritedTextFile = objFileSystem.CreateTextFile(szFileName, -1, 0)
        
        sql = "select ID, sum(Piece) as TotalPiece, sum(Volume) as TotalVolume, sum(Weight) as TotalWeight"
        sql = sql + " from FreightForm where VesselID = '" + szID + "'"
        
        sql = sql + " group by ID order by ID"
        
        set rs = conn.execute(sql)      
        
        while not rs.eof
        	szStrTmp = RIGHT (Space(5) & Trim(rs("ID")),5)
            szStrTmp = szStrTmp & RIGHT (Space(6) & Trim(rs("TotalPiece")),6)
        	szStrTmp = szStrTmp & RIGHT (Space(8) & Trim(FormatNumber(rs("TotalVolume"), 2)),8)
        	szStrTmp = szStrTmp & RIGHT (Space(8) & Trim(FormatNumber(rs("TotalWeight"), 2)),8)
        	szStrTmp = szStrTmp & RIGHT (Space(8) & Trim(szVesselNo),8)
        	szStrTmp = szStrTmp & RIGHT (Space(8) & Trim(szVesselName),8)
        	
        	szDate = CStr(int(nYear)) & "." & CStr(RIGHT("0" & int(nMonth), 2)) & "." & CStr(RIGHT("0" & int(nDay), 2)) 
        	szStrTmp = szStrTmp & RIGHT (Space(15) & Trim(szDate),15)
        	
            objWritedTextFile.WriteLine szStrTmp
            
        	rs.movenext
        wend
        
        objWritedTextFile.Close        
        '==========================================================
        rs.close
        conn.close 
        set rs=nothing
        set conn=nothing
        response.redirect "VOfun04a.asp?Status=Modify&ID=" + szID
        
    end if
%>