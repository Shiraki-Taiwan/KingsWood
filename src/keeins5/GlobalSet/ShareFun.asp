<%    
    '將字串轉成n位
    function FormatString (szStr, nDesiredLen)
    	
    	dim nStrLen
    	nStrLen = Len(szStr)
    	
    	dim i, szNewStr
    	for i = nStrLen to CLng(nDesiredLen) - 1
    		szNewStr = szNewStr & "0"
    	next
    	
    	szNewStr = szNewStr & szStr
    	
    	FormatString = szNewStr
    	
    end function
        
    function GetVesselListIDBySerialNum(a_nSerialNum)
    	dim sql, rs
    	set rs = nothing
            
        sql = request ("Sql")
        
        if sql = "" then    
            sql = "select * from VesselList order by Date DESC, DataCount DESC ,ID DESC"
        end if
        
        set rs = conn.execute(sql)
        
        dim nCounter
        nCounter = 1
        
        
        while not rs.eof
            if nCounter = a_nSerialNum then
                szVesselListID = rs("ID")
                szVesselLine = rs("VesselLine")
            end if
            
            nCounter = nCounter + 1
            rs.movenext        
        wend
        
        rs.close
        conn.close 
        set rs=nothing
        set conn=nothing
        
    end function
    
    
    function IsIrregularVolume(nBoard, fVolume)
    	dim fVolumeTmp
    	fVolumeTmp = fVolume
    	
    	
    	'板數大於1, 體積先除板數, 稍後再看是否體積仍超大或超小
    	if nBoard > 1 then	    
		    fVolumeTmp = fVolumeTmp / nBoard
		end if
            
        if fVolumeTmp>3.8 OR fVolumeTmp < 0.1 then
            IsIrregularVolume = 1
        else
            IsIrregularVolume = 0
        end if    
    end function
    
    
    '27-Apr2005: 若有堆量, 且長寬高小於30cm, 變色    
    function IsIrregularLength(bIsPL, fLength)   
    	If fLength > 600 OR fLength < 10 OR (bIsPL = TRUE And fLength < 30) Then
    	    IsIrregularLength = 1
        else
        	IsIrregularLength = 0
        end if
    end function
    
    function IsIrregularHeight(bIsPL, fHeight)
    	if fHeight > 226 OR fHeight < 10  OR (bIsPL = TRUE And fHeight < 30)  Then
            IsIrregularHeight = 1
        else
        	IsIrregularHeight = 0
        end if
    end function
    
    function IsIrregularWidth(bIsPL, fWidth)
    	if fWidth > 600 OR fWidth < 10  OR (bIsPL = TRUE And fWidth < 30)  then
            IsIrregularWidth = 1
        else
        	IsIrregularWidth = 0
        end if
    end function
    
        
    '17-Jun2005: 用來計算體積
    function VolumeCalculator(a_nBoard, a_bIsPL, a_nPiece, a_fLength, a_fWidth, a_fWeight)
	    Dim fVolumn
	    Dim DivNum
	    DivNum = 1000000
	    fVolumn = 0
	    
	    
	    Dim nBoard, nPiece, nLength, nWidth, nHeight, bIsPL
	    nBoard = a_nBoard
	    If nBoard = 0 Then
	        nBoard = 1
	    End If
	    
	    bIsPL = a_bIsPL
	    
	    nPiece = a_nPiece
	    If nPiece = 0 Then
	        nPiece = 1
	    End If 
	       
	    nLength = a_fLength
	    nWidth = a_fWidth
	    nHeight = a_fWeight
	    
	    If Trim(nLength) <> "" And Trim(nWidth) <> "" And Trim(nHeight) <> "" Then
	        fVolumn = nLength * nWidth * nHeight / DivNum
	        
	        '不是堆量，要乘件數
	        If bIsPL = "0" Then
	            fVolumn = fVolumn * nPiece
	        Else    '堆量, 要乘板數
	            fVolumn = fVolumn * nBoard
	        End If
	        
	        fVolumn = FormatNumber (fVolumn, 2)
    	End If
    	
    	VolumeCalculator = fVolumn
   	end function		    
   	
   	
   	
    '17-Jun2005: 處理字串中的特殊字元
    function ParseSpecialCharInAString (a_szStr)
	    dim szSubStr, iIndex, szTmpStr, szResult
	    szTmpStr = ""
	    szSubStr = a_szStr
	    
	    '處理單引號
	    iIndex = Instr(szSubStr, "'")
	    while iIndex <> 0
	        szTmpStr = szTmpStr + Left(szSubStr, iIndex - 1) + "''"
	        szSubStr = Mid(szSubStr, iIndex + 1, Len(szSubStr))
	        iIndex = Instr(szSubStr, "'")       
	    wend
	    
	    szResult = szTmpStr + szSubStr
	    
	    '處理雙引號
	    szSubStr = szResult	    
	    iIndex = Instr(szSubStr, """")
	    while iIndex <> 0
	        szTmpStr = szTmpStr + Left(szSubStr, iIndex - 1) + "''"
	        szSubStr = Mid(szSubStr, iIndex + 1, Len(szSubStr))
	        iIndex = Instr(szSubStr, """")       
	    wend
	    
	    ParseSpecialCharInAString = szResult
	 end function
	 
	 '---------------- Format 數字格式--------------
	Function MyFormatNumber(a_num, a_dotnum)
	
	    dim nResult, nTmpStr
	
		nResult = FormatNumber(a_num, a_dotnum)
		nTmpStr = CStr(nResult)
		
		nResult = Replace (nTmpStr, ",", "")
		
		MyFormatNumber = nResult
	
	End Function

    Function IIf( expr, truepart, falsepart )
		IIf = falsepart
		If expr Then IIf = truepart
	End Function
%>