<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@TWKIL-HM") < 0 then
		response.redirect "../Login/Login.asp"
	end if


    '===========Parse 單號===================
    dim szStartIDTmp(100), szEndIDTmp(100), nCount, nPos
    
    function ParseFormID()
        
        nCount = 0
        
        'szSubStr = szFormID
        nPos = Instr(szSubStr, ",")
        
        while (nPos <> 0)
            szStartIDTmp(nCount) = LTrim(Left(szSubStr, nPos - 1))
            szSubStr = LTrim(Mid (szSubStr, nPos + 1, Len(szSubStr))) 
            
            '找 "-"
            nPos = Instr(szStartIDTmp(nCount), "-")
            if nPos = 0 then
                szEndIDTmp(nCount) = szStartIDTmp(nCount)            
            else
                szEndIDTmp(nCount) = LTrim(Mid (szStartIDTmp(nCount), nPos + 1, Len(szStartIDTmp(nCount))))
                szStartIDTmp(nCount) = LTrim(Left(szStartIDTmp(nCount), nPos - 1))            
            end if
              
            nCount = nCount + 1
            nPos = Instr(szSubStr, ",")
            
        wend
        
        szStartIDTmp(nCount) = LTrim(szSubStr)
        '找 "-"
        nPos = Instr(szStartIDTmp(nCount), "-")
        if nPos = 0 then
            szEndIDTmp(nCount) = szStartIDTmp(nCount)            
        else
            szEndIDTmp(nCount) = LTrim(Mid (szStartIDTmp(nCount), nPos + 1, Len(szStartIDTmp(nCount))))
            szStartIDTmp(nCount) = LTrim(Left(szStartIDTmp(nCount), nPos - 1))            
        end if
        
    end function
     
%>