<!-- #include file = ../GlobalSet/conn.asp -->
<!--攬貨商-單號對照資料-->

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@KAOINS7") < 0 then
		response.redirect "../Login/Login.asp"
	end if

function CheckInput (tStart,tEnd)
	DIM A(3),B(3),C(3) ' A for char type of tStart, B for tEnd, differ for diffs of tStart,tEnd
	DIM DifferCounts,CountType,NumbersA,NumbersB,CheckError
	CheckError = False	'check數字被分開
	DifferCounts=0
	CountType=0	'default為全數字
	NumbersA=0
	NumbersB=0
'response.write cstr((Mid(tStart,1,1) <> Mid(nEndIDTmp ,1,1)))&"<BR>"		
'response.write cstr((Mid(tStart,2,1) <> Mid(nEndIDTmp ,2,1)))&"<BR>"		
'response.write cstr((Mid(tStart,3,1) <> Mid(nEndIDTmp ,3,1)))&"<BR>"		
'response.write cstr((Mid(tStart,4,1) <> Mid(nEndIDTmp ,4,1)))&"<BR>"		
	for j=0 to 3
'response.write "nStartIDTmp="&Mid(tStart ,j+1,1)&"@<BR>"
'response.write "nEndIDTmp="&Mid(tEnd ,j+1,1)&"@<BR>"	
		if IsNumeric(Mid(tStart,j+1,1)) then
			A(j)=True
			NumbersA=NumbersA+1
		end if
		if IsNumeric(Mid(tEnd,j+1,1)) then
			B(j)=True
			NumbersB=NumbersB+1
		end if		
		C(j)=0
		if(Mid(tStart,j+1,1) <> Mid(nEndIDTmp ,j+1,1))then
			C(j)=1
			DifferCounts=DifferCounts+1	
		end if
	next	  
'response.write "DifferCounts="&DifferCounts&"<BR>"	
'response.write "NumbersA="&NumbersA&",NumbersB="&NumbersB&"<BR>"
  if ((NumbersA=4) and (NumbersB=4)) then '全數字
  	CheckInput=0  	
'--------------判斷計數的方式	
	elseif(DifferCounts=1) then'僅一位不同,判斷不同的是文字OR數字
		for j=0 to 3
			if(C(j)<>0) then
				if (not A(j)) AND (not B (j))	then'都是文字 用A-Z,Type 1
					CountType=1
				elseif (A(j) AND B(j)) then	'都是數字,Type 2
					CountType=2
				else 
			  	response.write "錯誤,"&tStart&"-"&tEnd&"中相異字元各為英數字"
			  	response.end
    			response.redirect "FFfun03b.asp?Status=Add&VesselLine=" + szVesselLine					
				end if				
			end if
		next
				
	elseif (DifferCounts>1) then'超過一位不同,不同的必需都是數字
		for j=0 to 3
			if (C(j)<>0) then
				if not(A(j)AND B(j)) then 
			  	response.write "錯誤,"&tStart&"-"&tEnd&"中相異二字元以上而相異字元不全為數字"         					
			  	response.end                                                          					
    			response.redirect "FFfun03b.asp?Status=Add&VesselLine=" + szVesselLine				
    		end if
			end if			
			'-------------檢查數字被隔開         
			if NumbersA=1 then
				if (A(1) OR A(2)) then'case 0x00 or 00x0
					CheckError = True
				end if
			end if
			if NumbersB=1 then
				if (B(1) OR B(2)) then'case 0x00 or 00x0
					CheckError = True
				end if
			end if
			if NumbersA=2 then
				if ((A(1)AND A(2)) OR (A(0)AND A(2)) OR(A(1)AND A(3))) then 'case 0xx0,x0x0,0x0x
					CheckError = True
				end if	
			end if
			if NumbersB=2 then
				if ((B(1)AND B(2)) OR (B(0)AND B(2)) OR(B(1)AND B(3))) then 'case 0xx0,x0x0,0x0x
					CheckError = True
				end if	
			end if	
			if CheckError then
						response.write "錯誤,"&tStart&"-"&tEnd&"中數字被隔開"
						response.end
			    	response.redirect "FFfun03b.asp?Status=Add&VesselLine=" + szVesselLine		
			end if  						
			CountType=2		'以數字計數
		next
	end if	
'response.write "CheckInput End <br>"	
CheckInput=CountType
end function

'---------------function addone, add one for either 純數字or not
function AddOne(Num,tStart,tEnd,cType)
  DIM TEMP(3)
	Dim Index
	WithDigital=False
	Num=ucase(Num)
'response.write "NUM="&Num&"<BR>"
	if cType=0 then	'全數字
		dim TempStr
		TempStr=""
		for i=len(cint(Num)+1) to 3 '
			TempStr=TempStr+"0"			
		next
		TempStr=TempStr+cStr(cint(Num)+1)
'response.write "TempStr="&TempStr&"<BR>"		
		AddOne=TempStr
	elseif cType=1 then	'有英文動文字
		for pt=0 to 3	'找不同的位置
			TEMP(pt)=Asc(Mid(Num ,pt+1,1))
			if (Mid(tStart,pt+1,1) <> Mid(nEndIDTmp ,pt+1,1)) then
				Index=pt
			end if
		next
		Temp(Index)=Temp(Index)+1
		AddOne=chr(TEMP(0))+chr(TEMP(1))+chr(TEMP(2))+chr(TEMP(3))
	elseif cType=2 then	'有英文 動數字
		DIM Num2,NumLen,TempString,NumberStart,NumberEnd,Digit,PassDigit
		NumberStart=9
		NumberEnd=9
		Num2=0
		NumLen=0
		Index=9
		Digit=0
		PassDigit=0
		
		for pt=4 to 1	step -1 '找數字
			if IsNumeric(Mid(Num,pt,1)) then
			  if 	NumberEnd =9 then 'first fonud
					NumberStart=pt
					NumberEnd=pt
					PassDigit=Digit
				end if				
				if 	(pt < 	NumberStart) AND (PassDigit=Digit) then NumberStart=pt
			else
				Digit=1
			end if
		next	
'response.write "NumberStart="&NumberStart&"<BR>"
'response.write "NumberEnd="&NumberEnd&"<BR>"			
		
		for	pt=NumberStart to NumberEnd
			Num2=Num2*10+cint(mid(Num,pt,1))
		next	
		Num2=Num2+1
'response.write "Num2="&Num2&"<BR>"		
		StrNum = cstr(Num2)
		for i= len(Num2) to (NumberEnd-NumberStart) '補零
			StrNum="0"+StrNum
		next
		
		TempString=""	
'response.write "--StrNum="&StrNum&"<BR>"		
'response.write "NumLen="&(NumberEnd-NumberStart)+1&"<BR>"
'response.write TempString&" A<BR>"	

		for pt=1 to NumberStart -1	'加頭
			TempString=TempString+Mid(Num,pt,1)
'			response.write TempString&" --B<BR>"		
		next
'response.write TempString&" B<BR>"		

		'-----------------加數字
		TempString=TempString+StrNum
'response.write TempString&" C<BR>"	
		
		for pt=NumberEnd+1 to 4 '加尾
			TempString=TempString+Mid(Num,pt,1)
		next
'response.write "AddOne="&TempString&"    AddOne End<BR>"	
		AddOne=TempString
	end if	'end of CType=2
 
end function



    '===============接收參數===============
    dim szStatus
    szStatus = request("Status")    '編輯Status
    
    dim szFormID, szOwnerID, szVesselLine
    
    if szStatus = "Modify" then        '單號
        szFormID = request("FormIDToModify")
    else
        szFormID = request("FormID")        
    end if
          
    szOwnerID   = request("OwnerID")      '攬貨商
    
    '21-Nov2004: 新增航線欄位, 因為不同航線可能會有同樣的單號
    szVesselLine = request ("VesselLine") '航線
    
    
    '===========Parse 單號===================
    '21-Nov2004: buffer的size不要太小, 如此才能容納大筆的單號
    dim szFormIDTmp(10000), nCount, nPos, szSubStr, i, nStartIDTmp, nEndIDTmp, nFormIDTmp
    nCount = 0
    
    szSubStr = Ucase(szFormID)
    nPos = Instr(szSubStr, ",")
    dim nStrLen    
'根據, 分成多項
    if nPos=0 then '沒有, 分開
      szFormIDTmp(nCount)=szSubStr
      'response.write "single case ="&szFormIDTmp(nCount)&"<BR>"
    else
      while (nPos <> 0)
        szFormIDTmp(nCount) = LTrim(Left(szSubStr, nPos - 1))
        szSubStr = LTrim(Mid (szSubStr, nPos + 1, Len(szSubStr))) 

        nPos = Instr(szSubStr, ",")
        nCount=nCount+1
      wend   
    end if

    dim szPreStartFormID, szPreEndFormID, bGotNumberPart, nPreStrLen
    szPreStartFormID = ""
    szPreEndFormID = ""
    szFormIDTmp(nCount) = LTrim(szSubStr)

  Dim ItemCount 
  Dim StartID(3),EndID(3),TEMP(3),TEMP2(3)
  ItemCount = nCount
  for i=0 to ItemCount	'對每一個item做拆解
    Dim WithDigital
    WithDigital = False
    nPos = Instr(szFormIDTmp(i), "-")			'找 "-"

    if nPos<>0 then 					'需拆解
     ' nCount=nCount+1
      nStartIDTmp = LTrim(Left(szFormIDTmp(i), nPos - 1))
      nEndIDTmp = LTrim(Mid(szFormIDTmp(i), nPos + 1, Len(szFormIDTmp(i))))
response.write "nStartIDTmp="+nStartIDTmp+"<BR>"
response.write "nEndIDTmp="+nEndIDTmp+"<BR>"
        
      for j=0 to 3
        StartID(j)=ASC(Mid(nStartIDTmp,j+1,1))
 		    EndID(j)=ASC(Mid(nEndIDTmp ,j+1,1))
      next
DIM CType
CType=CheckInput (nStartIDTmp,nEndIDTmp)
response.write "CheckInput="&CType&"<BR>"
response.write "Addone="&AddOne(nStartIDTmp,nStartIDTmp,nEndIDTmp,cType)&"<BR>"
'response.end

   'write into buffer one by one
          for k=0 to 999
             if nStartIDTmp<> AddOne(nEndIDTmp,nStartIDTmp,nEndIDTmp,cType) then 'Start/End 相反
             		 response.write "Start/End change<br>"
                 ncount=ncount+1             
                 szFormIDTmp(nCount) =nStartIDTmp
                 nStartIDTmp=AddOne(nStartIDTmp,nStartIDTmp,nEndIDTmp,cType)
                 'response.write "Ncount="&nCount&"非純數字 szFormIDTmp="&szFormIDTmp(nCount)&"<BR>"        	                           
             end if
          next         
   
response.write "ncount="&nCount&"<BR>"
    else						'single case, just copy again
      nCount=nCount+1
      szFormIDTmp(nCount)=szFormIDTmp(i)
    end if   
  next

    '===============寫入資料庫===============  
    dim sql, rs 
    set rs = nothing 
response.write "szStatus="&szStatus&"<BR>"
    '=========新增=========
    if szStatus = "Add" then 
response.write "ItemCount="&ItemCount&"<BR>"
        for i = ItemCount+1 to nCount
            '檢查編號是否重複  
response.write "i="&i&"szFormIDTmp="&szFormIDTmp(3)&"<BR>"	
response.write "i="&i&"SQL="&sql&"<BR>"
'response.end           
            sql = "select FormID from FormToOwner where FormID='" + szFormIDTmp(i) + "' and VesselLine = '" + szVesselLine + "'"             
            set rs = conn.execute(sql)    
            
            if rs.eof then
                sql = "insert into FormToOwner(FormID, OwnerID, VesselLine) values('" + szFormIDTmp(i) + "','" + szOwnerID + "'"
                sql = sql + ",'" + szVesselLine + "')"  
'response.write "i="&i&"SQL="&sql&"<BR>"		             
                conn.execute(sql)                
            else
                '編號重複, 回上一頁
                'set rs=nothing
                'set conn=nothing
                'response.write("<meta http-equiv=""refresh"" content=""2; url=FFfun03b.asp""><center>編號重複</center>")    
            end if
        next
       
        rs.close
        conn.close 
        set rs=nothing
        set conn=nothing
'response.end        
        response.redirect "FFfun03b.asp?Status=Add&VesselLine=" + szVesselLine
        
    '=========修改=========
    elseif szStatus = "Modify" then
    
        set rs = nothing 
      	sql = "update FormToOwner set OwnerID = '"+ szOwnerID 
      	sql = sql + "' where FormID = '"+ szFormID +"' and VesselLine = '" + szVesselLine + "'"
      	
      	conn.execute(sql)
      	
      	conn.close 
        set conn=nothing
        response.redirect "FFfun03b.asp?Status=Add&VesselLine=" + szVesselLine
    
    '=========刪除=========
    elseif szStatus = "Delete" then
      
      	for i = ItemCount+1 to nCount
          	sql = "delete from FormToOwner where FormID ='" + szFormIDTmp(i) + "' and VesselLine = '" + szVesselLine + "'"          	
          	if szOwnerID <> "0" then
          	    sql = sql + " and OwnerID = '" + szOwnerID + "'"
          	end if        
	
          	conn.execute(sql)
      	next
      	
      	conn.close 
        set conn=nothing
        response.redirect "FFfun03b.asp?Status=Add&VesselLine=" + szVesselLine
        
    end if
%>