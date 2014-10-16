Dim nFieldDiff
nFieldDiff = 15
    
    
Function ChangeTextColor(a_nIndex, nLength, nWidth, nHeight, fVolume, bIsPL)
    
    Dim ColorYellow, ColorWhite, ColorPink
    ColorYellow = "#FFFFA0"
    ColorWhite = "#ffffff"
    ColorPink  = "#FFC8FF"
    
    Dim bChangeBg
    bChangeBg = 0
    
      
    '若有板數,體積先除以板數
    Dim nBoard
    If document.form(5+a_nIndex*nFieldDiff).value <> "" Then
        nBoard = document.form(5+a_nIndex*nFieldDiff).value
        If IsNumeric(nBoard) Then
            nBoard = CLng(nBoard)
        Else
            nBoard = 1  'just for aborting error
        End If
        
        If nBoard > 0 Then
            fVolume = fVolume / nBoard
        End if
    End If
      
        
    '體積大於3.8, 或小於0.1 用不同頻色表示
    Dim ColorVolume
    ColorVolume = ""
    If (fVolume > 3.8) OR (fVolume < 0.1) Then
        bChangeBg = 1
        ColorVolume = ColorPink
    End If
    
    Dim ColorLength
    ColorLength = ""
    '27-Apr2005: 若有堆量, 且長寬高小於30cm, 變色
    If nLength > 600 OR nLength < 10 OR (bIsPL = 1 And nLength < 30) Then
        bChangeBg = 1
        ColorLength = ColorPink
    End If
    
    Dim ColorWidth
    ColorWidth = ""
    If nWidth > 600 OR nWidth < 10 OR (bIsPL = 1 And nWidth < 30) Then
        bChangeBg = 1
        ColorWidth = ColorPink
    End If
    
    Dim ColorHeight
    ColorHeight = ""
    If nHeight > 226 OR nHeight < 10 OR (bIsPL = 1 And nHeight < 30) Then
        bChangeBg = 1
        ColorHeight = ColorPink
    End If
    
    
    If (bChangeBg = 1) Then
        document.form(2+a_nIndex*nFieldDiff).style.backgroundColor  = ColorYellow
        document.form(3+a_nIndex*nFieldDiff).style.backgroundColor  = ColorYellow
        document.form(4+a_nIndex*nFieldDiff).style.backgroundColor  = ColorYellow
        document.form(5+a_nIndex*nFieldDiff).style.backgroundColor  = ColorYellow
        document.form(6+a_nIndex*nFieldDiff).style.backgroundColor  = ColorYellow
        document.form(7+a_nIndex*nFieldDiff).style.backgroundColor  = ColorYellow
        document.form(8+a_nIndex*nFieldDiff).style.backgroundColor  = ColorYellow
        
        If ColorLength <> "" Then
            document.form(9+a_nIndex*nFieldDiff).style.backgroundColor  = ColorLength
        Else            
            document.form(9+a_nIndex*nFieldDiff).style.backgroundColor  = ColorYellow
        End If
        
        If ColorWidth <> "" Then
            document.form(10+a_nIndex*nFieldDiff).style.backgroundColor = ColorWidth
        Else
            document.form(10+a_nIndex*nFieldDiff).style.backgroundColor = ColorYellow
        End If
        
        If ColorHeight <> "" Then
            document.form(11+a_nIndex*nFieldDiff).style.backgroundColor = ColorHeight
        Else
            document.form(11+a_nIndex*nFieldDiff).style.backgroundColor = ColorYellow
        End If
        
        If ColorVolume <> "" Then
            document.form(12+a_nIndex*nFieldDiff).style.backgroundColor = ColorPink  
        Else
            document.form(12+a_nIndex*nFieldDiff).style.backgroundColor = ColorYellow  
        End If
                          
        document.form(13+a_nIndex*nFieldDiff).style.backgroundColor = ColorYellow
        document.form(14+a_nIndex*nFieldDiff).style.backgroundColor = ColorYellow        
    Else        
        document.form(2+a_nIndex*nFieldDiff).style.backgroundColor  = ColorWhite
        document.form(3+a_nIndex*nFieldDiff).style.backgroundColor  = ColorWhite
        document.form(4+a_nIndex*nFieldDiff).style.backgroundColor  = ColorWhite
        document.form(5+a_nIndex*nFieldDiff).style.backgroundColor  = ColorWhite
        document.form(6+a_nIndex*nFieldDiff).style.backgroundColor  = ColorWhite
        document.form(7+a_nIndex*nFieldDiff).style.backgroundColor  = ColorWhite
        document.form(8+a_nIndex*nFieldDiff).style.backgroundColor  = ColorWhite
        document.form(9+a_nIndex*nFieldDiff).style.backgroundColor  = ColorWhite
        document.form(10+a_nIndex*nFieldDiff).style.backgroundColor = ColorWhite
        document.form(11+a_nIndex*nFieldDiff).style.backgroundColor = ColorWhite
        document.form(12+a_nIndex*nFieldDiff).style.backgroundColor = ColorWhite                    
        document.form(13+a_nIndex*nFieldDiff).style.backgroundColor = ColorWhite
        document.form(14+a_nIndex*nFieldDiff).style.backgroundColor = ColorWhite
    End If
    
End Function

Function ChangeTextColor2(a_nIndex, nLength, nWidth, nHeight, fVolume, bIsPL)
    
    Dim ColorYellow, ColorWhite, ColorPink
    ColorYellow = "#FFFFA0"
    ColorWhite = "#ffffff"
    ColorPink  = "#FFC8FF"
    
    Dim bChangeBg
    bChangeBg = 0
    
    
    '若有板數,體積先除以板數
    Dim nBoard
    If document.form(10+a_nIndex*nFieldDiff).value <> "" Then
        nBoard = document.form(10+a_nIndex*nFieldDiff).value
        If IsNumeric(nBoard) Then
            nBoard = CLng(nBoard)
        Else
            nBoard = 1  'just for aborting error
        End If
        
        If nBoard > 0 Then
            fVolume = fVolume / nBoard
        End if
    End If
    
    
    '體積大於3.8, 或小於0.1 用不同頻色表示
    Dim ColorVolume
    ColorVolume = ""
    If (fVolume > 3.8) OR (fVolume < 0.1) Then
        bChangeBg = 1
        ColorVolume = ColorPink
    End If
    
    Dim ColorLength
    ColorLength = ""
    '27-Apr2005: 若有堆量, 且長寬高小於30cm, 變色
    If nLength > 600 OR nLength < 10 OR (bIsPL = 1 And nLength < 30) Then
        bChangeBg = 1
        ColorLength = ColorPink
    End If
    
    Dim ColorWidth
    ColorWidth = ""
    If nWidth > 600 OR nWidth < 10  OR (bIsPL = 1 And nWidth < 30)  Then
        bChangeBg = 1
        ColorWidth = ColorPink
    End If
    
    Dim ColorHeight
    ColorHeight = ""
    If nHeight > 226 OR nHeight < 10  OR (bIsPL = 1 And nHeight < 30)  Then
        bChangeBg = 1
        ColorHeight = ColorPink
    End If
    
    If (bChangeBg = 1) Then
        document.form(7+a_nIndex*nFieldDiff).style.backgroundColor  = ColorYellow
        document.form(8+a_nIndex*nFieldDiff).style.backgroundColor  = ColorYellow
        document.form(9+a_nIndex*nFieldDiff).style.backgroundColor  = ColorYellow
        document.form(10+a_nIndex*nFieldDiff).style.backgroundColor = ColorYellow
        document.form(11+a_nIndex*nFieldDiff).style.backgroundColor = ColorYellow
        document.form(12+a_nIndex*nFieldDiff).style.backgroundColor = ColorYellow
        document.form(13+a_nIndex*nFieldDiff).style.backgroundColor = ColorYellow
        
        If ColorLength <> "" Then
            document.form(14+a_nIndex*nFieldDiff).style.backgroundColor  = ColorLength
        Else            
            document.form(14+a_nIndex*nFieldDiff).style.backgroundColor  = ColorYellow
        End If
        
        If ColorWidth <> "" Then
            document.form(15+a_nIndex*nFieldDiff).style.backgroundColor = ColorWidth
        Else
            document.form(15+a_nIndex*nFieldDiff).style.backgroundColor = ColorYellow
        End If
        
        If ColorHeight <> "" Then
            document.form(16+a_nIndex*nFieldDiff).style.backgroundColor = ColorHeight
        Else
            document.form(16+a_nIndex*nFieldDiff).style.backgroundColor = ColorYellow
        End If
        
        If ColorVolume <> "" Then
            document.form(17+a_nIndex*nFieldDiff).style.backgroundColor = ColorPink  
        Else
            document.form(17+a_nIndex*nFieldDiff).style.backgroundColor = ColorYellow  
        End If
                            
        document.form(18+a_nIndex*nFieldDiff).style.backgroundColor = ColorYellow
        document.form(19+a_nIndex*nFieldDiff).style.backgroundColor = ColorYellow        
    Else        
        document.form(7+a_nIndex*nFieldDiff).style.backgroundColor  = ColorWhite
        document.form(8+a_nIndex*nFieldDiff).style.backgroundColor  = ColorWhite
        document.form(9+a_nIndex*nFieldDiff).style.backgroundColor  = ColorWhite
        document.form(10+a_nIndex*nFieldDiff).style.backgroundColor = ColorWhite
        document.form(11+a_nIndex*nFieldDiff).style.backgroundColor = ColorWhite
        document.form(12+a_nIndex*nFieldDiff).style.backgroundColor = ColorWhite
        document.form(13+a_nIndex*nFieldDiff).style.backgroundColor = ColorWhite
        document.form(14+a_nIndex*nFieldDiff).style.backgroundColor = ColorWhite
        document.form(15+a_nIndex*nFieldDiff).style.backgroundColor = ColorWhite
        document.form(16+a_nIndex*nFieldDiff).style.backgroundColor = ColorWhite
        document.form(17+a_nIndex*nFieldDiff).style.backgroundColor = ColorWhite                    
        document.form(18+a_nIndex*nFieldDiff).style.backgroundColor = ColorWhite
        document.form(19+a_nIndex*nFieldDiff).style.backgroundColor = ColorWhite 
    End If
    
End Function

'-----------------------計算體積----------------------------------------
Function VolumeCalculator(a_nIndex)
    
    Dim fVolumn
    Dim DivNum
    DivNum = 1000000
    fVolumn = 0
    
    
    Dim nBoard, nPiece, nLength, nWidth, nHeight, bIsPL
    nBoard = document.form(5+a_nIndex*nFieldDiff).value
    If nBoard = "" Then
        nBoard = 1
    End If
    
    bIsPL = document.form(6+a_nIndex*nFieldDiff).value
    
    nPiece = document.form(7+a_nIndex*nFieldDiff).value
    If nPiece = "" Then
        nPiece = 1
    End If 
       
    nLength = document.form(9+a_nIndex*nFieldDiff).value
    nWidth = document.form(10+a_nIndex*nFieldDiff).value
    nHeight = document.form(11+a_nIndex*nFieldDiff).value
    
    If Trim(nLength) <> "" And Trim(nWidth) <> "" And Trim(nHeight) <> "" Then
        fVolumn = nLength * nWidth * nHeight / DivNum
        
        '不是堆量，要乘件數
        If bIsPL = "0" Then
            fVolumn = fVolumn * nPiece
        Else    '堆量, 要乘板數
            fVolumn = fVolumn * nBoard
        End If
        
        document.form(12+a_nIndex*nFieldDiff).value = FormatNumber (fVolumn, 2)                
        
        ChangeTextColor a_nIndex, nLength, nWidth, nHeight, fVolumn, bIsPL 
    End If
    
End Function


'用於查詢,修改...的體積計算, 因為欄位不一樣...
Function VolumeCalculator2(a_nIndex)
    Dim fVolumn
    Dim DivNum
    DivNum = 1000000
    fVolumn = 0
    
    
    Dim nBoard, nPiece, nLength, nWidth, nHeight, bIsPL
    nBoard = document.form(10+a_nIndex*nFieldDiff).value
    If nBoard = "" Then
        nBoard = 1
    End If
    
    bIsPL = document.form(11+a_nIndex*nFieldDiff).value
    
    nPiece = document.form(12+a_nIndex*nFieldDiff).value
    If nPiece = "" Then
        nPiece = 1
    End If 
       
    nLength = document.form(14+a_nIndex*nFieldDiff).value
    nWidth = document.form(15+a_nIndex*nFieldDiff).value
    nHeight = document.form(16+a_nIndex*nFieldDiff).value
    
    If Trim(nLength) <> "" And Trim(nWidth) <> "" And Trim(nHeight) <> "" Then
        fVolumn = nLength * nWidth * nHeight / DivNum
        
        '不是堆量，要乘件數
        If bIsPL = "0" Then
            fVolumn = fVolumn * nPiece
        Else    '堆量, 要乘板數
            fVolumn = fVolumn * nBoard
        End If
        
        document.form(17+a_nIndex*nFieldDiff).value = FormatNumber (fVolumn, 2)                
        
        ChangeTextColor2 a_nIndex, nLength, nWidth, nHeight, fVolumn, bIsPL
    End If
    
    '09-May2005: update上方的總件數與總體積
    UpdateTotalPieceAndVolume()    
    
End Function

'-----------------------update上方的總件數與總體積----------------------------------------
Function UpdateTotalPieceAndVolume()
    Dim i, nPieceSum, fVolumeSum, nPiece, bSkipAnyPiece
    nPieceSum = 0
    fVolumeSum = 0
    bSkipAnyPiece = false
    
    For i = 0 to document.form.DataCounter.value - 1
        nPiece = document.form(12+i*nFieldDiff).value
        If nPiece = "" Then
            nPiece = 0
            bSkipAnyPiece = true
        End If
        
        nPieceSum = nPieceSum + CLng(nPiece)
        fVolumeSum = fVolumeSum + document.form(17+i*nFieldDiff).value
    Next
    
    document.form.StoreSum_Piece.value = nPieceSum
    
    
    If bSkipAnyPiece Then
        document.all.td_TotalPiece.innerHTML = "<font size=6></font>"
        document.all.td_TotalVolume.innerHTML = "<font size=6></font>"
    Else
        document.all.td_TotalPiece.innerHTML = "<font size=6>" + CStr(document.form.StoreSum_Piece.value)  + "</font>"
        
        If document.form.NeededForestry.value = "" Then
            fVolumeSum = FormatNumber (fVolumeSum, 2)
            document.form.StoreSum_Volume.value = fVolumeSum
            document.all.td_TotalVolume.innerHTML = "<font size=6>" + document.form.StoreSum_Volume.value + "</font>"
        Else
            document.all.td_TotalVolume.innerHTML = "<font size=6 color=#FF0000>" + FormatNumber (document.form.NeededForestry.value, 2) + "</font>"
        End If
    
    End If
    
    
End Function


'-----------------------預測才積----------------------------------------
Function Predict_Forestry()
    
    If document.form.PredictVolume.value <> "" Then
        document.form.PredictForestry.value = document.form.PredictVolume.value * 35.3445 
        'document.form.PredictForestry.value = FormatNumber(document.form.PredictForestry.value, 2)    
        document.form.PredictForestry.value = FormatNumber(document.form.PredictForestry.value, 0)    
    End If
    
End Function

'-----------------------預測體積----------------------------------------
Function Predict_Volume()

    If document.form.PredictForestry.value <> "" Then
        document.form.PredictVolume.value = document.form.PredictForestry.value  * 0.0283 
        document.form.PredictVolume.value = FormatNumber(document.form.PredictVolume.value, 2)    
    End If
    
End Function


'-----------------------小計算機----------------------------------------
Function SimpleCalculator(a_nCurElement)
    Dim szExpression
    szExpression = Trim(document.form(a_nCurElement).value)
    
    '如果只是一個數字,就不用計算了!!
    If IsNumeric(szExpression) or document.form(a_nCurElement).value = "" Then 
        document.form(a_nCurElement).dir="rtl"
        Exit Function
    End If
    
    Dim szSubExpression, szPtr, nStrCnt, szExpArray(20), iExpIndex
    nStrCnt = 1
    iExpIndex = 0
    
    szSubExpression = szExpression
    szPtr = Mid(szSubExpression, 1, 1)
    
    While szPtr <> ""
        
        If Asc(szPtr) >= 48 and Asc(szPtr) <= 57 Then        ' 字元在0-9
            nStrCnt = nStrCnt + 1   
            szPtr = Mid(szSubExpression, nStrCnt, 1) 
        Else    
            '存數字
            szExpArray(iExpIndex) = Trim(Mid(szSubExpression, 1, nStrCnt-1)) 
            If IsNumeric(szExpArray(iExpIndex)) Then
                szExpArray(iExpIndex) = CLng(szExpArray(iExpIndex))
            End If
                   
            iExpIndex = iExpIndex + 1
            
            szSubExpression = Trim(Mid(szSubExpression, nStrCnt))
            
            '存符號
            szExpArray(iExpIndex) = Trim(Mid(szSubExpression, 1, 1))            
            iExpIndex = iExpIndex + 1            
            
            szSubExpression = Trim(Mid(szSubExpression, 2))
            
            nStrCnt = 1        
            
            szPtr = Mid(szSubExpression, nStrCnt, 1)
            
        End If
    Wend
    
    szExpArray(iExpIndex) = Trim(Mid(szSubExpression, 1, nStrCnt-1))
    If IsNumeric(szExpArray(iExpIndex)) Then
        szExpArray(iExpIndex) = CLng(szExpArray(iExpIndex))
    End If
    
    '計算
    Dim nTotal, indexTmp, FindAny, i
    FindAny = 1
    
    While iExpIndex <> 0 And FindAny = 1    
        FindAny = 0     
        For i = 0 To iExpIndex
            
            If szExpArray(i) = "*" Then
                 
                nTotal = szExpArray(i - 1) * szExpArray(i + 1)
                szExpArray(i - 1) = nTotal
                
                indexTmp = i
               
                For j = i + 2 To iExpIndex
                    szExpArray(indexTmp) = szExpArray(j)
                    indexTmp = indexTmp + 1
                Next
                
                iExpIndex = iExpIndex - 2
                
                FindAny = 1
                 
                Exit For
                
            ElseIf szExpArray(i) = "/" Then
                
                nTotal = szExpArray(i - 1) / szExpArray(i + 1)
                szExpArray(i - 1) = nTotal
                
                indexTmp = i
                
                For j = i + 2 To iExpIndex
                    szExpArray(indexTmp) = szExpArray(j)
                    indexTmp = indexTmp + 1
                Next
                
                iExpIndex = iExpIndex - 2
                
                FindAny = 1
                Exit For
                
            End If
        Next        
        
    Wend
       
    FindAny = 1
    
    While iExpIndex > 0 And FindAny = 1
        FindAny = 0
        For i = 0 To iExpIndex
            If szExpArray(i) = "+" Then
                nTotal = szExpArray(i - 1) + szExpArray(i + 1)
                szExpArray(i - 1) = nTotal
                
                indexTmp = i
                
                For j = i + 2 To iExpIndex
                    szExpArray(indexTmp) = szExpArray(j)
                    indexTmp = indexTmp + 1
                Next
                
                iExpIndex = iExpIndex - 2
                
                FindAny = 1
                Exit For
            End If
        Next
    Wend
    
    FindAny = 1
    
    While iExpIndex > 0 And FindAny = 1
        FindAny = 0
        For i = 0 To iExpIndex
            If szExpArray(i) = "-" Then
                nTotal = szExpArray(i - 1) - szExpArray(i + 1)
                szExpArray(i - 1) = nTotal
                
                indexTmp = i
                
                For j = i + 2 To iExpIndex
                    szExpArray(indexTmp) = szExpArray(j)
                    indexTmp = indexTmp - 1
                Next
                
                iExpIndex = iExpIndex - 2
                
                FindAny = 1
                Exit For
            End If
        Next
    Wend
    
    SimpleCalculator = nTotal
    
End Function

Function OnPieceFocus(a_nCurElement)
    document.form(a_nCurElement).dir="ltr"
End Function



Function CallInputBox(a_nCurElement)
    If document.form(a_nCurElement).value = "=" Then        
        document.form(a_nCurElement).value = InputBox ("請輸入運算式", "")
        
        Dim nTotal
        
        nTotal = SimpleCalculator(a_nCurElement)
        
        document.form(a_nCurElement).value = nTotal
    	document.form(a_nCurElement).dir="rtl"
    	
    	'08-Jun2005: 算完後, 直接跳到下一格
    	document.form(a_nCurElement+2).focus()
    End If
End Function

'---------------- Format 數字格式--------------
Function MyFormatNumber(a_num, a_dotnum)
	
    dim nResult, nTmpStr

	nResult = FormatNumber(a_num, a_dotnum)
	nTmpStr = CStr(nResult)
	
	nResult = Replace (nTmpStr, ",", "")
	
	MyFormatNumber = nResult
	
End Function
'----------------計算總重----------------------------------
Function WeightCalculator(a_index)
    
    dim nBoard, nPiece, nWeight
    
    nPiece = document.form(7+a_index*nFieldDiff).value
    
    If nPiece = "" Then
        nPiece = 1
    End If
        
      
    nWeight = document.form(13+a_index*nFieldDiff).value
    If nWeight <> "" Then
        document.form(14+a_index*nFieldDiff).value = MyFormatNumber(nPiece * nWeight, 1)
    End If     
    
End Function

'---------------- 計算總重: 用於查詢/修改, 因為欄位個數不同--------------
Function WeightCalculator2(a_index)
    
    dim nBoard, nPiece, nWeight
    
    nPiece = document.form(12+a_index*nFieldDiff).value
    
    If nPiece = "" Then
        nPiece = 1
    End If  
    
      
    nWeight = document.form(18+a_index*nFieldDiff).value
    If nWeight <> "" Then
        document.form(19+a_index*nFieldDiff).value = MyFormatNumber(nPiece * nWeight, 1)
    End If 
        
    
End Function


