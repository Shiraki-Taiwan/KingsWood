<!-- #include file = ../GlobalSet/conn.asp -->
<!-- #include file = ../GlobalSet/ShareFun.asp -->
<!-- #include file = FSShare.asp -->
<!--攬貨報告書列印/傳真-->
<html>
<head>
	<title>攬貨報告書列印/傳真</title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<link href="../GlobalSet/sg.css" rel="stylesheet" type="text/css" />
</head>

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@TAOINS3") < 0 then
		response.redirect "../Login/Login.asp"
	end if

	'20-Jun2005: 延長Script的timeout時間, 否則若一次送太多信, 可能會出現timeout
	Server.ScriptTimeOut = 300

    '===============接收參數===============
    dim nFoundCount, szFreightOwnerID, i, j
    dim szFaxNo, szStatus, szMailAddr, szMailAddrCC, szFreightOwnerName, szContact, szPhone, szRemark
    dim szHotKey
    
    szHotKey = request ("HotKey")		'是否使用快速鍵
    
    
    nFoundCount = request ("FoundCount")
    
    dim szSelectType, szReportType
    
    szVesselListID = request ("VesselListID")           '航次
    
    
    szSelectType = request ("SelectType")               '依攬貨商或單號...
    szReportType = request ("ReportType")               '報表格式
    
    szStatus = request ("Status")
    
    dim szVesselName, szVesselList, szVesselDate, szVesselOwner, szFormID
    szVesselName = request ("VesselName")        '船名
    szVesselList = request ("VesselList")        '航次
    szVesselDate = request ("VesselDate")        '結關日
    szVesselOwner= request ("VesselOwner")       '船公司
    szFormID     = request ("FormID")            '起始單號
     
    dim szCompanyID
    szCompanyID = request ("CompanyID")         '公司名稱
    
    dim szVesselLine, szContainer
    szVesselLine = request ("VesselLine")            '航線
    szContainer =  request ("Container")             '貨櫃場
    
    dim sql2, rs2
    
    '===========Parse 單號===================
    ParseFormID()
    
    '查航線
    dim szVesselLineName 
    set rs = nothing   
    sql = "select * from VesselLine where ID = '" + szVesselLine + "'"
    set rs = conn.execute(sql)
    if not rs.eof then
        szVesselLineName = rs("Name")
    end if
    
    '查詢公司資料
    dim szCmpChtName, szCmpEngName, szCmpContact, szCmpPhone, szCmpMobile, szCmpFaxNo
    sql = "select * from CompanyInfo where ID = '" + szCompanyID +"'"
    set rs = conn.execute(sql)
    if not rs.eof then
        szCmpChtName = rs("ChineseName")
        szCmpEngName = rs("EnglishName")
        szCmpContact = rs("Contact_1")
        szCmpPhone   = rs("Phone_1")
        szCmpMobile  = rs("Mobile_1")
        szCmpFaxNo   = rs("FaxNo_1")
    end if
    
    dim szFilePath, szFileName
    dim nRndNum
    dim szStrTmp0, szStrTmp1, szStrTmp2,  szStrTmp3,  szStrTmp4,  szStrTmp5, szStrTmp6
       
            
    dim nTotalPiece, nTotalForestry, nTotalWeight, fNeededVolume
    nTotalPiece = 0
    nTotalForestry = 0
    nTotalWeight = 0
            
    dim nColSize, nColSize2
    nColSize = 11
    nColSize2 = 8
    nColSize3 = 13
    dim szStr0, szStr1, szStr2, szStr3, szStr4, szStr5
    dim nStrLen, nSpaceCount
    dim nPrevFormID, bIsFirst
    dim nNeededVolume
    dim WeightTmp
    dim TotalWeightTmp
    
    if szSelectType = 1 then  '依單號
    	
        dim szSubStr
        szSubStr = request("UnParseedID")
        ParseFormID()
        
        dim szSubject
        szSubject = request ("Subject")
        szMailAddr = request ("MailAddress")
        szMailAddrCC = request ("MailAddressCC")
        
        if szSubject = "" then
            szSubject = "攬貨報告書"
        end if
        
        if szMailAddr = "" then
            rs.close
            conn.close
            
            set rs=nothing
            set conn=nothing
            response.write("<font size=6><br><meta http-equiv=""refresh"" content=""2; url=FSfun01a.asp""><center>沒有e-Mail位址</center>")
            response.end
        else
            
            Call InitFileNameAndPath()
                
            Set objFileSystem = CreateObject ("Scripting.FileSystemObject")
            
            '寫txt檔            
            
            szFileName = szFilePath + CStr(nRndNum) + ".txt"
            
            Set objWritedTextFile = objFileSystem.CreateTextFile(szFileName, -1, 0)
            
            
            if szReportType = 0 or szReportType = 1 then  '總表        
                Call WriteFreightFormReportHead()
                
                '查資料
                for i = 0 to nCount
 					
 					'查倉單資料
                    set rs = nothing                            
                    
                    sql = "select ID, IsChecked, NeededForestry, sum(Piece) as TotalPiece, sum(Volume) as TotalVolume, sum(TotalWeight) as TotalWeightSum"
                    sql = sql + " from FreightForm where VesselID = '" + szVesselListID + "' "
                    
                    if szStartIDTmp(i) <> "" and szEndIDTmp(i) <> "" then
                        '把單號format成4位
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
                        
                        sql2 = "select HasDelivered from FreightForm where ID = '" + CStr(rs("ID")) + "' and  VesselID = '" + szVesselListID + "'"
                        sql2 = sql2 + " and HasDelivered = false"
                        set rs2 = conn.execute(sql2)                    
                    
                        if not rs2.eof then
                            bHasDelivered = false
                        else
                            bHasDelivered = true
                        end if
                        
                        Call WriteFreightFormReportBody(rs("ID"), bHasDelivered, rs("IsChecked"), rs("NeededForestry"), rs("TotalPiece"), rs("TotalVolume"), MyFormatNumber(rs("TotalWeightSum"),1))
                                        
                        '總數
                        nTotalPiece = nTotalPiece + rs("TotalPiece")
                        
                        if rs("NeededForestry") <> 0 then
                            nTotalForestry = nTotalForestry + rs("NeededForestry")
                        else
                            nTotalForestry = nTotalForestry + rs("TotalVolume")
                        end if
                        
                        nTotalWeight = nTotalWeight + rs("TotalWeightSum")
                        
                        '註記成"已輸出過"
                        sql2 = "update FreightForm set HasDelivered = 1  where ID = '" + CStr(rs("ID")) + "' and  VesselID = '" + szVesselListID + "'"
                        conn.execute(sql2) 
                        
                        rs.movenext
                    wend 
                next
                
                Call WriteFreightFormReportTail(nTotalPiece, nTotalForestry, MyFormatNumber(nTotalWeight,1))
                
            else   '寫入標頭
                       
                nStrLen = 0
                nSpaceCount = 0
                
                nStrLen = Len(Trim(szCmpChtName))
                nSpaceCount = (56 - nStrLen * 2) / 2
                szStr1 = Space (nSpaceCount) & Trim(szCmpChtName) & Space(nSpaceCount)
                objWritedTextFile.WriteLine szStr1
                
                nStrLen = Len(Trim(szCmpEngName))
                nSpaceCount = (56 - nStrLen) / 2
                szStr1 = Space (nSpaceCount) & Trim(szCmpEngName) & Space(nSpaceCount)
                objWritedTextFile.WriteLine szStr1
                
                objWritedTextFile.WriteLine "   "
                 
                objWritedTextFile.WriteLine "                        度  量  資  料                       "
                objWritedTextFile.WriteLine "============================================================="
                
                objWritedTextFile.WriteLine "船名：" & szVesselName & "    航次：" & szVesselList &  "    結關日期：" & szVesselDate
                
                objWritedTextFile.WriteLine ""
                
                szStrTmp1 = RIGHT (Space(nColSize) & Trim("倉位"),2)
                szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("單號"),3)
                szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("板數"),3)
                szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("堆量"),3)
                szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("件數"),3)
                szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("包裝"),3)
                szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("長"),3)
                szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("寬"),3)
                szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("高"),3)
                szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("體積"),4)
                szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("單重"),5)
                szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("總重"),5)
                objWritedTextFile.WriteLine szStrTmp1
            
                objWritedTextFile.WriteLine "============================================================="
                                
                
                '寫入詳細資料
                
                bIsFirst = TRUE 
                
                nTotalPiece = 0
                fTotalVolume = 0
                
                
                
                for i = 0 to nCount                     
                    
                    '查倉單資料
                    set rs = nothing
                    
                    sql = "select FreightForm.* from FreightForm where VesselID = '" + szVesselListID + "'"
                    
                    if szStartIDTmp(i) <> "" and szEndIDTmp(i) <> "" then
                    
                        '把單號format成4位
                        szStartIDTmpFmt = szStartIDTmp(i)
                        szEndIDTmpFmt = szEndIDTmp(i)
                        
                        szStartIDTmpFmt = FormatString (szStartIDTmpFmt, 4)
                        szEndIDTmpFmt = FormatString (szEndIDTmpFmt, 4)
                        
                        sql = sql + " and ID >= '" + szStartIDTmpFmt + "' and ID <= '" + szEndIDTmpFmt + "'"
                    end if
                    
                    '20-Jul2005: 以輸入順序顯示 (The priority is higher than PageNo)
                    sql = sql + " order by ID, SN, PageNo"
                    
                    set rs = conn.execute(sql)
                                        
                    while not rs.eof
                        
                        if bIsFirst = TRUE then
                            nPrevFormID = rs("ID")
                            bIsFirst = FALSE
                        elseif nPrevFormID <> rs("ID") then
                            nPrevFormID = rs("ID")
                            objWritedTextFile.WriteLine "============================================================="
                            
                            szStrTmp1 = "    TOTAL：" & RIGHT (Space(nColSize) & Trim(CStr(nTotalPiece)),12)
                            
                            if nNeededVolume <> 0 then
                                szStrTmp1 = szStrTmp1 & "           "
                                szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(FormatNumber(nNeededVolume, 2)),12)
                            else
                                szStrTmp1 = szStrTmp1 & "           "
                                szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(FormatNumber(fTotalVolume, 2)),12)
                            end if
                            
                            objWritedTextFile.WriteLine szStrTmp1
                            objWritedTextFile.WriteLine ""
                            
                            nTotalPiece = 0
                            fTotalVolume = 0
                        end if
                                                    
                        if rs("Board") = 0 then
                            nBoardTmp = ""
                        else
                            nBoardTmp = rs("Board")
                        end if
                                
                        szStrTmp1 = RIGHT (Space(nColSize) & Trim(rs("Storehouse")),4)
                        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(rs("ID")),5)
                        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(nBoardTmp),4)
                        
                        if rs("IsPL") = TRUE then
                            szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("堆量"),3)
                        else
                            szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(""),5)
                        end if
                        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(rs("Piece")),5)
                        
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
                                
                        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(szPackageStyle),5)
                        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(rs("Length")),4)
                        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(rs("Width")),4)
                        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(rs("Height")),4)
                        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(FormatNumber(rs("Volume"))),6)
                        
                        '如果單重為0, 則不要show
                                                        
                        if rs("Weight") = 0 then
                            WeightTmp = ""
                        else
                            WeightTmp = MyFormatNumber(rs("Weight"), 1)
                        end if
                                
                        szStrTmp1 = szStrTmp1 & RIGHT (Space(18) & Trim(WeightTmp),8)
                        
                        '如果總重為0, 則不要show
                                                        
                        if rs("TotalWeight") = 0 then
                            TotalWeightTmp = ""
                        else
                            TotalWeightTmp = MyFormatNumber(rs("TotalWeight"), 1)
                        end if
                                
                        szStrTmp1 = szStrTmp1 & RIGHT (Space(16) & Trim(TotalWeightTmp),7)
                        objWritedTextFile.WriteLine szStrTmp1
                        
                        
                        nTotalPiece = nTotalPiece + rs("Piece")
                        fTotalVolume = fTotalVolume + rs("Volume")
                        
                        nNeededVolume = rs("NeededForestry")
                        
                        '註記成"已輸出過"
                        sql2 = "update FreightForm set HasDelivered = 1  where ID = '" + CStr(rs("ID")) + "' and  VesselID = '" + szVesselListID + "'"
                        conn.execute(sql2) 
                        
                        rs.movenext
                    wend
                next
                
                '最後一筆的結尾
                objWritedTextFile.WriteLine "============================================================="
                            
                szStrTmp1 = "    TOTAL：" & RIGHT (Space(nColSize) & Trim(CStr(nTotalPiece)),12)
                
                if nNeededVolume <> 0 then
                    szStrTmp1 = szStrTmp1 & "           "
                    szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(FormatNumber(nNeededVolume, 2)),12)
                else
                    szStrTmp1 = szStrTmp1 & "           "
                    szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(FormatNumber(fTotalVolume, 2)),12)
                end if
                
                objWritedTextFile.WriteLine szStrTmp1
                objWritedTextFile.WriteLine ""
                
            end if
            
            '2005-08-16: 要記得close檔案, 要不會無法刪除!!
            objWritedTextFile.Close            
            '送出郵件
            Call SendMail(szSubject, szMailAddr, szMailAddrCC, szFileName)
                
            objFileSystem.DeleteFile(szFileName)
            
        end if
        
    elseif szSelectType = 2 then  '依攬貨商     
        dim bNeedSend   
        for j = 1 to nFoundCount
          szFreightOwnerID = request ("FreightOwnerID_" + CStr(j))       '攬貨商
          szFaxNo = request ("FaxNo_" + CStr(j))                         '傳真電話
          szMailAddr = request("MailAddr_" + CStr(j))                    'E-Mail

          szMailAddrCC = request("MailAddrCC_" + CStr(j))                'E-Mail副本
          szFreightOwnerName = request ("FreightOwnerName_" + CStr(j))   '攬貨商名稱
          szContact = request ("Contact_" + CStr(j))                     '聯絡人
          szPhone = request ("Phone_" + CStr(j))                         '聯絡電話        
          szRemark = request ("Remark_" + CStr(j))                      '備註
          
          '17-Jun2005: 要記得查新的單號前得重新initialize
					nTotalPiece = 0
		    	nTotalForestry = 0
		    	nTotalWeight = 0 
          
          Call InitFileNameAndPath()
         
          Set objFileSystem = CreateObject ("Scripting.FileSystemObject")
          
          
          '寫txt檔            
          
          szFileName = szFilePath + CStr(nRndNum) + ".txt"
          
          Set objWritedTextFile = objFileSystem.CreateTextFile(szFileName, -1, 0)
          
          
          '05-Mar2005:改show傳真號碼    
          objWritedTextFile.WriteLine szFreightOwnerName + "  " + szVesselLineName + "    " + szFaxNo + "  " + szContact + "    " + "NO#" + szFreightOwnerID
          objWritedTextFile.WriteLine "   "
          
          
					if szReportType<>2 then  '2009APR 新增詳細列表     
            Call WriteFreightFormReportHead()        

                             
          	  for i = 0 to nCount        
          	      '查倉單資料
          	      set rs = nothing 
		      	                  
          	      sql = "select ID, IsChecked, NeededForestry, HasDelivered, sum(Piece) as TotalPiece, sum(Volume) as TotalVolume, sum(TotalWeight) as TotalWeightSum"
          	      sql = sql + " from FreightForm, FormToOwner where VesselID = '" + szVesselListID + "' "
          	      sql = sql + " and FreightForm.ID = FormToOwner.FormID and FormToOwner.OwnerID = '" + szFreightOwnerID + "'"
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
          	      '當用hotkey時, 只有有新增資料的攬貨商才要mail
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
          	      
          	      dim bHasDelivered
          	      
          	      while not rs.eof
          	          if rs("HasDelivered") = FALSE then
          	              bHasDelivered = false
          	          else
          	              bHasDelivered = true
          	          end if
          	      
          	          Call WriteFreightFormReportBody(rs("ID"), bHasDelivered, rs("IsChecked"), rs("NeededForestry"), rs("TotalPiece"), rs("TotalVolume"), MyFormatNumber(rs("TotalWeightSum"),1))
          	          
          	          
          	          '總數
          	          nTotalPiece = nTotalPiece + rs("TotalPiece")
          	          
          	          if rs("NeededForestry") <> 0 then
          	              nTotalForestry = nTotalForestry + rs("NeededForestry")
          	          else
          	              nTotalForestry = nTotalForestry + rs("TotalVolume")
          	          end if
          	          
          	          nTotalWeight = nTotalWeight + rs("TotalWeightSum")
          	      
          	          rs.movenext
          	      wend
          	      
          	       
          	  next
          	  
					else	'2009APR 新增詳細列表         
									
                	'寫入標頭                               	

									szMailAddr=request("MailAddress")
									szMailAddrCC=request("MailAddressCC")
								
'response.write "j="&j&" nFoundCount="&nFoundCount&"<BR>"							
                	if j >1 then    
                		bNeedSend = FALSE           'default  
                	else
          	      	bNeedSend = TRUE
          	      end if
'response.write "szMailAddr="&szMailAddr&"  "&bNeedSend&"<BR>" 
 
                	                        
                	
                	nStrLen = 0
                	nSpaceCount = 0
                	
                	nStrLen = Len(Trim(szCmpChtName))
                	nSpaceCount = (56 - nStrLen * 2) / 2
                	szStr1 = Space (nSpaceCount) & Trim(szCmpChtName) & Space(nSpaceCount)
                	objWritedTextFile.WriteLine szStr1
                	
                	nStrLen = Len(Trim(szCmpEngName))
                	nSpaceCount = (56 - nStrLen) / 2
                	szStr1 = Space (nSpaceCount) & Trim(szCmpEngName) & Space(nSpaceCount)
                	objWritedTextFile.WriteLine szStr1
                	
                	objWritedTextFile.WriteLine "   "
                	 
                	objWritedTextFile.WriteLine "                        度  量  資  料                       "
                	objWritedTextFile.WriteLine "============================================================="
                	
                	objWritedTextFile.WriteLine "船名：" & szVesselName & "    航次：" & szVesselList &  "    結關日期：" & szVesselDate
                	
                	objWritedTextFile.WriteLine ""
                	
                	szStrTmp1 = RIGHT (Space(nColSize) & Trim("倉位"),2)
                	szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("單號"),3)
                	szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("板數"),3)
                	szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("堆量"),3)
                	szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("件數"),3)
                	szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("包裝"),3)
                	szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("長"),3)
                	szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("寬"),3)
                	szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("高"),3)
                	szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("體積"),4)
                	szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("單重"),5)
                	szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("總重"),5)
                	objWritedTextFile.WriteLine szStrTmp1
                	
                	objWritedTextFile.WriteLine "============================================================="
                	                
                	
                	'寫入詳細資料                	
                	bIsFirst = TRUE                 	
                	nTotalPiece = 0
                	fTotalVolume = 0
                	
                	for i = 0 to nCount                     
                	                	    
                	    '查倉單資料
                	    set rs = nothing
                	    
                	    'sql = "select FreightForm.* from FreightForm where VesselID = '" + szVesselListID + "'"
                	    '
                	    'if szStartIDTmp(i) <> "" and szEndIDTmp(i) <> "" then
                	    '
                	    '    '把單號format成4位
                	    '    szStartIDTmpFmt = szStartIDTmp(i)
                	    '    szEndIDTmpFmt = szEndIDTmp(i)
                	    '    
                	    '    szStartIDTmpFmt = FormatString (szStartIDTmpFmt, 4)
                	    '    szEndIDTmpFmt = FormatString (szEndIDTmpFmt, 4)
                	    '    
                	    '    sql = sql + " and ID >= '" + szStartIDTmpFmt + "' and ID <= '" + szEndIDTmpFmt + "'"
                	    'end if
                	    '
                	    ''20-Jul2005: 以輸入順序顯示 (The priority is higher than PageNo)
                	    'sql = sql + " order by ID, SN, PageNo"
sql = "select FreightForm.* from FreightForm , FormToOwner where VesselID = '" + szVesselListID + "' and FreightForm.ID = FormToOwner.FormID and FormToOwner.OwnerID ='"+ FormatString(szFreightOwnerID, 3) +"' and FormToOwner.VesselLine='" + szVesselLine + "' order by ID"                                                       	
                	    
                	    set rs = conn.execute(sql)

                	    while not rs.eof
                	        
                	        if bIsFirst = TRUE then
                	            nPrevFormID = rs("ID")
                	            bIsFirst = FALSE
                	        elseif nPrevFormID <> rs("ID") then
                	            nPrevFormID = rs("ID")
                	            objWritedTextFile.WriteLine "============================================================="
                	            
                	            szStrTmp1 = "    TOTAL：" & RIGHT (Space(nColSize) & Trim(CStr(nTotalPiece)),12)
                	            
                	            if nNeededVolume <> 0 then
                	                szStrTmp1 = szStrTmp1 & "           "
                	                szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(FormatNumber(nNeededVolume, 2)),12)
                	            else
                	                szStrTmp1 = szStrTmp1 & "           "
                	                szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(FormatNumber(fTotalVolume, 2)),12)
                	            end if
                	            
                	            objWritedTextFile.WriteLine szStrTmp1
                	            objWritedTextFile.WriteLine ""
                	            
                	            nTotalPiece = 0
                	            fTotalVolume = 0
                	        end if
                	                                    
                	        if rs("Board") = 0 then
                	            nBoardTmp = ""
                	        else
                	            nBoardTmp = rs("Board")
                	        end if
                	                
                	        szStrTmp1 = RIGHT (Space(nColSize) & Trim(rs("Storehouse")),4)
                	        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(rs("ID")),5)
                	        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(nBoardTmp),4)
                	        
                	        if rs("IsPL") = TRUE then
                	            szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("堆量"),3)
                	        else
                	            szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(""),5)
                	        end if
                	        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(rs("Piece")),5)
                	        
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
                	                
                	        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(szPackageStyle),5)
                	        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(rs("Length")),4)
                	        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(rs("Width")),4)
                	        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(rs("Height")),4)
                	        szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(FormatNumber(rs("Volume"))),6)
                	        
                	        '如果單重為0, 則不要show
                	                                        
                	        if rs("Weight") = 0 then
                	            WeightTmp = ""
                	        else
                	            WeightTmp = MyFormatNumber(rs("Weight"), 1)
                	        end if
                	                
                	        szStrTmp1 = szStrTmp1 & RIGHT (Space(18) & Trim(WeightTmp),8)
                	        
                	        '如果總重為0, 則不要show
                	                                        
                	        if rs("TotalWeight") = 0 then
                	            TotalWeightTmp = ""
                	        else
                	            TotalWeightTmp = MyFormatNumber(rs("TotalWeight"), 1)
                	        end if
                	                
                	        szStrTmp1 = szStrTmp1 & RIGHT (Space(16) & Trim(TotalWeightTmp),7)
                	        objWritedTextFile.WriteLine szStrTmp1
                	        
                	        
                	        nTotalPiece = nTotalPiece + rs("Piece")
                	        fTotalVolume = fTotalVolume + rs("Volume")
                	        
                	        nNeededVolume = rs("NeededForestry")
                	        
                	        '註記成"已輸出過"
                	        sql2 = "update FreightForm set HasDelivered = 1  where ID = '" + CStr(rs("ID")) + "' and  VesselID = '" + szVesselListID + "'"
                	        conn.execute(sql2) 
                	        
                	        rs.movenext
                	    wend
                	next
                	
                	'最後一筆的結尾
                	objWritedTextFile.WriteLine "============================================================="
                	            
                	szStrTmp1 = "    TOTAL：" & RIGHT (Space(nColSize) & Trim(CStr(nTotalPiece)),12)
                	
                	if nNeededVolume <> 0 then
                	    szStrTmp1 = szStrTmp1 & "           "
                	    szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(FormatNumber(nNeededVolume, 2)),12)
                	else
                	    szStrTmp1 = szStrTmp1 & "           "
                	    szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(FormatNumber(fTotalVolume, 2)),12)
                	end if
                	
                	objWritedTextFile.WriteLine szStrTmp1
                	objWritedTextFile.WriteLine ""            
					end if
          Call WriteFreightFormReportTail(nTotalPiece, nTotalForestry, MyFormatNumber(nTotalWeight,1))
          
          
          
          '23-Apr2005: 順便列出超長超大的尺寸資料
          sql = "select * from FreightForm, FormToOwner where VesselID = '" + szVesselListID + "' "
          sql = sql + " and FreightForm.ID = FormToOwner.FormID and FormToOwner.OwnerID = '" + szFreightOwnerID + "'"
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
              
              If ((rs("Length") > 600 OR rs("Height") > 226)) then
                  '只有第一筆才show出欄位名稱
                  if fFirstData = 1 then
                      fFirstData = 0
                      
                      objWritedTextFile.WriteLine "超大超小尺寸資料:"
                      
                      szStrTmp1 = RIGHT (Space(nColSize) & Trim("頁次"),2)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("單號"),2)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("板數"),3)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("堆量"),2)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("件數"),3)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("包裝"),2)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("長"),3)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("寬"),3)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("高"),3)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("體積"),4)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("單重"),2)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("總重"),2)
                      objWritedTextFile.WriteLine szStrTmp1
                  
                      szStrTmp1 = RIGHT (Space(nColSize) & Trim("===="),4)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("===="),4)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("===="),4)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("===="),4)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("===="),4)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("===="),4)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("====="),5)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("====="),5)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("====="),5)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("===="),4)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("===="),4)
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("===="),4)
                      objWritedTextFile.WriteLine szStrTmp1

                  end if
                  
                  szStrTmp1 = RIGHT (Space(nColSize) & Trim(rs("PageNo")),3)
                  szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(rs("ID")),5)
                  szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(nBoardTmp),4)
                  
                  if rs("IsPL") = TRUE then
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim("堆量"),3)
                  else
                      szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(""),5)
                  end if
                  szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(rs("Piece")),4)
                  
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
                  	WeightTmp = MyFormatNumber(rs("Weight"),1)
                  end if
                          
                  szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(szPackageStyle),4)
                  szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(rs("Length")),4)
                  szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(rs("Width")),4)
                  szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(rs("Height")),4)
                  szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(FormatNumber(rs("Volume"))),6)
                  szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(WeightTmp),4)
                  szStrTmp1 = szStrTmp1 & RIGHT (Space(nColSize) & Trim(MyFormatNumber(rs("TotalWeight"),1)),3)
                  objWritedTextFile.WriteLine szStrTmp1
          				objWritedTextFile.WriteLine "==================================================="
              end if
              
              '25-Apr2005: 註記成"已輸出過"
              sql = "update FreightForm set HasDelivered = 1 where SN = " + CStr(rs("SN"))
              conn.execute(sql) 
              
              rs.movenext
          wend
          
          '07-Jun2005: 加入備註             
          dim bFirstLine
          bFirstLine = 0
          
          if Trim(szRemark) <> "" then
          	objWritedTextFile.WriteLine ""
          	szStrTmp1 = "備註:" & szRemark
          	
          	while Len(szStrTmp1) > 25
          		if bFirstLine = 0 then
          			objWritedTextFile.WriteLine Mid(szStrTmp1, 1, 25)
          			szStrTmp1 = Mid(szStrTmp1, 26)
          			bFirstLine = 1
          		else
          			objWritedTextFile.WriteLine "    " & Mid(szStrTmp1, 1, 23)
          			szStrTmp1 = Mid(szStrTmp1, 24)
          		end if
          	wend
          	
          	if bFirstLine = 0 then
          		objWritedTextFile.WriteLine szStrTmp1
          	else
          		objWritedTextFile.WriteLine "    " & szStrTmp1
          	end if
          
          
            '寫入資料庫
            dim rs3
            sql = "select * from FreightReportRemark where OwnerID = '" + szFreightOwnerID + "'"
            set rs3 = conn.execute(sql)
            
            if not rs3.eof then
            	sql = "update FreightReportRemark set Remark = '" + szRemark + "' where OwnerID = '" + szFreightOwnerID + "'"  
            	conn.execute(sql)
            else
            	sql = "insert into FreightReportRemark(OwnerID, Remark) values ('" + FormatString(szFreightOwnerID,3) + "','" + szRemark + "')"       	
            	conn.execute(sql)
            end if                                        
            
            rs3.close
            set rs3 = nothing 
          end if  
          
          '''''''''''''''''''''''''''''''
          
          objWritedTextFile.Close                 
          
          dim szYear, szMonth, szDay, szHour, szMin, szSec    
          szYear = Mid(CStr(Year(Date)), 3) 
          
          szMonth = Month(Date)
          if szMonth < 10 then
              szMonth = "0" & CStr(szMonth)
          end if
          
          szDay = Day(Date)
          if szDay < 10 then
              szDay = "0" & CStr(szDay)
          end if
          
          szHour = Hour(Now)
          
          if szHour < 10 then
              szHour = "0" & CStr(szHour)
          end if
          
          szMin = Minute(Now)
          
          if szMin < 10 then
              szMin = "0" & CStr(szMin)
          end if
          
          szSec = Second(Now)
          
          if szSec < 10 then
              szSec = "0" & CStr(szSec)
          end if
          
          
          '寫ini檔
            if szStatus = "Fax" then
                if bNeedSend then
                    dim szINIFileName
                    szINIFileName = szFilePath + CStr(nRndNum) + ".ini"
                    
                    Set objWritedINIFile = objFileSystem.CreateTextFile(szINIFileName, -1, 0)
                    objWritedINIFile.WriteLine "[FAXFLAG]"
                    objWritedINIFile.WriteLine "NEEDFAX = 1"
                    
                    objWritedINIFile.WriteLine "[FAXINFO]"
                    objWritedINIFile.WriteLine "FAXNO = " & szFaxNo
                    objWritedINIFile.WriteLine "FILENAME = " & szFileName 
                    objWritedINIFile.WriteLine "FONTSIZE = 180"
                
                    objWritedINIFile.WriteLine "[LOGINFO]"
                    objWritedINIFile.WriteLine "LogFileName = " & szFilePath & "FailedList\" & szYear & szMonth & szDay & "_" & szHour & szMin & szSec & ".txt"
                    objWritedINIFile.WriteLine "JobID = NO#" & szFreightOwnerID & "_" & szFreightOwnerName & "_" & szVesselName & "_" & szVesselList & "_" & szVesselDate
                
                    objWritedINIFile.Close
                else
                    objFileSystem.DeleteFile(szFileName)
                end if
            
            elseif szStatus = "Mail" then	     		
                                  '30-Mar2005: 若沒有Mail address, 則不send      
                                                                                     
                if szMailAddr <> "" and bNeedSend then                
                    Call SendMail(szFreightOwnerName + "攬貨報告書", szMailAddr, szMailAddrCC, szFileName)
                end if
                    
                objFileSystem.DeleteFile(szFileName)
            elseif szStatus = "Test" then	
				Call SendMail("【測試郵件】上林公證", "eric0629@gmail.com", "", "")
            end if

        next
        
    end if
    
    
    'Initialize檔名及路徑
    function InitFileNameAndPath()
        'random檔名
        Randomize
                    
        nRndNum=Int((50000000)*Rnd)+1
        
        szFilePath = Request.ServerVariables ("PATH_TRANSLATED")
        
        nPos = InStrRev (szFilePath, "\")
        
        szFilePath = Mid (szFilePath, 1, nPos)
        
        if szStatus = "Fax" then
            szFilePath = szFilePath + "FaxFiles\"
        elseif szStatus = "Mail" then
            szFilePath = szFilePath + "MailFiles\"
        end if
    end function
    
    '寫入攬貨報告書標頭
    function WriteFreightFormReportHead()
        dim szStr0, szStr1, szStr2, szStr3, szStr4, szStr5
        
        dim nStrLen, nSpaceCount
        nStrLen = 0
        nSpaceCount = 0
        
        nStrLen = Len(Trim(szCmpChtName))
        nSpaceCount = (52 - nStrLen * 2) / 2
        szStr1 = Space (nSpaceCount) & Trim(szCmpChtName) & Space(nSpaceCount)
        objWritedTextFile.WriteLine szStr1
        
        nStrLen = Len(Trim(szCmpEngName))
        nSpaceCount = (52 - nStrLen) / 2
        szStr1 = Space (nSpaceCount) & Trim(szCmpEngName) & Space(nSpaceCount)
        objWritedTextFile.WriteLine szStr1
        
        objWritedTextFile.WriteLine "   "
        
        objWritedTextFile.WriteLine "           LIST OF CARGO MEASURE & WEIGHT          "
        objWritedTextFile.WriteLine "==================================================="
        
        szStr1 = "船名&航次:"
        szStr1 = Left ((Trim(szStr1) & Space(7)),7)  
        szStr2 = szVesselName & "  " & szVesselList
        
        nStrLen = Len(szStr2)
        
        if nStrLen <= 6 then
            szStr3 = vbtab & vbtab & "    "
        elseif nStrLen <= 14 then
            szStr3 = vbtab & "    "
        'else
            'szStr3 = vbtab
        end if
        
        szStr3 = szStr3 & "結關日期:" + szVesselDate
         
        objWritedTextFile.WriteLine szStr1 & szStr2 & szStr3
        
        szStr1 = "船 公 司 :"
        szStr1 = Left ((Trim(szStr1) & Space(8)),8)  
        szStr2 = szVesselOwner
        
        nStrLen = Len(szStr2)
        if nStrLen <= 5 then
            szStr3 = vbtab & vbtab & "    "
        else
            szStr3 = vbtab & "    "
        end if
          
        szStr3 = szStr3 & "電話:" + szCmpPhone + " " + szCmpContact
         
        objWritedTextFile.WriteLine szStr1 & szStr2 & szStr3
        
        
        szStr1 = "貨 櫃 場 :"
        szStr1 = Left ((Trim(szStr1) & Space(8)),8)  
        szStr2 = szContainer
        
        nStrLen = Len(szStr2)
        if nStrLen <= 6 then
            szStr3 = vbtab & "    "	    
        elseif nStrLen <= 13 then
            szStr3 = "    "	    
        else
            szStr3 = vbtab
        end if
        
        szStr3 = szStr3 & "傳真:" + szCmpFaxNo + " " + szCmpContact
        
        objWritedTextFile.WriteLine szStr1 & szStr2 & szStr3
        
        szStr1 = "航線/港口:"
        szStr1 = Left ((Trim(szStr1) & Space(7)),7)  
        szStr2 = szVesselLineName
        
        nStrLen = Len(szStr2)
        if nStrLen <= 6 then
            szStr3 = vbtab & "    "
        else
            szStr3 = vbtab
        end if
        szStr3 = szStr3 & "手機:" + szCmpMobile + " " + szCmpContact
        
        objWritedTextFile.WriteLine szStr1 & szStr2 & szStr3
        
        objWritedTextFile.WriteLine "==================================================="
        
        
        szStr0 = "     "
        szStr1 = "  核對"
        szStr2 = "   " & "單號"
        szStr3 = "   " & "件數"
        szStr4 = "   " & "體積(立方公尺)"
        szStr5 = "   " & "重量"
        objWritedTextFile.WriteLine szStr0 & szStr1 & szStr2 & szStr3 & szStr4 & szStr5
        
        szStr0 = "==== "
        szStr1 = "======"
        szStr2 = " " & "======"
        szStr3 = " " & "======"
        szStr4 = "  " & "=============="
        szStr5 = "  " & "======"
        
        objWritedTextFile.WriteLine szStr0 & szStr1 & szStr2 & szStr3 & szStr4 & szStr5
    end function
    
    
    '寫入攬貨報告書內文
    function WriteFreightFormReportBody(a_ID, a_HasDelivered, a_IsChecked, a_NeededForestry, a_TotalPiece, a_TotalVolume, a_TotalWeight)        
        dim szStr0, szStr1, szStr2, szStr3, szStr4, szStr5
        
        if a_HasDelivered = TRUE then
            szStr0 = "     "
        else
            szStr0 = "新增 "
        end if
        
        if a_IsChecked = FALSE then
            szStr1 = "未核對"
        else
            szStr1 = "      "
        end if
        
        fNeededVolume = a_NeededForestry
        
        szStr2 = "   " & CStr(a_ID)
        
        szStr3 = CStr(a_TotalPiece)
        szStr3 = RIGHT (Space(nColSize) & Trim(szStr3),7)
        
        if a_NeededForestry <> 0 then
            szStr4 = CStr(FormatNumber(a_NeededForestry,2))
        else  
            szStr4 = CStr(FormatNumber(a_TotalVolume,2))
        end if
        
        szStr4 = RIGHT (Space(nColSize) & Trim(szStr4),15)
         
        szStr5 = CStr(a_TotalWeight)
        szStr5 = RIGHT (Space(nColSize) & Trim(szStr5),9)  
        objWritedTextFile.WriteLine szStr0 & szStr1 & szStr2 & szStr3 & szStr4 & szStr5
    end function
     
     
    '寫入攬貨報告書結尾
    function WriteFreightFormReportTail(a_nTotalPiece, a_nTotalForestry, a_nTotalWeight)
        dim szStr0, szStr1, szStr2, szStr3, szStr4, szStr5
                        
        szStr0 = "==== "
        szStr1 = "======"
        szStr2 = " " & "======"
        szStr3 = " " & "======"
        szStr4 = "  " & "=============="
        szStr5 = "  " & "======"
        objWritedTextFile.WriteLine szStr0 & szStr1 & szStr2 & szStr3 & szStr4 & szStr5
        
        szStr0 = "    "
        szStr1 = "    "
        szStr2 = "    " & "Total:"
        szStr3 = CStr(a_nTotalPiece)
        szStr3 = RIGHT (Space(nColSize) & Trim(szStr3),7)  
        szStr4 = CStr(FormatNumber(a_nTotalForestry,2))
        szStr4 = RIGHT (Space(nColSize) & Trim(szStr4),15)  
        szStr5 = CStr(a_nTotalWeight)
        szStr5 = RIGHT (Space(nColSize) & Trim(szStr5),9)  
           
        objWritedTextFile.WriteLine szStr0 & szStr1 & szStr2 & szStr3 & szStr4 & szStr5
        
        objWritedTextFile.WriteLine ""
        objWritedTextFile.WriteLine ""
    end function
    
    
    '寄信
    function SendMail(a_szSubject, a_szRecipient, a_szRecipientCC, szAttachment)
        '查詢Mail info
        
        sql2 = "select * from MailInfo"
        set rs2 = conn.execute(sql2)
        
        dim szSenderMail, szMailServer
        if not rs2.eof then
            szMailServer = rs2("MailServer")
            szSenderMail = rs2("SenderMailAddr")            
        end if
        
        rs2.close
        set rs2 = nothing
        
        'if szMailServer = "" then
        '    szMailServer = "agaii.com.tw"   ' default                
        'end if
        '
        'if szSenderMail = "" then
        '    szSenderMail = "Rain0510@seed.net.tw"
        'end if
	
        'Set JMail = Server.CreateOBject( "JMail.Message" )
        'JMail.logging=true
	    'JMail.silent=true
        'Jmail.From = szSenderMail
        'JMail.Subject = a_szSubject
        'JMail.AddRecipient a_szRecipient
        
        'if a_szRecipientCC <> "" then
        '	JMail.AddRecipientCC a_szRecipientCC
        'end if
        
        'JMail.Body = "如附件"
        'JMail.Charset = "big5"
        'JMail.AddAttachment(szAttachment)
        
		'if not JMail.Send( szMailServer ) then
		'	Response.write "<pre>" & JMail.log & "</pre>"
		'end if

        if szMailServer = "" then
            szMailServer = "smtp.gmail.com"   ' default                
        end if
        
        if szSenderMail = "" then
            szSenderMail		= "kingswood.notice@gmail.com"
        end if

		dim Jmail, mailsmtp
		mailsmtp				= "msa.hinet.net"

		Set Jmail				= Server.CreateOBject("JMail.Message")
		Jmail.Charset			= "utf-8"
		Jmail.ISOEncodeHeaders	= true
		'Jmail.Charset			="big5"
		Jmail.Silent			= true
		'JMail.ContentType		= "text/html"
		Jmail.From				= szSenderMail
		Jmail.FromName			= "上林公證"
		Jmail.Subject			= a_szSubject
		Jmail.AddRecipient a_szRecipient
		Jmail.Body				= "如附件"
        
        if a_szRecipientCC <> "" then
        	JMail.AddRecipientCC a_szRecipientCC
        end if

		if szAttachment <> "" then
			JMail.AddAttachment(szAttachment)
        end if
		
		if not Jmail.Send(mailsmtp) then
			Response.write "<pre>" & JMail.log & "</pre>"
		end if

		 Jmail.close
		 Set Jmail=nothing
        'Set myMail = Server.CreateOBject( "CDO.Message" )
		''============================================================
		'' 使用外部 SMTP
		''============================================================
		''設定是否使用外部 SMTP
		''1 代表使用 local smtp, 2 為外部 smtp
		'myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		''SMTP Server domain name
		'myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = szMailServer
		''Server port, gmail use ssl smtp authentication, port number is 465
		'myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
		''Authentication method, ssl or not, Username and password for the SMTP Server
		''cdoBasic 基本驗證
		'myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		'myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
		'myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "kingswood.notice"
		'myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "abcde12345!"
		'
		'myMail.Configuration.Fields.Update
		''============================================================
		'' End of 使用外部 SMTP
		''============================================================
		'myMail.Subject		= a_szSubject
		'myMail.From			= szSenderMail
		'myMail.To			= a_szRecipient
		'myMail.TextBody		= "如附件"
		'
        'if szAttachment <> "" then
		'	myMail.AddAttachment szAttachment
        'end if
		'
        'if a_szRecipientCC <> "" then
		'	myMail.Cc		= a_szRecipientCC
        'end if
        '
		'myMail.Send
		'
		'set myMail			= nothing
    end function
    
    
    rs.close
    conn.close
    
    set rs=nothing
    set conn=nothing
    
    dim szStr
    
    if szHotKey = "MailAll" or szHotKey = "MailNew" then
        if szStatus = "Fax" then
            szStr = "<font size=6><br><meta http-equiv=""refresh"" content=""2; url=../FreightForm/FFfun02b.asp?Status=FirstLoad&VesselListID=" & szVesselListID & """><center>傳真資料已送至Server!</center>"
            response.write(szStr)
        elseif szStatus = "Mail" then
            szStr = "<font size=6><br><meta http-equiv=""refresh"" content=""2; url=../FreightForm/FFfun02b.asp?Status=FirstLoad&VesselListID=" & szVesselListID & """><center>電子郵件已送出!</center>"
            response.write(szStr)
        elseif szStatus = "Test" then
			response.Write("end")
        end if
    else
        if szStatus = "Fax" then
            response.write("<font size=6><br><meta http-equiv=""refresh"" content=""2; url=FSfun01a.asp""><center>傳真資料已送至Server!</center>")
        elseif szStatus = "Mail" then
            response.write("<font size=6><br><meta http-equiv=""refresh"" content=""2; url=FSfun01a.asp""><center>電子郵件已送出!</center>")
        elseif szStatus = "Test" then
			response.Write("end")
        end if
    end if
%>  


</html>