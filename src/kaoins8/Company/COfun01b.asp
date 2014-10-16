<!-- #include file = ../GlobalSet/conn.asp -->
<!-- #include file = ../GlobalSet/ShareFun.asp -->
<!--公司資料設定:新增-->

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@KAOINS8") < 0 then
		response.redirect "../Login/Login.asp"
	end if


    '===============接收參數===============
    dim szStatus
    szStatus = request("Status")    '編輯Status
    
    dim szID, szChineseName, szEnglishName, szFaxNo_1, szFaxNo_2, szPhone_1, szPhone_2, szAddress
    dim szContact_1, szContact_2, szMobile_1, szMobile_2, szContainerYard
    
    if szStatus = "Modify" then        '編號
        szID = request("IDToModify")
    else
        szID = request("ID")        
    end if
          
    szChineseName   = request("ChineseName")      '公司中文名稱
    szEnglishName   = request("EnglishName")      '公司英文名稱
    szFaxNo_1       = request("FaxNo_1")          '傳真電話1
    szFaxNo_2       = request("FaxNo_2")          '傳真電話2
    szPhone_1       = request("Phone_1")          '聯絡電話1
    szPhone_2       = request("Phone_2")          '聯絡電話2
    szAddress       = request("Address")          '地址
    szContact_1     = request("Contact_1")        '聯絡人
    szContact_2     = request("Contact_2")        '聯絡人
    szMobile_1      = request("Mobile_1")         '手機
    szMobile_2      = request("Mobile_2")         '手機
    szContainerYard = request("ContainerYard") '貨櫃場
    
    '處理英文名稱中的 單引號 字元
    szEnglishName = ParseSpecialCharInAString (szEnglishName)
    
        

    '===============寫入資料庫===============  
    dim sql, rs 
    
    '=========新增=========
    if szStatus = "Add" then 
        '檢查編號是否重複  
        sql = "select ID from CompanyInfo where ID='" + szID + "'"
        set rs = conn.execute(sql)    
        
        if rs.eof then
            sql = "insert into CompanyInfo(ID, ChineseName, EnglishName, FaxNo_1, FaxNo_2, Phone_1, Phone_2, Address, Contact_1, Contact_2, Mobile_1, Mobile_2, ContainerYard)" &_
                  " values('" + szID + "','" + szChineseName + "','" + szEnglishName + "','" + szFaxNo_1 + "','" + szFaxNo_2 &_
                  "','" +  szPhone_1 + "','" + szPhone_2 + "','" + szAddress + "','" + szContact_1 + "','" + szContact_2 + "','" + szMobile_1 + "','" + szMobile_2 + "','" + szContainerYard + "')"
            conn.execute(sql)
            
            rs.close
            conn.close
            set rs=nothing
            set conn=nothing
            response.redirect "COfun01a.asp"
        else
            '編號重複, 回上一頁
            
            rs.close
            conn.close
            set rs=nothing
            set conn=nothing
            response.write("<font size=6><br><meta http-equiv=""refresh"" content=""2; url=COfun01a.asp""><center>編號重複</center>")    
        end if
        
    '=========修改=========
    elseif szStatus = "Modify" then
        sql = "update CompanyInfo set ChineseName = '"+ szChineseName + "', EnglishName = '" + szEnglishName &_
      	      "', FaxNo_1 = '" + szFaxNo_1 + "', FaxNo_2 = '" + szFaxNo_2 + "', Phone_1 = '" + szPhone_1 &_
      	      "', Phone_2 = '" + Phone_2 + "', Address = '" + szAddress +"', Contact_1 = '" + szContact_1 &_
      	      "', Contact_2 = '" + szContact_2 + "', Mobile_1 = '" + szMobile_1 + "', Mobile_2 = '" + szMobile_2 &_
      	      "', ContainerYard='" + szContainerYard + "' where ID = '"+ szID +"'"
      	
      	conn.execute(sql)
      	
      	conn.close        
      	set conn=nothing
      	
        response.redirect "COfun01a.asp"
    
    '=========刪除=========
    elseif szStatus = "Delete" then
        sql = "delete from CompanyInfo where ID ='" + szID + "'"
      	conn.execute(sql)
      	
      	conn.close
      	set conn=nothing
        response.redirect "COfun01a.asp"
        
    end if
%>