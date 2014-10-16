<!-- #include file = ../GlobalSet/conn.asp -->
<!--攬貨商資料設定:新增-->

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@AGCMT") < 0 then
		response.redirect "../Login/Login.asp"
	end if


    '===============接收參數===============
    dim szStatus
    szStatus = request("Status")    '編輯Status
    
    dim szID, szName, szFaxNo_1, szFaxNo_2, szContact_1, szContact_2, szPhone_1, szPhone_2, szAddress, szMailAddr, szMailAddrCC
      
    if szStatus = "Modify" then     '編號
        szID = request("IDToModify")
    else
        szID = request("ID")        
    end if
    
    szName          = request("Name")             '公司名稱
    szFaxNo_1       = request("FaxNo_1")          '傳真電話1
    szFaxNo_2       = request("FaxNo_2")          '傳真電話2
    szContact_1     = request("Contact_1")        '聯絡人1
    szContact_2     = request("Contact_2")        '聯絡人2
    szPhone_1       = request("Phone_1")          '聯絡電話1
    szPhone_2       = request("Phone_2")          '聯絡電話2    
    szMailAddr      = request("MailAddr")         'E-Mail
    szMailAddrCC    = request("MailAddrCC")       'E-Mail副本
    szAddress       = request("Address")          '地址
    

    '===============寫入資料庫===============  
    dim sql, rs 
    set rs = nothing 
  
    
    '=========新增=========
    if szStatus = "Add" then 
        '檢查編號是否重複  
        sql = "select ID from FreightOwner where ID='" + szID + "'"
        set rs = conn.execute(sql)    
        
        if rs.eof then
            sql = "insert into FreightOwner (ID, Name, FaxNo_1, FaxNo_2, Contact_1, Contact_2, Phone_1, Phone_2, MailAddr, MailAddrCC, Address)" &_
                  " values('" + szID + "','" + szName + "','" + szFaxNo_1 + "','" + szFaxNo_2 + "','" + szContact_1 &_
                  "','" + szContact_2 + "','" +  szPhone_1 + "','" + szPhone_2 + "','" + szMailAddr + "','" + szMailAddrCC + "','" + szAddress + "')"
            conn.execute(sql)
            
            rs.close
            conn.close 
            set rs=nothing
            set conn=nothing
            response.redirect "FOfun01a.asp"
        else
            '編號重複, 回上一頁
            rs.close
            conn.close 
            set rs=nothing
            set conn=nothing
            response.write("<font size=6><br><meta http-equiv=""refresh"" content=""2; url=FOfun01a.asp""><center>編號重複</center>")    
        end if
        
    '=========修改=========
    elseif szStatus = "Modify" then
        sql = "update FreightOwner set Name = '"+ szName + "', FaxNo_1 = '" + szFaxNo_1 + "', FaxNo_2 = '" + szFaxNo_2 &_
      	      "', Contact_1 = '" + szContact_1 + "', Contact_2 = '" + szContact_2 + "', Phone_1 = '" + szPhone_1 &_
      	      "', Phone_2 = '" + szPhone_2 + "', MailAddr = '" + szMailAddr + "', MailAddrCC ='" + szMailAddrCC + "', Address = '" + szAddress +"' where ID = '"+ szID +"'"
      	      
      	conn.execute(sql)
      	
      	conn.close 
        set conn=nothing
        response.redirect "FOfun01a.asp"
    
    '=========刪除=========
    elseif szStatus = "Delete" then
        sql = "delete from FreightOwner where ID ='" + szID + "'"
      	conn.execute(sql)
      	
      	conn.close 
        set conn=nothing
        response.redirect "FOfun01a.asp"
    end if
%>