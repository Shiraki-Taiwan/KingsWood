<!-- #include file = ../GlobalSet/conn.asp -->
<!--船名資料列印查詢-->

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@CHINA-LOGISTICS") < 0 then
		response.redirect "../Login/Login.asp"
	end if


    '===============接收參數===============
    dim szStatus
    szStatus = request("Status")    '編輯Status
    
    dim szID, szName
    
    if szStatus = "Modify" then     '編號
        szID = request("IDToModify")
    else
        szID = request("ID")        
    end if
    
    szName   = request("Name")      '名稱
    

    '===============寫入資料庫===============  
    dim sql, rs 
    set rs = nothing 
  
    '=========新增=========
    if szStatus = "Add" then 
        '檢查編號是否重複  
        sql = "select ID from PackageStyle where ID='" + szID + "'"
        set rs = conn.execute(sql)    
        
        if rs.eof then
            sql = "insert into PackageStyle (ID, Name) values('" + szID + "','" + szName + "')"
            conn.execute(sql)
            
            rs.close
            conn.close 
            set rs=nothing
            set conn=nothing
            response.redirect "PSfun01a.asp"
        else
            '編號重複, 回上一頁
            rs.close
            conn.close 
            set rs=nothing
            set conn=nothing
            response.write("<font size=6><br><meta http-equiv=""refresh"" content=""2; url=PSfun01a.asp""><center>編號重複</center>")    
        end if
        
    '=========修改=========
    elseif szStatus = "Modify" then
        sql = "update PackageStyle set Name = '"+ szName +"' where ID = '"+ szID +"'"
      	conn.execute(sql)
      	
      	conn.close 
        set conn=nothing
        response.redirect "PSfun01a.asp"
            
    '=========刪除=========
    elseif szStatus = "Delete" then
        sql = "delete from PackageStyle where ID ='" + szID + "'"
      	conn.execute(sql)
      	
      	conn.close 
        set conn=nothing
        response.redirect "PSfun01a.asp"
        
    end if
%>