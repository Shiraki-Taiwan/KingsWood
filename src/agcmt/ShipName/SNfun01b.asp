<!-- #include file = ../GlobalSet/conn.asp -->
<!--船名資料設定:新增-->

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@AGCMT") < 0 then
		response.redirect "../Login/Login.asp"
	end if


    '===============接收參數===============
    dim szStatus
    szStatus = request("Status")    '編輯Status
    
    dim szID, szName, szOwner   
       
    if szStatus = "Modify" then     '編號
        szID = request("IDToModify")
    else
        szID = request("ID")        
    end if
    
    szName   = request("Name")      '船名
    szOwner  = request("Owner")     '船公司
    

    '===============寫入資料庫===============  
    dim sql, rs 
    set rs = nothing 
  
    '=========新增=========
    if szStatus = "Add" then 
        '檢查編號是否重複  
        sql = "select ID from VesselInfo where ID='" + szID + "'"
        set rs = conn.execute(sql)    
        
        if rs.eof then
            sql = "insert into VesselInfo (ID, Name, Owner) values('" + szID + "','" + szName + "','" + szOwner + "')"
            conn.execute(sql)
            
            rs.close
            conn.close 
            set rs=nothing
            set conn=nothing
            response.redirect "SNfun01a.asp"
        else
            '編號重複, 回上一頁
            
            rs.close
            conn.close 
            set rs=nothing
            set conn=nothing
            response.write("<font size=6><br><meta http-equiv=""refresh"" content=""2; url=SNfun01a.asp""><center>船碼重複</center>")    
        end if
        
    '=========修改=========
    elseif szStatus = "Modify" then
        sql = "update VesselInfo set Name = '"+ szName +"', Owner ='" + szOwner + "' where ID = '"+ szID +"'"
      	conn.execute(sql)
      	
      	conn.close 
        set conn=nothing
        response.redirect "SNfun01a.asp"
        
    '=========刪除=========
    elseif szStatus = "Delete" then
        sql = "delete from VesselInfo where ID ='" + szID + "'"
      	conn.execute(sql)
      	
      	conn.close 
        set conn=nothing
        response.redirect "SNfun01a.asp"
    end if
%>