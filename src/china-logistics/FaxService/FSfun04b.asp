<!-- #include file = ../GlobalSet/conn.asp -->
<!--e-Mail設定-->
<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@CHINA-LOGISTICS") < 0 then
		response.redirect "../Login/Login.asp"
	end if


    '===============接收參數===============
    dim szMailServer, szSenderMail
    
    szMailServer   = request("MailServer") 
    szSenderMail   = request("SenderMail") 
    

    '===============寫入資料庫===============  
    dim sql
    
    '先刪除
  	sql = "delete from MailInfo"
  	conn.execute(sql)
  	
  	'再新增
  	
  	sql = "insert into MailInfo (MailServer, SenderMailAddr) values('" + szMailServer + "','" + szSenderMail + "')"
    conn.execute(sql)

    response.write("<font size=6><br><meta http-equiv=""refresh"" content=""2; url=FSfun04a.asp""><center>修改完成</center>")    
    
%>