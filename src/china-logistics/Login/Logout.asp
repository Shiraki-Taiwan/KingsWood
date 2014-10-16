<%	
	Session.Contents.Remove("GroupType@CHINA-LOGISTICS")
	Session.Contents.Remove("Owner@CHINA-LOGISTICS")
	
	'Session.Contents("GroupType@CHINA-LOGISTICS")	= -1 
	'Session.Contents("Owner@CHINA-LOGISTICS")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>