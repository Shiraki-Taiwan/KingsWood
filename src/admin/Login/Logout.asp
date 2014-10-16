<%	
	Session.Contents.Remove("GroupType@ADMIN")
	Session.Contents.Remove("Owner@ADMIN")
	
	'Session.Contents("GroupType@ADMIN")	= -1 
	'Session.Contents("Owner@ADMIN")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>