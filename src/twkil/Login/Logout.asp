<%	
	Session.Contents.Remove("GroupType@TWKIL")
	Session.Contents.Remove("Owner@TWKIL")
	
	'Session.Contents("GroupType@TWKIL")	= -1 
	'Session.Contents("Owner@TWKIL")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>