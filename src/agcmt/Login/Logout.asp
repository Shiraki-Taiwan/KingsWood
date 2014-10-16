<%	
	Session.Contents.Remove("GroupType@AGCMT")
	Session.Contents.Remove("Owner@AGCMT")
	
	'Session.Contents("GroupType@AGCMT")	= -1 
	'Session.Contents("Owner@AGCMT")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>