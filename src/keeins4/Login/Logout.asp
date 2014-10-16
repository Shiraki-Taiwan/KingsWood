<%	
	Session.Contents.Remove("GroupType@KEEINS4")
	Session.Contents.Remove("Owner@KEEINS4")
	
	'Session.Contents("GroupType@KEEINS4")	= -1 
	'Session.Contents("Owner@KEEINS4")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>