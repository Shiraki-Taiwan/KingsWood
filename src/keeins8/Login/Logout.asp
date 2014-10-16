<%	
	Session.Contents.Remove("GroupType@KEEINS8")
	Session.Contents.Remove("Owner@KEEINS8")
	
	'Session.Contents("GroupType@KEEINS8")	= -1 
	'Session.Contents("Owner@KEEINS8")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>