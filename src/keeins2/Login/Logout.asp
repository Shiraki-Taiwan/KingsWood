<%	
	Session.Contents.Remove("GroupType@KEEINS2")
	Session.Contents.Remove("Owner@KEEINS2")
	
	'Session.Contents("GroupType@KEEINS2")	= -1 
	'Session.Contents("Owner@KEEINS2")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>