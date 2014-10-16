<%	
	Session.Contents.Remove("GroupType@KEEINS5")
	Session.Contents.Remove("Owner@KEEINS5")
	
	'Session.Contents("GroupType@KEEINS5")	= -1 
	'Session.Contents("Owner@KEEINS5")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>