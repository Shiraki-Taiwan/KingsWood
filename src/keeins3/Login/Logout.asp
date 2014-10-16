<%	
	Session.Contents.Remove("GroupType@KEEINS3")
	Session.Contents.Remove("Owner@KEEINS3")
	
	'Session.Contents("GroupType@KEEINS3")	= -1 
	'Session.Contents("Owner@KEEINS3")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>