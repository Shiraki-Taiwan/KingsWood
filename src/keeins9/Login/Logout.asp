<%	
	Session.Contents.Remove("GroupType@KEEINS9")
	Session.Contents.Remove("Owner@KEEINS9")
	
	'Session.Contents("GroupType@KEEINS9")	= -1 
	'Session.Contents("Owner@KEEINS9")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>