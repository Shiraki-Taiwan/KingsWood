<%	
	Session.Contents.Remove("GroupType@KEEINS6")
	Session.Contents.Remove("Owner@KEEINS6")
	
	'Session.Contents("GroupType@KEEINS6")	= -1 
	'Session.Contents("Owner@KEEINS6")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>