<%	
	Session.Contents.Remove("GroupType@DEFAULT")
	Session.Contents.Remove("Owner@DEFAULT")
	
	'Session.Contents("GroupType@DEFAULT")	= -1 
	'Session.Contents("Owner@DEFAULT")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>