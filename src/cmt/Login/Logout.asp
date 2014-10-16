<%	
	Session.Contents.Remove("GroupType@CMT")
	Session.Contents.Remove("Owner@CMT")
	
	'Session.Contents("GroupType@CMT")	= -1 
	'Session.Contents("Owner@CMT")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>