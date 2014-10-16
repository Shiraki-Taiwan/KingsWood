<%	
	Session.Contents.Remove("GroupType@TPEINS01")
	Session.Contents.Remove("Owner@TPEINS01")
	
	'Session.Contents("GroupType@TPEINS01")	= -1 
	'Session.Contents("Owner@TPEINS01")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>