<%	
	Session.Contents.Remove("GroupType@TAIINS2")
	Session.Contents.Remove("Owner@TAIINS2")
	
	'Session.Contents("GroupType@TAIINS2")	= -1 
	'Session.Contents("Owner@TAIINS2")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>