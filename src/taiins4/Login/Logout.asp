<%	
	Session.Contents.Remove("GroupType@TAIINS4")
	Session.Contents.Remove("Owner@TAIINS4")
	
	'Session.Contents("GroupType@TAIINS4")	= -1 
	'Session.Contents("Owner@TAIINS4")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>