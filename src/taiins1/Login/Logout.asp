<%	
	Session.Contents.Remove("GroupType@TAIINS1")
	Session.Contents.Remove("Owner@TAIINS1")
	
	'Session.Contents("GroupType@TAIINS1")	= -1 
	'Session.Contents("Owner@TAIINS1")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>