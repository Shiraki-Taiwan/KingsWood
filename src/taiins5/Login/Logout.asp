<%	
	Session.Contents.Remove("GroupType@TAIINS5")
	Session.Contents.Remove("Owner@TAIINS5")
	
	'Session.Contents("GroupType@TAIINS5")	= -1 
	'Session.Contents("Owner@TAIINS5")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>