<%	
	Session.Contents.Remove("GroupType@TAIINS3")
	Session.Contents.Remove("Owner@TAIINS3")
	
	'Session.Contents("GroupType@TAIINS3")	= -1 
	'Session.Contents("Owner@TAIINS3")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>