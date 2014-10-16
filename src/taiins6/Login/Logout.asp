<%	
	Session.Contents.Remove("GroupType@TAIINS6")
	Session.Contents.Remove("Owner@TAIINS6")
	
	'Session.Contents("GroupType@TAIINS6")	= -1 
	'Session.Contents("Owner@TAIINS6")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>