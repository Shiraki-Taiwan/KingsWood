<%	
	Session.Contents.Remove("GroupType@TAOINS4")
	Session.Contents.Remove("Owner@TAOINS4")
	
	'Session.Contents("GroupType@TAOINS4")	= -1 
	'Session.Contents("Owner@TAOINS4")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>