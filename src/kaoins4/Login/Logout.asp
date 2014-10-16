<%	
	Session.Contents.Remove("GroupType@KAOINS4")
	Session.Contents.Remove("Owner@KAOINS4")
	
	'Session.Contents("GroupType@KAOINS4")	= -1 
	'Session.Contents("Owner@KAOINS4")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>