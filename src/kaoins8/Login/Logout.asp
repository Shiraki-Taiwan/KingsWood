<%	
	Session.Contents.Remove("GroupType@KAOINS8")
	Session.Contents.Remove("Owner@KAOINS8")
	
	'Session.Contents("GroupType@KAOINS8")	= -1 
	'Session.Contents("Owner@KAOINS8")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>