<%	
	Session.Contents.Remove("GroupType@KAOINS10")
	Session.Contents.Remove("Owner@KAOINS10")
	
	'Session.Contents("GroupType@KAOINS10")	= -1 
	'Session.Contents("Owner@KAOINS10")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>