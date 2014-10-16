<%	
	Session.Contents.Remove("GroupType@KAOINS2")
	Session.Contents.Remove("Owner@KAOINS2")
	
	'Session.Contents("GroupType@KAOINS2")	= -1 
	'Session.Contents("Owner@KAOINS2")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>