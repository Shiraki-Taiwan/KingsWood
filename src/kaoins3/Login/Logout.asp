<%	
	Session.Contents.Remove("GroupType@KAOINS3")
	Session.Contents.Remove("Owner@KAOINS3")
	
	'Session.Contents("GroupType@KAOINS3")	= -1 
	'Session.Contents("Owner@KAOINS3")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>