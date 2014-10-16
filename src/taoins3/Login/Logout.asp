<%	
	Session.Contents.Remove("GroupType@TAOINS3")
	Session.Contents.Remove("Owner@TAOINS3")
	
	'Session.Contents("GroupType@TAOINS3")	= -1 
	'Session.Contents("Owner@TAOINS3")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>