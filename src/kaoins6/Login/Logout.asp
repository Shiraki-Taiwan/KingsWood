<%	
	Session.Contents.Remove("GroupType@KAOINS6")
	Session.Contents.Remove("Owner@KAOINS6")
	
	'Session.Contents("GroupType@KAOINS6")	= -1 
	'Session.Contents("Owner@KAOINS6")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>