<%	
	Session.Contents.Remove("GroupType@KAOINS9")
	Session.Contents.Remove("Owner@KAOINS9")
	
	'Session.Contents("GroupType@KAOINS9")	= -1 
	'Session.Contents("Owner@KAOINS9")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>