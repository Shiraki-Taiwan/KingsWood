<%	
	Session.Contents.Remove("GroupType@TWKIL-HM")
	Session.Contents.Remove("Owner@TWKIL-HM")
	
	'Session.Contents("GroupType@TWKIL-HM")	= -1 
	'Session.Contents("Owner@TWKIL-HM")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>