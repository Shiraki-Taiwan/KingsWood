<%	
	Session.Contents.Remove("GroupType@CMT-EMC")
	Session.Contents.Remove("Owner@CMT-EMC")
	
	'Session.Contents("GroupType@CMT-EMC")	= -1 
	'Session.Contents("Owner@CMT-EMC")		= ""

	'response.redirect "Login.asp"
	response.redirect "/"
%>