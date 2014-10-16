<%@include file = "../GlobalSet/FunTP.jsp"%>
<%@page contentType="text/html;charset=Big5" language="java" import="java.sql.*"%>
<%
	String FormID=request.getParameter("FormID");
	String FieldID=request.getParameter("FieldID");
	String SQLTable=request.getParameter("SQLTable");
	String SQLField=request.getParameter("SQLField");
	String SQLField2=request.getParameter("SQLField2");
	String SQLCond2=request.getParameter("SQLCond2");
%>
<form name=f1 method="post" action="Fuzzyb.jsp">
<input type=hidden name="FormID" value=<%=FormID%>>
<input type=hidden name="FieldID" value=<%=FieldID%>>
<input type=hidden name="SQLTable" value=<%=SQLTable%>>
<input type=hidden name="SQLField" value=<%=SQLField%>>
<input type=hidden name="val" value="">
</form>
<Script Language=javascript>
f1.val.value=window.opener.<%=FormID%>.<%=FieldID%>.value;
f1.submit();
</Script>
