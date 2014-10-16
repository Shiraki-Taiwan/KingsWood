<%@include file = "../GlobalSet/FunTP.jsp"%>
<%@page contentType="text/html;charset=Big5" language="java" import="java.sql.*,java.util.Enumeration"%>

<%
Connection con=null;
try
{
String FormID=request.getParameter("FormID");
String FieldID=request.getParameter("FieldID");
String SQLTable=request.getParameter("SQLTable");
String SQLField=request.getParameter("SQLField");
String val=request.getParameter("val");

con=DriverManager.getConnection((String)session.getAttribute("DB"));
Statement sta=con.createStatement();
ResultSet rs=null;

String SQLValue="",Cond="",cid;
int Num=0;
String sql,id="";

Cond=SQLField+" like '%"+val;

//chao修改 2003/7/15 start
if(SQLTable.equals("STeam"))
 sql="Select STeamID,"+SQLField+" From "+SQLTable+" Where "+Cond+"%' Group By STeamID, "+SQLField+" order by "+SQLField;
else
 sql="Select id,CName,"+SQLField+" From "+SQLTable+" Where "+Cond+"%' Group By id,CName,"+SQLField+" order by "+SQLField;

//chao修改 2003/7/15 end

rs=sta.executeQuery(sql);
%>
<Script language=javascript>
function clickurl1(val1)
{
	window.opener.<%=FormID%>.<%=FieldID%>.value=val1;
	window.close();
}
</Script>
<form name="f1" method="post" action="LSQuy02c.jsp">
	<table border=0>
		<!--Form Body Begin-->
		<tr><th colspan=8 align=center>查詢結果</th></tr>		
		<%while(rs.next()){
		SQLValue=rs.getString(SQLField);
		SQLValue.replace('\'',' ');
		SQLValue.replace('\"',' ');
		SQLValue.replace('(',' ');
		SQLValue.replace(')',' ');
		%>
		<tr><td align=center>
		<a href="#" onclick="javascript:clickurl1('<%=SQLValue%>')">
		<%
			Num++;
			out.print(SQLValue); 
		%>
		</a></td></tr>      
		<%}
		if(Num==0)
		{
		%>
		<tr><td>無相關資料</td></tr>
		<tr><td><a href="#" onclick="javascript:clickurl1('')">關閉視窗</a></td></tr>
		<%}%>
		<!--Form Body End-->
<%
rs.close(); 
con.close(); 
} 
catch(Exception e) 
{ 
out.println(e.getMessage()); 
}
%>
</body>
</html>
