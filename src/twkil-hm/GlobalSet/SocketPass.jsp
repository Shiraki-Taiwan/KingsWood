<%@include file = "../GlobalSet/FunTP.jsp"%>
<%@page contentType="text/html;charset=Big5" language="java"
import="java.net.*"
import="java.io.*"
%>
<%
out.println("執行結果:<BR>");
String arg1,arg2,RIp;
Socket Con;
try
{
	arg1= request.getParameter("FileName");
	arg2= request.getParameter("SQLString");
	RIp= request.getRemoteAddr();
	out.print("本地端 IP:<BR><FONT COLOR=GREEN>" + RIp + "</FONT>");
	try
	{
		Con= new Socket(RIp,1024);
	}
	catch(SocketException e)
	{
		out.print("<BR><FONT COLOR=RED>連線失敗:</FONT><BR>");
		out.print(e.getMessage());
		Con= null;
		return;
	}

	try
	{
		OutputStream SocketWrite = Con.getOutputStream();
		SocketWrite.write(((String)(arg1+"[SQL]")).getBytes());
		SocketWrite.write(((String)(arg2+"[END]")).getBytes());
	}
	catch(SocketException e)
	{
		out.print("<BR><FONT COLOR=RED>資料傳回失敗:</FONT><BR><BR>");
		out.print(e.getMessage());
		Con= null;
		return;
	}

	out.print("<BR><FONT COLOR=BLUE>連線成功!!</FONT>");
	out.print("<BR>傳回參數1: <BR>");
	out.print(arg1);
	out.print("<BR>傳回參數2: <BR>");
	out.print(arg2);
}
catch(Exception e)
{
	out.println(e.getMessage());
}
%>
</table>
</form>
</center>
<%@include file = "../GlobalSet/BPanel.jsp"%>
</body>
</html>
