<%@include file = "../GlobalSet/Global.jsp"%>
<%@page contentType="text/html;charset=Big5" language="java" import="java.sql.*"%>
<%
String sql;
try 
{ 
	Connection con1=DriverManager.getConnection((String)session.getAttribute("DB")); 
	Connection con2=DriverManager.getConnection((String)session.getAttribute("DB")); 
	Statement sta1=con1.createStatement(); 
	Statement sta2=con2.createStatement(); 
	ResultSet rs=sta1.executeQuery("SELECT dbo.Customer.CustomerID,dbo.Goods.CStorageGoodsID,SUM(dbo.PerDayStorageGoodsHistory.QtyNumber) AS TotalNumber, COUNT(dbo.Goods.CStorageGoodsID) AS TrayNumber,dbo.StorageLocation.Kind,dbo.PerDayStorageGoodsHistory.RecordDate FROM dbo.Goods INNER JOIN dbo.StorageGoods ON dbo.Goods.id = dbo.StorageGoods.CSgid INNER JOIN dbo.PerDayStorageGoodsHistory ON dbo.StorageGoods.id = dbo.PerDayStorageGoodsHistory.CSgid INNER JOIN dbo.Customer ON dbo.Goods.Cid = dbo.Customer.id INNER JOIN dbo.StorageLocation ON dbo.StorageGoods.Slid = dbo.StorageLocation.id	GROUP BY  dbo.PerDayStorageGoodsHistory.CSgid,dbo.PerDayStorageGoodsHistory.RecordDate, dbo.Customer.CustomerID,dbo.Goods.CStorageGoodsID, dbo.StorageLocation.Kind ORDER BY dbo.PerDayStorageGoodsHistory.RecordDate,dbo.PerDayStorageGoodsHistory.CSgid");
	while(rs.next())
	{
		sql="Insert into SGHistoryAccount(CustomerID,CStorageGoodsID,TotalNumber,TrayNumber,Kind,RecordDate) Values('"+rs.getString("CustomerID")+"','"+rs.getString("CStorageGoodsID")+"',"+rs.getString("TotalNumber")+","+rs.getString("TrayNumber")+","+rs.getString("Kind")+",'"+rs.getString("RecordDate")+"')";
		//out.println(sql);
		sta2.executeUpdate(sql);
	}
	rs.close();
	con1.close();
	con2.close();
} 
catch(Exception e) 
{ 
out.println(e.getMessage()); 
}
%>