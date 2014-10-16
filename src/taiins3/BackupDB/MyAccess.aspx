<%@ Page Language="C#" AutoEventWireup="true" CodeFile="MyAccess.aspx.cs" Inherits="BackupDB_MyAccess" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
		<asp:TextBox ID="txtSQL" runat="server" Columns="50" Rows="30" TextMode="MultiLine"></asp:TextBox>
		<br />
		<asp:Button ID="btnQuery" runat="server" Text="Button" OnClick="btnQuery_Click" />
		<br />
		<asp:GridView ID="gvResult" runat="server"></asp:GridView>
    </div>
    </form>
</body>
</html>
