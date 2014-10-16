<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PrintOfSelect.aspx.cs" Inherits="FaxService_PrintOfSelect" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>航運系統</title>
</head>
<body style="background-image: url(../image/Logo.gif); background-position: center; background-repeat: no-repeat; background-position-y: top;">
	<div style="margin: 0px auto; margin-top: 100px; text-align: center;">
		<span style="font-size: large;">上 林 公 證 有 限 公 司<br />
			KINGSWOOD SURVEY & MEASURER LTD</span>
		<br />
		<br />
		<span style="font-size: small;">版權所有 &copy 2013. All Rights Reserved.</span>
	</div>
	<form id="form1" runat="server">
		<div style="width: 95%; margin: 0 auto;">
			<table style="background-color: #c9e0f8; width: 100%; text-align: center; border-width: 0px;">
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td colspan="4" style="font-size: 18.0pt; text-align: center;">總表查詢</td>
				</tr>
				<tr>
					<td colspan="4">
						<hr />
					</td>
				</tr>
				<tr style="font-family: Arial; font-size: 14.0pt">
					<td style="width: 1%;"></td>
					<td style="width: 20%;">船名：<%= szVesselName %></td>
					<td style="width: 20%;">航次：<%= szVesselList %></td>
					<td style="width: 69%;">結關日期：<%= szVesselDate %></td>
				</tr>
			</table>
			<table class="grid_view" style="background-color: #c9e0f8; width: 100%; border: 1px solid #000000; text-align: center;">
				<tr style="text-align: right; font-family: Arial; font-size: 14.0pt; border: 1px solid #000000;">
<%
	if (this.szGroupType.Equals("1", StringComparison.OrdinalIgnoreCase))
	{
%>
					<td>&nbsp;</td>
					<td>攬貨商編號</td>
					<td>攬貨商</td>
					<td>核對</td>
					<td>單號S/O</td>
					<td>件數</td>
					<td>體積</td>
					<td>板數</td>
					<td>總重</td>
<%
	}
	else
	{
%>
					<td>&nbsp;</td>
					<td>攬貨商</td>
					<td>核對</td>
					<td>單號S/O</td>
					<td>件數</td>
					<td>體積</td>
					<td>總重</td>
<%
	}
%>
				</tr>
<%
	if (this.szGroupType.Equals("1", StringComparison.OrdinalIgnoreCase))
	{
		foreach(DbData data in dbData)
		{
%>
				<tr style="text-align: right; font-family: Arial; font-size: 14.0pt; border: 1px solid #000000;">
					<td></td>
					<td><%= data.FO_ID %></td>
					<td><%= data.FO_Name %></td>
					<td><%= data.IsCheckedText %></td>
					<td><%= data.ID %></td>
					<td><%= data.TotalPiece %></td>
					<td><%= data.TotalVolume %></td>
					<td><%= data.TotalBoard %></td>
					<td><%= data.TotalWeightSum %></td>
				</tr>
<%
		}
	}
	else
	{
		foreach(DbData data in dbData)
		{
%>
				<tr style="text-align: right; font-family: Arial; font-size: 14.0pt; border: 1px solid #000000;">
					<td></td>
					<td><%= data.FO_Name %></td>
					<td><%= data.IsCheckedText %></td>
					<td><%= data.ID %></td>
					<td><%= data.TotalPiece %></td>
					<td><%= data.TotalVolume %></td>
					<td><%= data.TotalWeightSum %></td>
				</tr>
<%
		}
	}
	if (this.szGroupType.Equals("1", StringComparison.OrdinalIgnoreCase))
	{
%>
				<tr>
					<td colspan="9">
						<table style="width: 100%;">
							<tr>
								<td style="border-bottom: 1px solid #000000; height: 5px;"></td>
							</tr>
							<tr>
								<td style="border-top: 1px solid #000000; height: 5px;"></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr style="text-align: right; font-family: Arial; font-size: 14.0pt;">
					<td></td>
					<td>TOTAL：</td>
					<td></td>
					<td></td>
					<td></td>
					<td><%= nTotalPiece %></td>
					<td><%= nTotalVolume %></td>
					<td><%= nTotalBoard %></td>
					<td><%= nTotalWeight %></td>
				</tr>
<%
	}
	else
	{
%>
				<tr>
					<td colspan="7">
						<table style="width: 100%;">
							<tr>
								<td style="border-bottom: 1px solid #000000; height: 5px;"></td>
							</tr>
							<tr>
								<td style="border-top: 1px solid #000000; height: 5px;"></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr style="text-align: right; font-family: Arial; font-size: 14.0pt;">
					<td>TOTAL：</td>
					<td></td>
					<td></td>
					<td></td>
					<td><%= nTotalPiece %></td>
					<td><%= nTotalVolume %></td>
					<td><%= nTotalWeight %></td>
				</tr>
<%
	}
%>
			</table>
		</div>
	</form>
	<script>
		window.print();
	</script>
</body>
</html>
