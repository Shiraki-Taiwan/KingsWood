<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PrintOfSelect2.aspx.cs" Inherits="FaxService_PrintOfSelect2" %>
<%@ Import Namespace="System.Collections.Generic" %>
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
					<td colspan="4" style="font-size: 18.0pt; text-align: center;">尺寸資料查詢</td>
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
					<td>頁次</td>
					<td>單號</td>
					<td>倉位</td>
					<td>板數</td>
					<td>堆量</td>
					<td>件數</td>
					<td>包裝</td>
					<td>長</td>
					<td>寬</td>
					<td>高</td>
					<td>體積</td>
					<td>單重</td>
					<td>總重</td>
				</tr>
<%
	foreach (KeyValuePair<string, DbSummaryData> summary in dbSummaryData)
	{
		foreach (DbData data in summary.Value.Items)
		{
%>
				<tr style="text-align: right; font-family: Arial; font-size: 14.0pt; border: 1px solid #000000;">
					<td><%= data.PageNo %></td>
					<td><%= data.ID %></td>
                    <td><%= data.Storehouse %></td>
                    <td><%= data.Board %></td>
					<td><%= data.IsPLText %></td>
					<td><%= data.Piece %></td>
                    <td><%= data.PackageStyle %></td>
                    <td><%= data.Length %></td>
					<td><%= data.Width %></td>
                    <td><%= data.Height %></td>
					<td><%= data.Volume %></td>
                    <td><%= data.Weight %></td>
					<td style="text-align: right;"><%= data.TotalWeight %></td>
				</tr>
<%
		}
%>
				<tr>
					<td colspan="13">
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
					<td>小計<!--頁次--></td>
					<td>&nbsp;<!--單號--></td>
					<td>&nbsp;<!--倉位--></td>
					<td>&nbsp;<!--板數--></td>
					<td>&nbsp;<!--堆量--></td>
					<td><%= summary.Value.nTotalPiece %></td>
					<td>&nbsp;<!--包裝--></td>
					<td>&nbsp;<!--長--></td>
					<td>&nbsp;<!--寬--></td>
					<td>&nbsp;<!--高--></td>
					<td><%= summary.Value.nTotalVolume %></td>
					<td>&nbsp;<!--單重--></td>
					<td>&nbsp;<!--總重--></td>
				</tr>
<%
	}
%>
				<tr>
					<td colspan="13">
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
					<td>總計<!--頁次--></td>
					<td>&nbsp;<!--單號--></td>
					<td>&nbsp;<!--倉位--></td>
					<td>&nbsp;<!--板數--></td>
					<td>&nbsp;<!--堆量--></td>
					<td><%= nTotalPiece %></td>
					<td>&nbsp;<!--包裝--></td>
					<td>&nbsp;<!--長--></td>
					<td>&nbsp;<!--寬--></td>
					<td>&nbsp;<!--高--></td>
					<td><%= nTotalVolume %></td>
					<td>&nbsp;<!--單重--></td>
					<td>&nbsp;<!--總重--></td>
				</tr>
				<%--<tr>
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
					<td><%= nTotalPiece %></td>
					<td></td>
					<td><%= nTotalWeight %></td>
					<td></td>
					<td><%= nTotalVolume %></td>
				</tr>--%>
			</table>
		</div>
	</form>
	<script>
		window.print();
	</script>
</body>
</html>
