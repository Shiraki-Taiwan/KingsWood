<%@ Page Language="C#" AutoEventWireup="true" CodeFile="UploadDb.aspx.cs" Inherits="BackupDB_UploadDb" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>上傳資料庫</title>
	<!-- #include file = ../head_bundle.html -->
	<style>
		body {
			margin: 0;
		}

		input.upload {
			font-size: 12px;
			width: 200px;
		}
	</style>
</head>
<body>
	<ul style="width: 170px;" class="mainmenu">
		<li><a>度量登錄作業</a>
			<ul>
				<li><a href="../Voyage/VOfun01a.asp">航次資料設定(CTRL + 7)</a></li>
				<li><a href="../FreightForm/FFfun02a.asp">進倉單資料查詢/修改/核定</a></li>
			</ul>
		</li>
	</ul>
	<ul style="width: 170px;" class="mainmenu">
		<li><a href="../GlobalSet/MainMenu.asp">主選單(Ctrl + Q)</a></li>
	</ul>
	<ul style="width: 160px;" class="mainmenu">
		<li><a href="../Login/Logout.asp">登　出</a></li>
	</ul>
	<ul style="border: solid 0px #3366CC; width: 800px;" class="mainmenu">
		<li><span style="font-size: medium; color: #0000ff">上林公證有限公司 版權所有 &copy; 2013. All Rights Reserved.<br />
			Kingswood Survey & Measurer Ltd.</span></li>
	</ul>
	<form id="form1" runat="server">
		<table cellspacing="0" cellpadding="0" style="border-width: 0px; width: 80%; margin: 0 auto;">
			<tr style="background-color: #3366cc;">
				<td style="width: 1px;">
					<img alt="" src="../image/coin2ltb.gif" /></td>
				<td style="text-align: center; width: 100%;">
					<span style="color: #fff;">匯入資料庫</span>
				</td>
				<td style="width: 1px;">
					<img alt="" src="../image/coin2rtb.gif" /></td>
			</tr>
		</table>
		<table cellspacing="0" cellpadding="0" style="border-width: 0px; width: 80%; margin: 0 auto; background-color: #ebebeb;">
			<tr>
				<td style="width: 120px;"></td>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td style="width: 120px; font-size: 0.8em; text-align: right;">選擇檔案：</td>
				<td style="font-size: 0.8em;">
					<asp:FileUpload ID="uploadDb" runat="server" CssClass="upload" />
				</td>
			</tr>
			<tr>
				<td style="width: 120px;"></td>
				<td>&nbsp;</td>
			</tr>
		</table>
		<table cellspacing="0" cellpadding="0" style="border-width: 0px; width: 80%; margin: 0 auto;">
			<tr>
				<td style="width: 1px;">
					<img alt="" src="../image/box1.gif" /></td>
				<td style="width: 1px;">
					<img alt="" src="../image/box2.gif" /></td>
				<td style="background-image: url(../image/box3.gif); background-repeat: repeat-x; width: 1px; text-align: center; vertical-align: middle;">&nbsp; </td>
				<td style="width: 1px;">
					<img alt="" src="../image/box4.gif" /></td>
				<td style="width: 100%; padding: 0px; vertical-align: top;">
					<div style="background-color: #3366cc; width: 100%; height: 38px; margin: 0px; text-align: center; vertical-align: middle;">
						<asp:Button ID="btnUpload" runat="server" Text="上傳" OnClick="btnUpload_Click" />
					</div>
				</td>
				<td style="width: 1px;">
					<img alt="" src="../image/box5.gif" /></td>
			</tr>
		</table>
	</form>
</body>
</html>
