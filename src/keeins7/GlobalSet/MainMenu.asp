<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@KEEINS7") < 0 then
		response.redirect "../Login/Login.asp"
	end if

	dim nGroupType	
	nGroupType = Session.Contents("GroupType@KEEINS7")
%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<link rel="stylesheet" type="text/css" href="../GlobalSet/main.css" media="screen, print" />
	<title>主選單</title>
	<script type="text/javascript" src="../Scripts/jquery-1.10.2.min.js"></script>
	<script type="text/javascript" src="../GlobalSet/ShareFun.js"></script>
	<script type="text/javascript">
		$(document).ready(function () {
			$("body").focus();
			$(document).keydown(CheckHotKey);
		});
	</script>
</head>
<body onkeydown="CheckHotKey()">
	<form name="form" method="post">
		<br />
		<table cellspacing="0" cellpadding="0" width="100%" border="0" align="center">
			<tr>
				<td>
					<table cellspacing="0" cellpadding="0" width="100%" border="0">
						<tbody>
							<tr bgcolor="#3366cc">
								<td width="1">
									<img height="26" src="../image/coin2ltb.gif" width="20"></td>
								<td align="middle" width="100%">
									<div align="center"><font color="#FFFFFF">主　選　單</font></div>
								</td>
								<td width="1">
									<img height="26" src="../image/coin2rtb.gif" width="20"></td>
							</tr>
						</tbody>
					</table>
				</td>
			</tr>
			<tr>
				<td>
					<table cellspacing="0" cellpadding="1" width="100%" bgcolor="#000000" border="0" height="8">
						<tbody>
							<tr>
								<td height="24">
									<table cellspacing="0" cellpadding="0" width="100%" bgcolor="#ebebeb" border="0" height="38">
										<tbody>
											<tr>
												<td align="left" bgcolor="#C9E0F8" height="2" width="3%"></td>
												<td align="left" bgcolor="#C9E0F8" height="2" width="41%">度量登錄作業</td>
												<td align="left" bgcolor="#C9E0F8" height="2" width="28%">基本資料設定</td>
												<td align="left" bgcolor="#C9E0F8" height="2" width="28%">列印/傳真/查詢</td>
											</tr>
											<tr>
												<td align="left" bgcolor="#C9E0F8" height="2" colspan="4">
													<hr>
												</td>
											</tr>
											<tr>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../Voyage/VOfun01a.asp">01.航次資料設定(CTRL + 7)</a></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../FreightOwner/FOfun01a.asp">01.攬貨商資料設定</a></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../Voyage/VOfun02a.asp">01.航次資料列印</a></td>
											</tr>
											<tr>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../FreightForm/FFfun02a.asp">02.進倉單資料查詢/修改/核定</a></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../Company/COfun01a.asp">02.公司資料設定</a></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../FreightOwner/FOfun02a.asp">02.攬貨商資料列印</a></td>
											</tr>
											<tr>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../FreightForm/FFfun03a.asp">03.攬貨商-單號對照資料</a></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../PackageStyle/PSfun01a.asp">03.包裝資料設定</a></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../FaxService/FSfun01a.asp">03.攬貨報告書列印傳真</a></td>
											</tr>
											<tr>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../Voyage/VOfun04a.asp">04.航次資料刪除</a></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../Voyage/VOfun03a.asp">04.航線資料設定</a></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../FaxService/FSfun02a.asp">04.傳真結果查詢</a></td>
											</tr>
											<tr>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../FaxService/FSfun04a.asp">05.e-Mail設定</a></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../FaxService/FSfun03a.asp">05.倉單資料查詢(CTRL + 9)</a></td>
											</tr>
											<%
                                	if nGroupType = 1 then
											%>
											<tr>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../Users/Usfun01a.asp">06.使用者設定</a></td>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
											</tr>
											<%
                                	end if
											%>
											<tr>
												<td align="left" bgcolor="#C9E0F8" height="2" colspan="4">
													<hr>
												</td>
											</tr>

											<%
                                	if nGroupType = 1 then
											%>
											<tr>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2">資料庫管理</td>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
											</tr>
											<tr>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../BackupDB/BKfun01a.asp">01.備份資料庫</a></td>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
											</tr>
											<tr>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../BackupDB/BKfun02a.asp">02.匯入資料庫</a></td>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
											</tr>
											<tr>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../BackupDB/publish.htm" target="_blank">03.上傳資料庫</a></td>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
											</tr>
											<tr>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"><a href="../BackupDB/DownloadDb.aspx" target="_blank">04.下載資料庫</a></td>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
												<td align="left" bgcolor="#C9E0F8" height="2"></td>
											</tr>
											<%
                                	end if
											%>
										</tbody>
									</table>
								</td>
							</tr>
						</tbody>
					</table>
				</td>
			</tr>

			<tr>
				<td>
					<table cellspacing="0" cellpadding="0" width="100%" border="0">
						<tbody>
							<tr>
								<td width="1">
									<img height="38" src="../image/box1.gif" width="20"></td>
								<td width="1">
									<img height="38" src="../image/box2.gif" width="9"></td>
								<td valign="center" align="middle" width="1" background="../image/box3.gif">&nbsp; </td>
								<td width="1">
									<img height="38" src="../image/box4.gif" width="27"></td>
								<td valign="center" align="middle" width="100%" bgcolor="#3366cc"></td>
								<td width="1">
									<img height="38" src="../image/box5.gif" width="20"></td>
							</tr>
						</tbody>
					</table>
				</td>
			</tr>
			<tr>
				<td>
					<br>
				</td>
			</tr>
			<tr>

				<td>
					<table cellspacing="0" cellpadding="0" width="100%" border="0">
						<tr>
							<td align="center" colspan="3"><font color="#0000FF" size="5">上 林 公 證 有 限 公 司</font></td>
						</tr>
						<tr>
							<td align="center" colspan="3"><font color="#0000FF" size="4">Kingswood Survey & Measurer Ltd.</font></td>
						</tr>
						<tr>
							<td align="center" colspan="3"><font color="#0000FF" size="4">版權所有 &copy 2005. All Rights Reserved. </font></td>
						</tr>
						<tr>
							<td>
								<br>
							</td>
						</tr>
						<tr>
							<td align="right"></td>
							<td align="left"><font color="#0000FF" size="2">公　司: 基隆市仁四路19巷18-2號3樓</font></td>
						</tr>
						<tr>
							<td align="right"></td>
							<td align="left"><font color="#0000FF" size="2">Office: 3FL., No.18-2, Lane 19, Jen 4 Rd Keelung, Taiwan</font></td>
						</tr>
						<tr>
							<td align="left" width="33%"></td>
							<td align="left" width="67%"><font color="#0000FF" size="2">基　隆: (02)2423-0896</font></td>
						</tr>
						<tr>
							<td align="left"></td>
							<td align="left"><font color="#0000FF" size="2">Keelung: (02)2428-9858</font></td>
						</tr>
						<tr>
							<td align="left"></td>
							<td align="left"><font color="#0000FF" size="2">F a x : (02)2423-4430</font></td>
						</tr>
						<tr>
							<td align="left"></td>
							<td align="left"><font color="#0000FF" size="2">E-mail: tslin3@yahoo.com.tw</font></td>
						</tr>
						<tr>
							<td>
								<br>
							</td>
						</tr>

					</table>
				</td>
			</tr>
		</table>
		<br />
	</form>
</body>
</html>
