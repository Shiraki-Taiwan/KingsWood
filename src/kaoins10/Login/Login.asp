<%
	Session.Contents("GroupType@KAOINS10") = -1
%>
<html>
<head>
	<title>航運系統</title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<link rel="stylesheet" type="text/css" href="../GlobalSet/main.css" media="screen, print" />
	<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.10.1/jquery.min.js"></script>
	<script type="text/javascript" src="../GlobalSet/ShareFun.js"></script>
	<script type="text/javascript" src="Login.js"></script>
	<script type="text/javascript">
		$(document).ready(function () {
			//$('form').submit(function () { $.get("Login.aspx", $('form').serialize()); });
		});
		if (top.location !== self.location)
			top.location = self.location;
	</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0"
	marginheight="0" onload="OnLoad()">
	<form name="form" method="post" action="Loginb.asp" onsubmit="javascript: return checkform();">
	<br />
	<br />
	<br />
	<br />
	<br />
	<table cellspacing="0" cellpadding="0" width="100%" border="0" align="center" >
		<tbody>
			<tr>
				<td width="0%" height="111">
					&nbsp;
				</td>
				<td valign="top" height="111">
					<div style="width: 38%; margin: 0 auto;">
					<table cellspacing="0" cellpadding="0" width="100%" border="0" hspace="0"
						vspace="0">
						<tbody>
							<tr>
								<td>
									<table cellspacing="0" cellpadding="0" width="100%" border="0">
										<tbody>
											<tr bgcolor="#3366cc">
												<td width="1">
													<img alt="" height="22" src="../image/coin2ltb.gif" width="20" />
												</td>
												<td align="center" width="100%">
													<div align="center">
														<font color="#FFFFFF">量 才 程 式</font></div>
												</td>
												<td width="1">
													<img alt="" height="22" src="../image/coin2rtb.gif" width="20" />
												</td>
											</tr>
										</tbody>
									</table>
								</td>
							</tr>
							<tr>
								<td>
									<table cellspacing="0" cellpadding="1" width="100%" bgcolor="#000000" border="0"
										height="8">
										<tbody>
											<tr>
												<td height="24">
													<table cellspacing="0" cellpadding="0" width="100%" bgcolor="#ebebeb" border="0">
														<tbody>
															<tr>
																<td bgcolor="#C9E0F8" colspan="3">
																	<br>
																</td>
															</tr>
															<tr>
																<td align="right" bgcolor="#C9E0F8" height="2" width="40%">
																	帳 號：
																</td>
																<td align="left" bgcolor="#C9E0F8" height="2" width="60%">
																	<input type="text" name="ID" size="10" maxlength="20" onfocus="SelectText(0)" onkeydown="ChangeFocus(1)"
																		onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
																</td>
															</tr>
															<tr>
																<td align="right" bgcolor="#C9E0F8" height="2">
																	密 碼：
																</td>
																<td align="left" bgcolor="#C9E0F8" height="2">
																	<input type="Password" name="Password" size="10" maxlength="20" onfocus="SelectText(1)"
																		onkeydown="ChangeFocus(2)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
																</td>
															</tr>
															<tr>
																<td bgcolor="#C9E0F8" colspan="3">
																	<br>
																</td>
															</tr>
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
													<img height="38" src="../image/box1.gif" width="20">
												</td>
												<td width="1">
													<img height="38" src="../image/box2.gif" width="9">
												</td>
												<td valign="center" align="middle" width="1" background="../image/box3.gif">
													&nbsp;
												</td>
												<td width="1">
													<img height="38" src="../image/box4.gif" width="27">
												</td>
												<td valign="center" align="middle" width="100%" bgcolor="#3366cc">
													<div align="center">
														<input type="button" style="background-color: #C9E0F8; border-style: outset; border-width: 2"
															value="登錄" onkeydown="OnLoginByKeyPress()" onmouseup="OnLogin()" onfocusin="SetFocusStyle(this, true, true)"
															onfocusout="SetFocusStyle(this, false, true)">
														<input type="button" style="background-color: #C9E0F8; border-style: outset; border-width: 2"
															value="重置" onkeydown="ResetByKeyPress()" onmouseup="Reset()" onfocusin="SetFocusStyle(this, true, true)"
															onfocusout="SetFocusStyle(this, false, true)">
													</div>
												</td>
												<td width="1">
													<img height="38" src="../image/box5.gif" width="20">
												</td>
											</tr>
										</tbody>
									</table>
								</td>
							</tr>
						</tbody>
					</table>
					</div>
				</td>
				<td width="1%" height="111">
					&nbsp;
				</td>
			</tr>
		</tbody>
	</table>
	</form>
	<br />
	<br />
	<div style="width: 38%; margin: 0 auto;">
	<table cellspacing="1" cellpadding="0" width="100%" border="0" style="background-image: url('../image/Logo.gif'); background-repeat: no-repeat; background-position-x: center;">
		<tbody>
			<tr>
				<td>
					<br>
					<br>
					<br>
				</td>
			</tr>
			<tr>
				<td valign="top" align="middle" height="220">
					<div align="center">
						<font size="5">上 林 公 證 有 限 公 司<br>
							KINGSWOOD SURVEY & MEASURER LTD</font><br>
						<br>
						<font size="4">	版權所有 &copy 2013. All Rights Reserved.</font></div>
				</td>
			</tr>
		</tbody>
	</table>
	</div>
</body>
</html>