<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<html>
	<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<title>《航運系統》貨櫃場選擇頁</title>
		<!-- #include file = head_bundle.html -->
		<style type="text/css">
			th, td {
				padding: 3px;
			}

				td a{
					margin: 0 auto;
				}
		</style>
		<script>
function CLocation() {
}
		</script>
	</head>
	<body>
		<!-- #include file = title.htm -->
<%
    nGroupType      = Session.Contents("GroupType@KEEINS8")

	if nGroupType < 0 then
		response.redirect "Login/Login.asp"
	elseif nGroupType = 1 then
		response.redirect "Voyage/VOfun01a.asp"
	elseif nGroupType = "" then
		response.redirect "Login/Login.asp"
	end if
%>
		<div style="margin: 0px auto; width: 720px;">
			<table style="width: 100%;">
				<thead style="background-color: #3366cc; border-color: #ffffff; text-align: center;">
					<tr style="color: #ffffff;">
						<th style="width: 25%;">北部地區</th>
						<th style="width: 25%;">桃園地區</th>
						<th style="width: 25%;">台中地區</th>
						<th style="width: 25%;">南部地區</th>
					</tr>
				</thead>
				<tbody style="background-color: #c9e0f8; font-weight: bold;">
					<tr>
						<td><a href="Voyage/VOfun01a.asp">北部／中國</a></td>
						<td><a href="Voyage/VOfun01a.asp">桃園／貿聯</a></td>
						<td><a href="Voyage/VOfun01a.asp">台中／中國</a></td>
						<td><a href="Voyage/VOfun01a.asp">高雄／高鳳</a></td>
					</tr>
					<tr>
						<td><a href="Voyage/VOfun01a.asp">北部／東亞</a></td>
						<td><a href="Voyage/VOfun01a.asp">桃園／中航物流</a></td>
						<td><a href="Voyage/VOfun01a.asp">台中／長榮</a></td>
						<td><a href="Voyage/VOfun01a.asp">高雄／亞太</a></td>
					</tr>
					<tr>
						<td><a href="Voyage/VOfun01a.asp">北部／環球</a></td>
						<td><a href="Voyage/VOfun01a.asp">桃園／東海</a></td>
						<td></td>
						<td><a href="Voyage/VOfun01a.asp">高雄／友聯</a></td>
					</tr>
					<tr>
						<td><a href="Voyage/VOfun01a.asp">北部／長榮</a></td>
						<td><a href="Voyage/VOfun01a.asp">桃園／欣隆</a></td>
						<td></td>
						<td><a href="Voyage/VOfun01a.asp">高雄／高雄-66W</a></td>
					</tr>
					<tr>
						<td><a href="Voyage/VOfun01a.asp">北部／長春</a></td>
						<td></td>
						<td></td>
						<td><a href="Voyage/VOfun01a.asp">高雄／高雄-69W</a></td>
					</tr>
					<tr>
						<td><a href="Voyage/VOfun01a.asp">北部／台陽</a></td>
						<td></td>
						<td></td>
						<td><a href="Voyage/VOfun01a.asp">高雄／高雄-70W</a></td>
					</tr>
					<tr>
						<td><a href="Voyage/VOfun01a.asp">北部／弘貿</a></td>
						<td></td>
						<td></td>
						<td><a href="Voyage/VOfun01a.asp">高雄／高雄-76W</a></td>
					</tr>
					<tr>
						<td><a href="Voyage/VOfun01a.asp">北部／台基</a></td>
						<td></td>
						<td></td>
						<td><a href="Voyage/VOfun01a.asp">高雄／高雄-78W</a></td>
					</tr>
					<tr>
						<td><a href="Voyage/VOfun01a.asp">北部／中央</a></td>
						<td></td>
						<td></td>
						<td><a href="Voyage/VOfun01a.asp">高雄／高雄-118W</a></td>
					</tr>
					<tr>
						<td></td>
						<td></td>
						<td></td>
						<td><a href="Voyage/VOfun01a.asp">高雄／高雄-121W</a></td>
					</tr>
				</tbody>
				<tbody style="border-top-width: 1px; border-top-style: double; border-top-color: #fff;">
					<tr>
						<td></td>
						<td></td>
						<td></td>
						<td></td>
					</tr>
				</tbody>
			</table>
		</div>
	</body>
</html>