<!-- #include file = ../GlobalSet/conn.asp -->
<!DOCTYPE html>
<html xmlns:ng="http://angularjs.org">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>使用者設定</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="Usfun01a.js"></script>
	<style type="text/css">
		th, td {
			height: 2px;
		}
		th {
			text-align: right;
		}
		td {
			text-align: left;
		}
		td.data-head {
			background-color: #3366cc;
		}
		a.add, a.remove {
		}

		a.tool_btn {
			border: none;
			vertical-align: middle;
			display: inline-block;
			cursor: pointer;
			width: 20px;
			height: 20px;
			text-decoration: none;
			color: transparent;
		}

			a.tool_btn:hover {
				background-position: 0px 1px;
			}

			a.tool_btn img {
				width: 16px;
				height: 16px;
			}
	</style>
	<script type="text/javascript">
		jQuery.fn.filterByText = function (textbox) {
			return this.each(function () {
				var select = this;
				var options = [];
				$(select).find('option').each(function () { options.push({ value: $(this).val(), text: $(this).text() }); });
				$(select).data('options', options);

				$(textbox).bind('change keyup', function () {
					var options = $(select).empty().data('options');
					var search = $.trim($(this).val());
					var regex = new RegExp(search, "gi");

					$.each(options, function (i) {
						var option = options[i];
						if (option.text.match(regex) !== null) {
							$(select).append(
								$('<option>').text(option.text).val(option.value)
							);
						}
					});
				});
			});
		};
	</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<%
	'開啟ACCESS資料庫作資料庫的連結讀取
	Set ConnUser = Server.CreateObject("ADODB.Connection")
	ConnUser.open "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" & Server.MapPath("../App_Data/HAS.accdb")

    dim sql, rs
    set rs = nothing
    
    '===========接參數===========
    dim szID, szPassword, szPassword2, szStatus, szOwner
    szID			= request ("ID")
    szPassword		= request ("Password")
    szPassword2		= request ("Password2")
	szOwner			= request ("Owner")
    szStatus		= request ("Status")
    
	if request ("GroupType") <> "" then
		nGroupType	= request ("GroupType")
	end if
    
    if szStatus = "" then
        szStatus = "Add"    'Default value
    end if
    
    '===========查詢欲修改的資料===========
    dim szIDTmp, szPasswordTmp, szOwnerTmp, nGroupTypeTmp

    if szStatus = "Modify" then
        sql = "select * from Users where ID = '" + szID + "'"
        set rs = ConnUser.execute(sql)
        
        if not rs.eof then
            szIDTmp			= rs ("ID")
            szPasswordTmp	= rs ("Password")
            nGroupTypeTmp	= rs ("GroupType")
			szOwnerTmp		= rs ("Owner")
    
            '清除參數
            szID			= ""
            szPassword		= ""
            szPassword2		= ""
			szOwner			= ""
            nGroupType		= 0
        end if
    end if

	dim szOwners

	szOwners				= Split(szOwnerTmp, ",")

    '查詢出底下的list
    set rs = nothing 
    sql = "select ID,Name from FreightOwner order by ID"
    set rs = conn.execute(sql) 
	
	Dim objDictionary
	Set objDictionary = CreateObject("Scripting.Dictionary")
	objDictionary.CompareMode = vbTextCompare 'makes the keys case insensitive'
	
    dim r_szIDTmp, r_szNameTmp, r_szSelected

	while not rs.eof
		r_szIDTmp			= rs("ID")
		r_szNameTmp			= rs("Name")

		objDictionary.Add r_szIDTmp, r_szNameTmp

		rs.movenext
	wend
%>
<form name="form" method="post" action="SNfun01b.asp" onsubmit="javascript: return checkform();" OnKeyDown="CheckHotKey()" style="width: 95%; margin: 0 auto;">
	<table cellspacing="0" cellpadding="0" width="100%" border="0">
		<tr>
		    <td class="data-head" style="width: 1px;"><img height=26 src="../image/coin2ltb.gif" width=20></td>
		    <td class="data-head" style="width: 100%; text-align: center; color: #ffffff;">使用者設定</td>
		    <td class="data-head" style="width: 1px;"><img height=26 src="../image/coin2rtb.gif" width=20></td>
		</tr>
	</table>
	<table cellspacing="0" cellpadding="0" width="100%" border="0" align="center" ng-controller="MyCntrl">
		<tr>
			<td style="background-color: #c9e0f8;">
				<script type="text/javascript">
					function MyCntrl($scope) {
						$scope.options = options;
						$scope.option = $scope.options[index]
						console.log($scope.option);
					}
				</script>
				<table cellspacing="0" cellpadding="0" border="0" style="width: 100%;">
					<tr>
						<th style="width: 35%;">帳　　號：</th>
						<td style="width: 100px;">
							<input type="text" name="ID" size="20" maxlength="20" value="<%=szIDTmp%>" onfocus="SelectText(0)" onkeydown="ChangeFocus(1)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
						</td>
						<td></td>
					</tr>
					<tr>
						<th>密　　碼：</th>
						<td>
							<input type="Password" name="Password" size="20" maxlength="20" value="<%=szPasswordTmp%>" onfocus="SelectText(1)" onkeydown="ChangeFocus(2)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
						</td>
						<td></td>
					</tr>
					<tr>
						<th>密碼確認：</th>
						<td align="left" bgcolor="#C9E0F8" height="2">
							<input type="Password" name="Password2" size="20" maxlength="20" value="<%=szPasswordTmp%>" onfocus="SelectText(2)" onkeydown="ChangeFocus(3)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
						</td>
						<td></td>
					</tr>
					<tr>
						<th>群　　組：</th>
						<td align="left" bgcolor="#C9E0F8" height="2">
<%
                                    	if nGroupTypeTmp = 1 then
%>
							<input checked type="radio" name="GroupType" value="1" style="border-style: solid; border-color: #C9E0F8" />管理者
							<input type="radio" name="GroupType" value="2" style="border-style: solid; border-color: #C9E0F8" />一般使用者
							<input type="radio" name="GroupType" value="3" style="border-style: solid; border-color: #C9E0F8" />檢視人員
<%
        	                        	elseif nGroupTypeTmp = 3 then
%>
							<input type="radio" name="GroupType" value="1" style="border-style: solid; border-color: #C9E0F8" />管理者
							<input type="radio" name="GroupType" value="2" style="border-style: solid; border-color: #C9E0F8" />一般使用者
							<input checked type="radio" name="GroupType" value="3" style="border-style: solid; border-color: #C9E0F8" />檢視人員
<%
										else
%>
							<input type="radio" name="GroupType" value="1" style="border-style: solid; border-color: #C9E0F8" />管理者
							<input checked type="radio" name="GroupType" value="2" style="border-style: solid; border-color: #C9E0F8" />一般使用者
							<input type="radio" name="GroupType" value="3" style="border-style: solid; border-color: #C9E0F8" />檢視人員
<%
       	                        		end if
%>
						</td>
						<td></td>
					</tr>
					<tr>
						<th style="vertical-align: top;">所屬廠商：</th>
						<td style="text-align: left; background-color: #c9e0f8;">
<%
	For Each szOwnerTmp2 In szOwners
		if szOwnerTmp2 <> "-" And szOwnerTmp2 <> "" then
			szOwnerTmp2 = Trim(szOwnerTmp2)
%>
							<div>
								<a class="tool_btn remove" title="移除"><img alt="移除" src="../image/remove.png" /></a><label>(<%= szOwnerTmp2 %>)<%= objDictionary(CStr(szOwnerTmp2)) %></label>
								<input type="hidden" name="Owner" value="<%= szOwnerTmp2 %>" />
							</div>
<%
		end if
	Next
%>
							<div></div>
							<a class="tool_btn add" title="加入"><img alt="加入" src="../image/add.png" /></a>
							<select name="OwnerList">
								<option value="">- 無 -</option>
<%
	For Each Entry In objDictionary
		'r_szSelected			= ""
		'
		'For Each szOwnerTmp2 In szOwners
		'	if Entry = szOwnerTmp2 then
		'		r_szSelected		= "selected"
		'	end if
		'Next
		'
		'Response.Write "<option value=""" + Entry + """ " + r_szSelected + ">" + Replace(objDictionary(Entry), "'", "") + "</option>"
		Response.Write "<option value=""" + Entry + """>(" + Entry + ")" + Replace(objDictionary(Entry), "'", "") + "</option>"
	Next
%>
							</select>
							<script type="text/javascript">
								$(function () {
									$("a.add").each(function () {
										$(this).click(function () {
											console.log('click');
											var t = $(this).next("select[name=OwnerList]").find(":selected").text();
											var v = $(this).next("select[name=OwnerList]").find(":selected").val()

											console.log(t);
											console.log(v);

											var html =
												"<div>" +
												"<a class=\"tool_btn remove\" title=\"移除\"><img alt=\"移除\" src=\"../image/remove.png\" /></a>" +
												"<label>" + t + "</label>" +
												//"<label>(" + v + ")" + t + "</label>" +
												"<input type=\"hidden\" name=\"Owner\" value=\"" + v + "\" />" +
												"</div>";

											console.log(html);
											$(this).prev("div").append(html);
										});
									});
									$("a.remove").each(function () {
										$(this).click(function () {
											$(this).parent("div").remove();
										});
									});
									$('select[name=OwnerList]').filterByText($('#filter'));
								});
							</script>
						</td>
						<td style="vertical-align: bottom;">
							<input id="filter" type="text" />
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	<table cellspacing="0" cellpadding="0" width="100%" border="0">
		<tr>
			<td style="width: 1px;"><img src="../image/box1.gif" /></td>
			<td style="width: 1px;"><img src="../image/box2.gif" /></td>
			<td style="width: 1px; vertical-align: middle; text-align: center; background-image: url('../image/box3.gif'); background-repeat: repeat-x;">&nbsp;</td>
			<td style="width: 1px;"><img src="../image/box4.gif" /></td>
			<td style="width: 100%; vertical-align: middle; text-align: center; background-color: #3366cc;">
				<input type="button" style="background-color: #C9E0F8; border-style: outset; border-width: 2px" value="儲存" onkeydown="AddByKeyPress()" onmouseup="Add()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)" />
				<input type="button" style="background-color: #C9E0F8; border-style: outset; border-width: 2px" value="查詢" onkeydown="SearchByKeyPress()" onmouseup="Search()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)" />
				<input type="button" style="background-color: #C9E0F8; border-style: outset; border-width: 2px" value="刪除" onkeydown="DeleteByKeyPress()" onmouseup="Delete()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)" />
				<input type="button" style="background-color: #C9E0F8; border-style: outset; border-width: 2px" value="重置" onkeydown="ResetByKeyPress()" onmouseup="Reset()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)" />
			</td>
			<td style="width: 1px;"><img src="../image/box5.gif" /></td>
		</tr>
	</table>
	<input type="hidden" name="Status" value="<%=szStatus%>" />
	<input type="hidden" name="IDToModify" value="<%=szIDTmp%>" />
	<br />
 </form>
<%
    '查詢出底下的list
    set rs = nothing 
    sql = "select * from Users where ID like '%" + szID + "%' order by GroupType, ID"
    set rs = ConnUser.execute(sql)   
%>
<a name="D"></a>
<table cellspacing="1" cellpadding="3" style="margin: 0 auto; width: 95%; border-width: 1px; background-color: #c9e0f8;">
    <tr style="background-color: #3366cc;">
		<td class="data-head" style="width: 160px;"><span style="color: #ffffff;">帳號</span></td>
		<td class="data-head" style="width: 160px;"><span style="color: #ffffff;">群組</span></td>
		<td class="data-head" style="width: 30%;"><span style="color: #ffffff;">廠商代碼</span></td>
		<td class="data-head" style="width: auto;"><span style="color: #ffffff;">廠商名稱</span></td>
    </tr>
<%
	dim r_szOwner
	dim r_szUserType

    while not rs.eof
		r_szOwner						= rs("Owner") + ""

		szOwners						= Split(r_szOwner, ",")
	
		dim r_szFontColor

		if Len(Trim(Replace(r_szOwner, ",", ""))) > 0 then
			r_szFontColor				= "style=""color: #ff0000;"""
		else
			r_szFontColor				= ""
		end if

		if rs("GroupType") = 1 then
			r_szUserType				= "管理者"
		elseif rs("GroupType") = 2 then
			r_szUserType				= "一般使用者"
		end if
%>
    <tr style="background-color: #c9e0f8;">
		<td><a <% = r_szFontColor %> href="Usfun01a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= rs("ID") %></a></td>
		<td><a <% = r_szFontColor %> href="Usfun01a.asp?Status=Modify&ID=<%= rs("ID") %>"><%= r_szUserType %></a></td>
		<td><% = r_szOwner %></td>
		<td>
<%
	For Each szOwnerTmp2 In szOwners
		if szOwnerTmp2 <> "-" And szOwnerTmp2 <> "" then
			szOwnerTmp2 = Trim(szOwnerTmp2)
%>
			<%= objDictionary(CStr(szOwnerTmp2)) %>,
<%
		end if
	Next
%>
		</td>
    </tr>
<%
        rs.movenext
    wend
%>
</table>
<%
    rs.close
    conn.close 
	ConnUser.close 
    
    set rs=nothing
    set conn=nothing
	set ConnUser=nothing
%>
</body>
</html>
