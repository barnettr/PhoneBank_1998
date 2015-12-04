<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<html>
<body TOPMARGIN="2" LEFTMARGIN="2" RIGHTMARGIN="0" LANGUAGE="VBScript">
	<table WIDTH="100%" COLS="3">
		<tr>
			<td ALIGN="CENTER" WIDTH="10%">
<%
				if not Session("IsClerk") then
%>
					<img SRC="images/log.gif" onClick="LogCall()" WIDTH="75" HEIGHT="36">
<%
				end if
%>
			</td>
			<td ALIGN="CENTER">
				<font SIZE="+2" face="verdana, arial, helvetica"><strong>Other Searches</strong></font>
			</td>
			<td ALIGN="CENTER" WIDTH="10%">
				<img SRC="images/bluebar2.gif" onClick="history.go(-1)" border="0">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0">
			</td>
		</tr>
	</table>
	<br>
	<center>
	<font size="2" face="verdana, arial, helvetica">Click one of the buttons above to make the applicable search.</font>
	</center>

</body>
</html>