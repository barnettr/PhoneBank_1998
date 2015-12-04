<%@ LANGUAGE=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<%
if Request.ServerVariables("REMOTE_USER") = "" Then
'if session("Logon") = "" Then
'  Response.Status = "401 Unauthorized"
	Response.Redirect "failedlogin.htm"
	Response.end
Else
	Session("Logon") = Request.ServerVariables("REMOTE_USER")
	Session("User") = Right(Session("Logon"),len(Session("Logon"))-InStr(Session("Logon"), "\"))
	Response.Cookies("UserName") = Session("User")
	Session("CurrSSN") = "-1"
	Response.Cookies("CurrSSN") = Session("CurrSSN")
	Response.Redirect "pbMain.asp?NotClerk=true"
	Response.end
end if
%>
