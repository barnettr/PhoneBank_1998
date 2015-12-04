<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<%
	Session.Abandon
	Response.Redirect "status.asp"
%>
