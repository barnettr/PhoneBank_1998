<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<% Dim locktype, OtherArea, UseEAS, Admin 
locktype = request.querystring("locktype")
OtherArea = request.querystring("OtherArea")
UseEAS = request.querystring("UseEAS")
Admin = request.querystring("Admin")
if LEN(UseEAS) = 5 then
  UseEAS = UseEAS
else
  UseEAS = LEFT(UseEAS, 4)
end if 
%>

<!--#include file="user.inc" -->
<HTML>
<BODY onLoad='DoDialog()'>
</BODY>
<!--#include file="VBFuncs.inc" -->
<SCRIPT LANGUAGE="VBSCRIPT">

	function DoDialog()
	dim sTemp, sResult

		sResult=showModalDialog("LogDialog.asp", "LogDialog", "dialogWidth:500px; dialogHeight:295px; help:no;")
		select case sResult
			case "normal", "override"
				WorkingStatus
				location.replace("WriteLog.asp?SSN=<%= Request.QueryString("SSN")%>&DepNum=<%= Request.QueryString("DepNum")%>&OtherArea=<%= OtherArea %>&locktype=<%= locktype %>&Admin=<%= Admin %>&UseEAS=<%= UseEAS %>&LogType=" & sResult)
			case "cancel", ""
				window.history.back
		end select
	end function
</SCRIPT>
</HTML>