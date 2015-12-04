<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
dim bLogSuccess, locktype, OtherArea, m_bUseEAS

OtherArea = request.querystring("OtherArea")
'm_bUseEAS = request.querystring("UesEAS")
'm_bUseEAS = LEFT(m_bUseEAS, 4)


function WritePhoneLog(sLogType, iSSN)
dim sSQL, sACSID, adorsACS, adocnLog, iLogSSN

	set adocnLog = Server.CreateObject("ADODB.Connection")
	adocnLog.Open Application("DataConn")

'    sSQL = "select ACSUserID from Userinformation where LogonID='" & Session("User") & "'"
    'sSQL = "select ACSUserTrackNumber from Userinformation where LogonID='" & Session("User") & "'"
    sSQL = "FindUserInfoUsingNetworkLogon " & Session("User")
    set adorsACS = adocnLog.execute(sSQL)
	'set adorsACS = Server.CreateObject("ADODB.Recordset")
	'adorsACS.Open sSql, adocnLog, adOpenForwardOnly, adLockOptimistic
	if adorsACS.EOF then
		WritePhoneLog = false
		adorsACS.Close
		set adorsACS = nothing
	else
'		sACSID = adorsACS("ACSUserID")
		sACSID = adorsACS("ACSUserTrackNumber")
		adorsACS.Close
		set adorsACS = nothing
		select case lcase(sLogType)
			case "override" 
				sSQL = "INSERT INTO PHONECALLS (SSN,CallSummaryCode,TimeOfCall,WhoTookCall)"
				sSQL = sSQL & " VALUES(" & iSSN & ",'OVR',GetDate(),'" & sACSID & "')"
			case else
				sSQL = "INSERT INTO PHONECALLS (SSN,TimeOfCall,WhoTookCall)"
				sSQL = sSQL & " VALUES(" & iSSN & ",GetDate(),'" & sACSID & "')"
		end select
	
		on error resume next
		adocnLog.Execute sSQL
		if adocnLog.Errors.Count > 0 then
			WritePhoneLog = false
		else
			WritePhoneLog = true
		end if
	end if
	adocnLog.Close
	set adocnLog = nothing
end function

bLogSuccess = WritePhoneLog(Request.QueryString("LogType"), Request.QueryString("SSN"))
if bLogSuccess then
	Session("CurrSSN") = Request.QueryString("SSN")
	Response.Cookies("CurrSSN") = Session("CurrSSN")
%>
	<HTML>
	<HEAD>
	</HEAD>
	<BODY onLoad="FinishLogUpdate()">
		<BR>
<!--		Writing information to the Phone Call log.  Please wait... -->
	<SCRIPT LANGUAGE=VBScript>
	
		sub FinishLogUpdate()
			top.NavFrame.iCurrSSN = <%= Request.QueryString("SSN")%>
			UpdateStatus
			if "<%= Request.QueryString("OpenPhoneSearch")%>" <> "" then
				location.replace("PhoneSearch.asp?SSN=<%= Request.QueryString("SSN")%>&AfterChange=true")
			else
				location.replace("PersonDetails.asp?SSN=<%= Request.QueryString("SSN")%>&DepNum=<%= Request.QueryString("DepNum")%>&OtherArea=<%= request.querystring("OtherArea") %>&locktype=<%= Request.QueryString("locktype")%>&Admin=<%= request.querystring("Admin") %>&UseEAS=<%= request.querystring("UseEAS") %>")
			end if
		end sub
	</SCRIPT>		
	</BODY>
<!--#include file="VBFuncs.inc" -->
	</HTML>
<%
'	if Request.QueryString("OpenPhoneSearch") <> "" then
'		Response.Redirect("PhoneSearch.asp?SSN=" & Request.QueryString("SSN") & "&AfterChange=true")
'	else
'		Response.Redirect("PersonDetails.asp?SSN=" & Request.QueryString("SSN") & "&DepNo=" & Request.QueryString("DepNo") & "&UseEAS=" & Request.QueryString("UseEAS"))
'	end if
else
%>
	<HTML>
	<HEAD>
	</HEAD>
	<BODY onLoad="UpdateStatus()">
		<BR>
		<font color="red" face="verdana, arial, helvetica" size="2"><b>Unable to successfully write to the Phone Call log.  Please contact your network administrator.</b></font>
	</BODY>
<!--#include file="VBFuncs.inc" -->
	</HTML>
<%
end if		
%>
