<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim i, adoConn, adoRS, sSQL, sACSID, sTemp, sMessage
dim sHREF
dim counter

	sHREF = "PhoneSearch.asp?AfterChange=true&SSN=" & Request.QueryString("CriteriaSSN")
	sHREF = sHREF & "&CallBack=" & Request.QueryString("CallBack") & "&SummCode=" & Request.QueryString("SummCode")
	sHREF = sHREF & "&LogonID=" & Request.QueryString("LogonID") & "&ACSUserID=" & Request.QueryString("ACSID") & "&FromDate=" & Request.QueryString("FromDate") & "&FromTime="
	sHREF = sHREF & Request.QueryString("FromTime") & "&ThruDate=" & Request.QueryString("ThruDate") & "&ThruTime=" & Request.QueryString("ThruTime")
	
	set adoConn = Server.CreateObject("ADODB.Connection")
	adoConn.Open Application("DataConn")
	
	if Request.QueryString("Delete") <> "" then
		'sSQL = "DELETE FROM PhoneCalls "
		'sSQL = sSQL & " where SSN=" & Request.QueryString("SSN") 
		'sSQL = sSQL & " and TimeOfCall='" & Request.QueryString("CallTime") & "'"
    sSQL = "PB_DeletePhoneCalls " & Request.QueryString("SSN") & ", '" & Request.QueryString("CallTime") & "'" 
		on error resume next
		adoConn.Execute sSQL
		if adoConn.Errors.Count > 0 then
			sMessage = "The Phone Call for SSN " & Request.QueryString("SSN") & " and a Time"
			sMessage = sMessage & " Of Call of " & Request.QueryString("CallTime") & " could not be deleted -- contact your network administrator."
		else
			adoConn.Close
			set adoConn = nothing
			Response.Redirect sHREF
		end if
	else
'		sSQL = "select ACSUserID from Userinformation where LogonID='" & Session("User") & "'"
		sSQL = "select ACSUserTrackNumber from Userinformation where LogonID='" & Session("User") & "'"
		set adoRS = Server.CreateObject("ADODB.Recordset")
		adoRS.Open sSql, adoConn, adOpenForwardOnly, adLockOptimistic
		if adoRS.EOF then
			sMessage = "Your ACS ID could not be located (Phone Call SSN " & Request.QueryString("SSN") & " and Time"
			sMessage = sMessage & " Of Call of " & Request.QueryString("CallTime") & ") -- contact your network administrator."
		else
'			sACSID = adoRS("ACSUserID")
			sACSID = adoRS("ACSUserTrackNumber")
			adoRS.Close
			set adoRS = nothing
			sSQL = "UPDATE PhoneCalls set"
      'sSQL ="PB_UpdatePhoneCalls "
			if Request.QueryString("CancelCallBack") <> "" then
				sSQL = sSQL & " CallBack='N'"
        'sSQL ="PB_UpdatePhoneCalls @P3_CallBack='N'"
			else
				sSQL = sSQL & " NameOfCaller='" & Request.QueryString("Caller") & "'"
				sSQL = sSQL & ", PhoneNumber='" & Request.QueryString("Phone") & "'"
				sSQL = sSQL & ", Extension='" & Request.QueryString("Ext") & "'"
				sSQL = sSQL & ", NoteText='" & Request.QueryString("Notes") & "'"
				sSQL = sSQL & ", CallBack='" & Request.QueryString("NewCallBack") & "'"
				sSQL = sSQL & ", CallSummaryCode='" & Request.QueryString("NewSummCode") & "'"
				sSQL = sSQL & ", WhoUpdatedRecord='" & sACSID & "'"
				sSQL = sSQL & ", WhenRecordUpdated=GetDate()"
        'sSQL = "PB_UpdatePhoneCalls @P5_NameOfCaller='" & request.querystring("Caller") & "', @P6_PhoneNumber='" & request.querystring("Phone") & "', @P7_Extension='" & request.querystring("Ext") & "', @P8_NoteText='" & request.querystring("Notes") & "', @P3_CallBack='" & request.querystring("NewCallBack") & "', @P4_CallSummaryCode='" & request.querystring("NewSummCode") & "', @P10_WhoUpdaterecord='" & sACSID & "', @P11_WhenRecordUpdated=GetDate(), @p1_SSN='" & request.querystring("SSN") & ", @P2_TimeOfCall='" & Request.QueryString("CallTime") & "'"
			end if
			sSQL = sSQL & " where SSN=" & Request.QueryString("SSN") 
			sSQL = sSQL & " and TimeOfCall='" & Request.QueryString("CallTime") & "'" 
			on error resume next
			adoConn.Execute (sSQL)
			if adoConn.Errors.Count > 0 then
        sMessage = "Database Errors Occured" & "<P>"
        sMessage = sMessage & sSQL & "<P>"
        for counter= 0 to adoConn.Errors.Count
          sMessage = sMessage & "Error #" & adoConn.errors(counter).number & "<P>"
          sMessage = sMessage & "Error desc. -> " & adoConn.errors(counter).description & "<P>"
        next
        sMessage = sMessage & "Please copy this Error Message and foreword to your Network Administrator."
				'sMessage = "Changes could not be saved for the Phone Call with SSN " & Request.QueryString("SSN") & " and a Time"
				'sMessage = sMessage & " Of Call of " & Request.QueryString("CallTime") & " -- contact your network administrator."
			else
				adoConn.Close
				set adoConn = nothing
				Response.Redirect sHREF
			end if
		end if
		adoConn.Close
		set adoConn = nothing
	end if
%>
<HTML>
<HEAD>
</HEAD>
<BODY onLoad='UpdateStatus()'>
	<table WIDTH="100%">
		<tr>
			<td ALIGN="CENTER">
				<font SIZE="+2"><strong>Data Save for Phone Call Log</strong></font>
			</td>
		</tr>
	</table>
	<BR>
		<%= sMessage%>
	<BR>
	<BR>
	<CENTER>
	<INPUT TYPE=BUTTON NAME=Ok WIDTH=100 VALUE=Ok onClick='Acknowledge()'>
	</CENTER>
</BODY>
<!--#include file="VBFuncs.inc" -->
<SCRIPT LANGUAGE=VBSCRIPT>

	sub Acknowledge()

		top.Details.location.replace("<%= sHREF%>")
		
	end sub
	
</SCRIPT>
</HTML>
