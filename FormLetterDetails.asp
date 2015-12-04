<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
dim i, adoConn, adoRS, sSQL, sTemp, adoCmd, adoParam1
dim m_sLetterNum
dim Admin, locktype

m_sLetterNum = Request.QueryString("LetterNum")

if m_sLetterNum = "" then
	Response.Write "Letter Number missing; contact your network administrator"
else
	set adoConn = Server.CreateObject("ADODB.Connection")
	adoConn.ConnectionTimeout = 300
  adoConn.CommandTimeout = 300
  set adoCmd = Server.CreateObject("ADODB.Command")
  adoCmd.CommandTimeout = 300
  set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
  
  adoCmd.CommandText = "pb_ACELetterDetail"
  adoCmd.CommandType = adCmdStoredProc
  Set adoCmd.ActiveConnection = adoConn
  Set adoParam1 = adoCmd.CreateParameter("@P1_LetterNumber", adChar, adParamInput, 13)
  adoCmd.Parameters.Append adoParam1
  adoCmd("@P1_LetterNumber") = m_sLetterNum
%>
	<html>
	<head>
<!--#include file="VBFuncs.inc" -->
	</head>
	<body LANGUAGE="VBScript" onLoad="UpdateScreen(3)">
	<table WIDTH="100%" COLS="3">
		<tr>
			<td ALIGN="CENTER" WIDTH="10%">
<%
				if not Session("IsClerk") then
%>
					<img SRC="images/log.gif" onClick="LogCall()">
<%
				end if
%>
			</td>
			<td ALIGN="CENTER">
				<font SIZE="+2" face="verdana, arial, helvetica"><strong>Form Letter Information</strong></font>
			</td>
			<td ALIGN="CENTER" WIDTH="20%">
				<img SRC="images/bluebar2.gif" onClick="history.back">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0">
			</td>
		</tr>
	</table>
<%
	adoRS.Open adoCmd
	if adoRS.EOF then
		Response.Write ("<p><ul><font color='red' face='verdana, arial, helvetica'>Unable to find all necessary information; please contact your network administrator.</font></ul>")
	else
%>
		<table WIDTH="100%" BORDER="1" COLS=6 RULES=GROUPS bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<COLGROUP SPAN=2>
			<COLGROUP SPAN=2>
			<COLGROUP SPAN=2>
		    <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				  <td WIDTH="15%" bgcolor="#cccccc"><strong>Letter Number:</strong></td>
				  <td STYLE='color:<%= Session("EmphColor")%>;' WIDTH="25%" bgcolor="#cccccc"><%= adoRS("LetterNumber")%></td>
		      <td WIDTH="15%" bgcolor="#cccccc"><strong>Name:</strong></td>
<%
					sTemp = "<td WIDTH='20%' bgcolor='#cccccc'><A HREF='PersonDetails.asp?SSN=" & adoRS("SSN") & "&DepNum=" & adoRS("DepNumber") & "&locktype=" & adoRS("locktype") & "&Admin=" & adoRS("SupervisorAccessOnly")
					sTemp = sTemp & "' onClick='javascript:return LogCheck(" & adoRS("SSN") & ")'>"
					sTemp = sTemp & adoRS("Name") & "</A></td>"
					Response.Write sTemp
%>					    
				
		      <td WIDTH="15%" bgcolor="#cccccc"><strong>Provider:</strong></td>
				  <td WIDTH="10%" bgcolor="#cccccc"><a HREF="ProviderDetails.asp?TaxID=<%= adoRS("TaxID")%>&amp;AssocNo=<%= adoRS("Associate")%>" onClick="WorkingStatus()"><%= adoRS("Provider")%></a></td>
		    </tr>
		    <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				  <td bgcolor="#F0F0F0"><strong>Date Created:</strong></td>
				  <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("DateCreated")%></td>
		      <td bgcolor="#F0F0F0"><strong>SSN:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("SSN")%></td>
		      <td bgcolor="#F0F0F0"><strong>TaxID:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("TaxID")%></td>
		    </tr>
		    <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				  <td bgcolor="#cccccc"><strong>LogonID:</strong></td>
				  <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("LogonID")%></td>
		      <td bgcolor="#cccccc"><strong>Dependent #:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("DepNumber")%></td>
		      <td bgcolor="#cccccc"><strong>Associate #:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("Associate")%></td>
		    </tr>
		</table>
		<br>	
		<table COLS=6 WIDTH="100%" CELLPADDING="1" CELLSPACING="1" border="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td WIDTH="16%" bgcolor="#cccccc"><strong>Fund:</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' WIDTH="16%" bgcolor="#cccccc"><%= adoRS("Fund")%></td>
		    <td WIDTH="15%" bgcolor="#cccccc"><strong>Status:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' WIDTH="23%" bgcolor="#cccccc"><%= adoRS("Status")%></td>
		    <td WIDTH="15%" bgcolor="#cccccc"><strong>Craft:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' WIDTH="15%" bgcolor="#cccccc"><%= adoRS("Craft")%></td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td bgcolor="#F0F0F0"><strong>Plan Code:</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("PlanCode")%></td>
		    <td bgcolor="#F0F0F0"><strong>Status Date:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("StatusDate")%></td>
		    <td bgcolor="#F0F0F0"><strong>Local:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("Local")%></td>
			</tr>
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td bgcolor="#cccccc"><strong>HW Type:</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("HWType")%></td>
		    <td bgcolor="#cccccc"><strong>Send To:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("SendTo")%></td>
		    <td bgcolor="#cccccc"><strong>Location:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("Location")%></td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td COLSPAN=4 bgcolor="#F0F0F0"></td>
		    <td bgcolor="#F0F0F0"><strong>Employer:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("Employer")%></td>
			</tr>
		</table>
    <br>
		<hr SIZE="3" NOSHADE>
		<br>	
		<table border="1" cellpadding="2" cellspacing="2" width="100%" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td bgcolor="#cccccc"><strong>Letter Code:</strong> <%= adoRS("LetterCode")%></td>
				<td bgcolor="#cccccc"><strong>Description:</strong> <%= adoRS("Description")%></td>
				<td bgcolor="#cccccc"><strong>Second Notice Days:</strong> <%= adoRS("SecondNoticeDays")%></td>
				<td bgcolor="#cccccc"><strong>Type:</strong> <%= adoRS("Type")%></td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
			  <td COLSPAN="4" bgcolor="#F0F0F0"><strong>File Location:</strong> <a HREF="file://<%= adoRS("FileLocation")& m_sLetterNum & ".doc"%>"><%= adoRS("FileLocation")& m_sLetterNum & ".doc"%></a></td>
			</tr>
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
			  <td COLSPAN="4" bgcolor="#cccccc"><strong>Letter Name:</strong> <%=adoRS("LetterName")%></td>
			</tr>
		</table>
		<br>	
		<table ALIGN="CENTER" BORDER="1" CELLPADDING="2" CELLSPACING="2" WIDTH="80%" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td WIDTH="50%" bgcolor="#cccccc"><strong>First Full Name:</strong> <%= adoRS("FirstFullName")%></td>
				<td WIDTH="50%" bgcolor="#cccccc"><strong>Second Full Name:</strong> <%= adoRS("SecondFullName")%></td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td bgcolor="#F0F0F0"><strong>First Name:</strong> <%= adoRS("FirstName")%></td>
				<td bgcolor="#F0F0F0"><strong>First Name:</strong> <%= adoRS("SecondFirstName")%></td>
			</tr>
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td bgcolor="#cccccc"><strong>Middle Name:</strong> <%= adoRS("MiddleName")%></td>
				<td bgcolor="#cccccc"><strong>Middle Name:</strong> <%= adoRS("SecondMiddleName")%></td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td bgcolor="#F0F0F0"><strong>First Name:</strong> <%= adoRS("FirstName")%></td>
				<td bgcolor="#F0F0F0"><strong>First Name:</strong> <%= adoRS("SecondFirstName")%></td>
			</tr>
		</table>
    <p>
		<table ALIGN="CENTER" BORDER="1" CELLPADDING="2" CELLSPACING="2" WIDTH="60%" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td bgcolor="#cccccc" align="right"><strong>Address:</strong></td>
				<td bgcolor="#cccccc"><%= adoRS("Address1")%></td>
			</tr>
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td bgcolor="#CCCCCC"></td>
				<td bgcolor="#CCCCCC"><%= adoRS("Address2")%></td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td bgcolor="#F0F0F0" align="right"><strong>City:</strong></td>
				<td bgcolor="#F0F0F0"><%= adoRS("City")%></td>
			</tr>
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td bgcolor="#CCCCCC" align="right"><strong>State:</strong></td>
				<td bgcolor="#CCCCCC"><%= adoRS("State")%></td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td bgcolor="#F0F0F0" align="right"><strong>ZIP Code:</strong></td>
				<td bgcolor="#F0F0F0"><%= adoRS("PostalCode")%></td>
			</tr>
		</table>
<%
'
' CC's
'
		if adoRS("CanGenerateCCs") then
			adoRS.Close
%>
			<br>	
			<font SIZE="2" face="verdana, arial, helvetica"><center><strong>CC's</strong></center></font>
			<table WIDTH="100%" BORDER="1" CELLPADDING="2" CELLSPACING="2" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
				<tr ALIGN="CENTER" BorderColorDark="#F0F0F0" BorderColorlight="#999999">
					<td bgcolor="#CCCCCC"><strong>Full Name</strong></td>
					<td bgcolor="#CCCCCC"><strong>First</strong></td>
					<td bgcolor="#CCCCCC"><strong>Middle</strong></td>
					<td bgcolor="#CCCCCC"><strong>Last</strong></td>
					<td bgcolor="#CCCCCC"><strong>Second Full Name</strong></td>
					<td bgcolor="#CCCCCC"><strong>First</strong></td>
					<td bgcolor="#CCCCCC"><strong>Middle</strong></td>
					<td bgcolor="#CCCCCC"><strong>Last</strong></td>
					<td bgcolor="#CCCCCC"><strong>Address</strong></td>
					<td bgcolor="#CCCCCC"><strong>City</strong></td>
					<td bgcolor="#CCCCCC"><strong>State</strong></td>
					<td bgcolor="#CCCCCC"><strong>ZIP</strong></td>
				</tr>
<%
				sSQL = "select *"
				sSQL = sSQL & " FROM FormLettersCC WHERE"
				sSQL = sSQL & " LetterNumber='" & m_sLetterNum & "' order by LastName"
				adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
				if not adoRS.EOF then
					Do While Not adoRS.EOF
%>
						<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
						    <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("FirstFullName")%></td>
						    <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("FirstName")%></td>
						    <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("MiddleName")%></td>
						    <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("LastName")%></td>
						    <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("SecondFullName")%></td>
						    <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("SecondFirstName")%></td>
						    <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("SecondMiddleName")%></td>
						    <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("SecondLastName")%></td>
						    <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("Address1") & " " & adoRS("Address2")%></td>
						    <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("City")%></td>
						    <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("State")%></td>
						    <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("PostalCode")%></td>
						</tr>
<% 
						adoRS.MoveNext
					Loop
				else
%>
					<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
					    <td COLSPAN="11" bgcolor="#F0F0F0"><b><font color="red">No CC's found.</font></b></td>
					</tr>
<%
				end if  
%>
			</table>
<%
		end if
		adoRS.Close
		set adoRS = nothing
		adoConn.Close
		set adoConn = nothing
%>	
		</body>
		</html>
<%
	end if
end if
%>