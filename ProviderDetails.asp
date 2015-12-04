<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim adoConn, adoRS, sSQL, sTemp
Dim m_iTaxID, m_iAssoc, m_sCredentials

m_iTaxID = Request.QueryString("TaxID")
m_iAssoc = Request.QueryString("AssocNo")

if m_iTaxID = "" or m_iAssoc = "" then
	Response.Write "Tax ID or Associate Number missing; contact your network administrator"
else
	set adoConn = Server.CreateObject("ADODB.Connection")
'***********************************************************************************
' Setting CommandTimeout to 0 waits indefinitely for response; note that per 
' ADO documentation Command object will not inherit the Connection setting (default
' is 30 seconds).  This appeared to help with response issues on seasql02.
'***********************************************************************************
	adoConn.CommandTimeout = 0
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
%>
	<html>
	<head>
<!--#include file="VBFuncs.inc" -->
	</head>
	<body LANGUAGE="VBScript" onLoad="UpdateScreen(1)">
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
				<font SIZE="+2" face="verdana, arial, helvetica"><strong>Provider Information</strong></font>
			</td>
			<td ALIGN="CENTER" WIDTH="20%">
				<img SRC="images/bluebar2.gif" onClick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0">
			</td>
		</tr>
	</table>
<%
	sSQL = "select "
	sSQL = sSQL & " c.description cred1Desc, c.description cred2desc, c.description"
	sSQL = sSQL & " cred3Desc from providers p left join Providercredentials c on"
	sSQL = sSQL & " (p.credentials=c.credential or p.credentials2=c.credential or" 
	sSQL = sSQL & " p.credentials3=c.credential)" 
	sSQL = sSQL & " WHERE p.TaxID = " & m_iTaxID & " AND p.Associate = " & m_iAssoc
	sSQL = sSQL & " and p.RecordType = 0 order by cred1Desc"
	adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
	if not adoRS.EOF then
		m_sCredentials = adoRS("cred1Desc")	
		adoRS.MoveNext
		do while not adoRS.EOF
			if trim(adoRS("cred1Desc") & " ") <> "" then
				m_sCredentials = m_sCredentials & ", " & adoRS("cred1Desc")
			end if
			adoRS.MoveNext
		loop
	end if
	adoRS.Close

	sSQL = "select p.TaxID, p.Associate, p.AlternateTaxID, p.FullName, n.Note,"
	sSQL = sSQL & " p.AddressLine1, p.AddressLine2, p.City, p.State," 
	sSQL = sSQL & " p.PostalCode, p.Credentials, p.Credentials2, p.Credentials3," 
	sSQL = sSQL & " s.Description 'Specialty Description'" 
	sSQL = sSQL & " FROM Providers p left join ProviderNotes n on (p.TaxID = n.TaxID" 
	sSQL = sSQL & " and p.Associate = n.Associate)" 
	sSQL = sSQL & " left join ProviderSpecialties s on (p.Specialty="
	sSQL = sSQL & " s.HCFASpecialty OR p.Specialty=s.PRUSpecialty)" 
	sSQL = sSQL & " WHERE p.TaxID = " & m_iTaxID & " AND p.Associate = " & m_iAssoc
	sSQL = sSQL & " and p.RecordType = 0"
	adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
	
	if adoRS.EOF then
		Response.Write ("<p><ul><font color='red' face='verdana, arial, helvetica'>Unable to find all necessary information; please contact your network administrator.</font></ul>")
		Response.Write "<BR>" & sSQL
	else
%>
		<table border="1" width="100%" COLS=6 RULES=GROUPS bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<COLGROUP SPAN=2>
			<COLGROUP SPAN=2>
			<COLGROUP SPAN=2>
		    <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				  <td WIDTH=10% bgcolor="#cccccc"><strong>Tax ID:</strong></td>
				  <td WIDTH=15% STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("TaxID")%></td></td>
		      <td WIDTH=10% bgcolor="#cccccc"><strong>Assoc. #:</strong></td>
		      <td WIDTH=10% STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("Associate")%></td>
		      <td WIDTH=10% bgcolor="#cccccc"><strong>Name:</strong></td>
		      <td WIDTH=45% STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("FullName")%></td>
		    </tr>
		</table>
		<br>
		<table cellpadding="1" cellspacing="1" width="100%" COLS=4 border="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td WIDTH="20%" VALIGN="TOP" bgcolor="#F0F0F0"><strong>Alternate Tax ID:</strong></td>
		    <td WIDTH="30%" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("AlternateTaxID")%></td>
				<td WIDTH="20%" VALIGN="TOP" bgcolor="#F0F0F0"><strong>Address</strong></td>
<%
				if trim(adoRS("AddressLine1") & " ") <> "" then
%>
				<td WIDTH="30%" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("AddressLine1")%></td>
<%
				else
%>
				<td WIDTH="30%" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("AddressLine2")%></td>
<%
				end if
%>
			</tr>
<%
				if (trim(adoRS("AddressLine1") & " ") <> "" AND trim(adoRS("AddressLine1") & " ") <> "") then
%>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
			  <td bgcolor="#F0F0F0"></td>
				<td bgcolor="#F0F0F0"></td>
				<td bgcolor="#F0F0F0"></td>
				<td WIDTH="30%" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("AddressLine2")%></td>
			</tr>
<%
				end if
%>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td VALIGN="TOP" bgcolor="#F0F0F0"><strong>Provider Specialty:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("Specialty Description")%></td>
				<td VALIGN="TOP" bgcolor="#F0F0F0"><strong>City</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("City")%></td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td VALIGN="TOP" bgcolor="#F0F0F0"><strong>Credentials:</strong></td>
			  <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= m_sCredentials%></td>
				<td VALIGN="TOP" bgcolor="#F0F0F0"><strong>State</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("State")%></td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td VALIGN="TOP" bgcolor="#F0F0F0"><strong>Phone Number:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">[To be added to Database]</td>
				<td VALIGN="TOP" bgcolor="#F0F0F0"><strong>ZIP</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("PostalCode")%></td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td VALIGN="TOP" bgcolor="#F0F0F0"><strong>Fax Number:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">[To be added to Database]</td>
        <td colspan="2" bgcolor="#F0F0F0">&nbsp;</td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td VALIGN="TOP" bgcolor="#F0F0F0"><strong>Notes:</strong></td>
		    <td COLSPAN=3 STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<%= adoRS("Note")%></td>
			</tr>
		</table>
    <br>
		<hr SIZE="3" NOSHADE>
<%
		adoRS.Close
'
' PPO Affiliates
'
		sSQL = "select p.PPOSponsor, s.Name, p.EffectiveDate, p.TermDate, p.EthixKey"
		sSQL = sSQL & " from PPOMembers p" 
		sSQL = sSQL & " left join PPO s on (p.PPOSponsor = s.PPOSponsor) WHERE"
		sSQL = sSQL & " p.TaxID = " & m_iTaxID & " AND p.Associate = " & m_iAssoc & " order by s.Name"
%>
		<br>	
		<font SIZE="3" face="arial, helvetica"><center><b>PPO Affiliates</b></center></font>
		<div align="left">
		<table border="1" cellpadding="2" cellspacing="2" width="100%" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr ALIGN="CENTER" BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td bgcolor="cccccc"><strong>PPO Name</strong></td>
				<td bgcolor="cccccc"><strong>PPO Code</strong></td>
				<td bgcolor="cccccc"><strong>PPO Key Code</strong></td>
				<td bgcolor="cccccc"><strong>Effective Date</strong></td>
				<td bgcolor="cccccc"><strong>Termination Date</strong></td>
			</tr>
<%
			adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
			if not adoRS.EOF then
				Do While Not adoRS.EOF
%>
					<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
					    <td bgcolor="#F0F0F0"><%=adoRS("Name")%></td>
					    <td bgcolor="#F0F0F0"><%=adoRS("PPOSponsor")%></td>
					    <td bgcolor="#F0F0F0"><%=adoRS("EthixKey")%></td>
					    <td bgcolor="#F0F0F0"><%=adoRS("EffectiveDate")%></td>
					    <td bgcolor="#F0F0F0"><%=adoRS("TermDate")%></td>
					</tr>
<% 
					adoRS.MoveNext
				Loop
			else
%>
				<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				    <td COLSPAN="5" bgcolor="#F0F0F0"><font color="red"><b>No PPO Affiliates found.</b></font></td>
				</tr>
<%
			end if  
%>
		</table>
		</div>
<%
		adoRS.Close
'
' Address History
'
%>
		<br>	
		<font SIZE="3" face="arial, helvetica"><center><b>Address History</b></center></font>
		<div align="left">
		<table border="1" cellpadding="2" cellspacing="2" width="100%" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr ALIGN="CENTER" BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td bgcolor="cccccc"><strong>Address</strong></td>
				<td bgcolor="cccccc"><strong>Updated On</strong></td>
				<td bgcolor="cccccc"><strong>By</strong></td>
			</tr>
<%
			sSQL = "select UserName, CONVERT(VARCHAR(12),UpdateTime,101) ChangeDate,"
			sSQL = sSQL & " AddressLine1 + ' ' + AddressLine2 Address, City, State, PostalCode"
			sSQL = sSQL & " FROM ProviderChanges"
			sSQL = sSQL & " WHERE TaxID = " & m_iTaxID & " AND Associate=" & m_iAssoc			
			sSQL = sSQL & " and RecordType = 0"
			sSQL = sSQL & " order by UpdateTime DESC"					
			adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
			if not adoRS.EOF then
				Do While Not adoRS.EOF
%>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td bgcolor="#F0F0F0"><%=adoRS("Address") + ", " & adoRS("City") + ", " & adoRS("State") + " " & adoRS("PostalCode")%></td>
				<td bgcolor="#F0F0F0"><%=adoRS("ChangeDate")%></td>
				<td bgcolor="#F0F0F0"><%=adoRS("UserName")%></td>
			</tr>
<% 
					adoRS.MoveNext
				Loop
			else
%>
				<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				    <td COLSPAN="3" bgcolor="#F0F0F0"><font color="red"><b>No address changes found.</b></font></td>
				</tr>
<%
			end if  
%>
		</table>
		</div>
<% 
		adoRS.Close
		adoConn.Close
		set adoRS = nothing
		set adoConn = nothing
%>	
		</body>
		</html>
<%
	end if
end if
%>