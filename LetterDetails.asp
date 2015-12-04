<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim i, adoConn, adoRS, sSQL, sTemp
dim m_sCreated, m_iSeq, adoCmd, adoParam1, adoParam2
dim adoParam3, adoParam4

m_sCreated = Request.QueryString("DateCreated")
m_iSeq = Request.QueryString("Seq")

if m_sCreated = "" or m_iSeq = "" then
	Response.Write ("<p><ul><font color='red' face='verdana, arial, helvetica' size='2'><b>Creation Date or Sequence Number missing; contact your network administrator.</b></font></ul>")
else
  
  set adoConn = Server.CreateObject("ADODB.Connection")
  adoConn.ConnectionTimeout = 0
	adoConn.CommandTimeout = 0
  set adoCmd = Server.CreateObject("ADODB.Command")
  adoCmd.CommandTimeout = 0
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
  
  adoCmd.CommandText = "PB_ListLettersSent"
  adoCmd.CommandType = adCmdStoredProc
  Set adoCmd.ActiveConnection = adoConn
  Set adoParam1 = adoCmd.CreateParameter("@P1_DateCreated", adChar, adParamInput, 8)
  adoCmd.Parameters.Append adoParam1
  Set adoParam2 = adoCmd.CreateParameter("@P2_Sequence", adInteger, adParamInput)
  adoCmd.Parameters.Append adoParam2
  adoCmd("@P1_DateCreated") = m_sCreated
  adoCmd("@P2_Sequence") = m_iSeq
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
				<font SIZE="+2" face="verdana, arial, helvetica"><strong>Letter Information</strong></font>
			</td>
			<td ALIGN="CENTER" WIDTH="20%">
				<img SRC="images/bluebar2.gif" onClick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0">
			</td>
		</tr>
	</table>
<%
	adoRS.Open adoCmd
	if adoRS.EOF then
		Response.Write ("<p><ul><font color='red' face='verdana, arial, helvetica' size='2'><b>Unable to find all necessary information; please contact your network administrator.</b></font></ul>")
	else
%>
		<table COLS=6 WIDTH="100%" BORDER="1" RULES=GROUPS bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<COLGROUP SPAN=2>
			<COLGROUP SPAN=2>
			<COLGROUP SPAN=2>
		    <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				  <td WIDTH="15%" bgcolor="#cccccc"><strong>Date Created:</strong></td>
				  <td STYLE='color:<%= Session("EmphColor")%>;' WIDTH="28%" bgcolor="#cccccc"><%= adoRS("DateCreated")%></td>
		      <td WIDTH="15%" bgcolor="#cccccc"><strong>Sequence:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' WIDTH="20%" bgcolor="#cccccc"><%= adoRS("Sequence")%></td>
		      <td WIDTH="15%" bgcolor="#cccccc"><strong>SSN:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' WIDTH="19%" bgcolor="#cccccc"><%= adoRS("SSN")%></td>
		    </tr>
		    <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				  <td bgcolor="#F0F0F0"><strong>Status:</strong></td>
				  <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("Status")%></td>
		      <td bgcolor="#F0F0F0"><strong>Status Date:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("StatusDate")%></td>
		      <td bgcolor="#F0F0F0"><strong>ACS User ID:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("ACSUserID")%></td>
		    </tr>
		    <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				  <td bgcolor="#cccccc"><strong>Letter Key:</strong></td>
				  <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("LetterKey")%></td>
		      <td bgcolor="#cccccc"><strong>Fund:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("Fund")%></td>
		      <td bgcolor="#cccccc"><strong>HWType:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("HWType")%></td>
		    </tr>
		    <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				  <td bgcolor="#F0F0F0"><strong>Name #1:</strong></td>
				  <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("Name_1")%></td>
		      <td bgcolor="#F0F0F0"><strong>Name #2:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("Name_2")%></td>
		      <td bgcolor="#F0F0F0"><strong>HWType:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("HWType")%></td>
		    </tr>
		</table>
		<br>	
		<table ALIGN="CENTER" BORDER="0" CELLPADDING="2" CELLSPACING="2" WIDTH="60%" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td bgcolor="#cccccc"><strong>Address:</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("Street")%></td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td bgcolor="#F0F0F0"><strong>City:</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("City")%></td>
			</tr>
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td bgcolor="#cccccc"><strong>State:</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("State")%></td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td bgcolor="#F0F0F0"><strong>ZIP Code:</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("PostalCode")%></td>
			</tr>
		</table>
<%
		adoRS.Close
%>
		<br>	
		<font SIZE="3" face="arial, helvetica"><center><b>Letter Lines</b></center></font>
		<table WIDTH="100%" BORDER="1" CELLPADDING="2" CELLSPACING="2" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr ALIGN="CENTER" BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td WIDTH="20%" bgcolor="#cccccc"><strong>Line Number</strong></td>
				<td WIDTH="80%" bgcolor="#cccccc"><strong>Line of Text</strong></td>
			</tr>
<%
      
      adoCmd.CommandText = "PB_ListLettersLines"
      adoCmd.CommandType = adCmdStoredProc
      Set adoCmd.ActiveConnection = adoConn
			adoRS.Open adoCmd
			if not adoRS.EOF then
				Do While Not adoRS.EOF
%>
					<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
					    <td bgcolor="#F0F0F0"><%= adoRS("LineNumber")%></td>
					    <td bgcolor="#F0F0F0"><%= adoRS("LineOfText")%></td>
					</tr>
<% 
					adoRS.MoveNext
				Loop
			else
%>
				<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				    <td COLSPAN="2" bgcolor="#F0F0F0"><font color="red" face="verdana, arial, helvetica" size="2"><b>No letter lines found.</b></font></td>
				</tr>
<%
			end if  
%>
		</table>
<%
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