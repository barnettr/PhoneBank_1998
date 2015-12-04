<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim adoConn, adoRS, sSQL, sTemp, sCriteria
Dim iPageNo, iRowCount, i, iPageCount
Dim m_sCheckAcct, m_iCheckNum, m_sPayee, m_iSSN
dim bUseSQL, bContinueProcessing, adoCmd, adoParam1
dim adoParam2, adoParam3, adoParam4


bContinueProcessing = true

function ConvertToClaimNum(LogDate,Sequence,Distr)
dim sClaimNum

	if trim(LogDate & " ") = "" then
		sClaimNum = ""
	else
		if month(LogDate) < 10 then
			sClaimNum = "0" & month(LogDate)
		else
			sClaimNum = month(LogDate)
		end if
		if day(LogDate) < 10 then
			sClaimNum = sClaimNum & "0" & day(LogDate)
		else
			sClaimNum = sClaimNum & day(LogDate)
		end if
		sClaimNum = sClaimNum & right(Year(LogDate),2)
		sClaimNum = sClaimNum & "_" & string(5-len(Sequence),"0") & Sequence
		sClaimNum = sClaimNum & "_" & Distr
	end if
	ConvertToClaimNum = sClaimNum
	
end function

if Request.Querystring("UserCriteria") = "Initial" then
	bUseSQL = false
else
	bUseSQL = true
  m_iSSN = request.querystring("SSN")
  m_iCheckNum = request.querystring("CheckNum")
  
  set adoConn = Server.CreateObject("ADODB.Connection")
  adoConn.ConnectionTimeout = 0
	adoConn.CommandTimeout = 0
  set adoCmd = Server.CreateObject("ADODB.Command")
  adoCmd.CommandTimeout = 0
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
	
	if Trim(Request.Querystring("ActionType")) = "0" then
		iPageNo = 1
	else
		if Trim(Request.Querystring("ActionType")) = "" then
			iPageNo = 1
		else
			iPageNo = Cint(Request.Querystring("ActionType"))
			if iPageNo < 1 Then 
				iPageNo = 1
			end if
		end if
	end if
	if Request.Querystring("CheckAcct") <> "" then
		m_sCheckAcct = Request.Querystring("CheckAcct")
	end if
	if Request.Querystring("CheckNum") <> "" then
		m_iCheckNum = Request.Querystring("CheckNum")
	end if
	if Request.Querystring("Payee") <> "" then
		m_sPayee = Request.Querystring("Payee")
	end if
	if Request.Querystring("SSN") <> "" then
		m_iSSN = Request.Querystring("SSN")
  else
   m_iSSN = 0
	end if
  
  adoCmd.CommandText = "pb_CheckSearch"
  adoCmd.CommandType = &H0004
  Set adoCmd.ActiveConnection = adoConn
  Set adoParam1 = adoCmd.CreateParameter("@P1_Acct", adVarChar, adParamInput, 6)
  adoCmd.Parameters.Append adoParam1
  Set adoParam2 = adoCmd.CreateParameter("@P2_Number", adInteger, adParamInput)
  adoCmd.Parameters.Append adoParam2
  Set adoParam3 = adoCmd.CreateParameter("@P3_Payee", adInteger, adParamInput)
  adoCmd.Parameters.Append adoParam3
  Set adoParam4 = adoCmd.CreateParameter("@P4_SSN", adInteger, adParamInput)
  adoCmd.Parameters.Append adoParam4
  adoCmd("@P1_Acct") = m_sCheckAcct
  adoCmd("@P2_Number") = m_iCheckNum
  adoCmd("@P3_Payee") = m_sPayee
  adoCmd("@P4_SSN") = m_iSSN 
  
end if
%>
<html>
<head>
<title>Details</title>
</head>
<body TOPMARGIN="2" LEFTMARGIN="2" RIGHTMARGIN="0" LANGUAGE="VBScript" onLoad="UpdateScreen(2)">
<link REL="STYLESHEET" HREF="styles/CritTable.css">
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
			<font SIZE="+2" face="verdana, arial, helvetica"><strong>Check Search</strong></font>
		</td>
		<td ALIGN="CENTER" WIDTH="20%">
			<img SRC="images/bluebar2.gif" onClick="history.back">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0">
		</td>
	</tr>
</table>
<%
	if bUseSQL then
		adoRS.Open adoCmd
		if adoRS.EOF then
			bContinueProcessing = false
			adoRS.Close
			adoConn.Close
			set adoRS = nothing
			set adoConn = nothing
			Response.write ("<p><ul><font color='red' face='verdana, arial, helvetica' size='2'><b>No matches were found -- try a different search.</b></font></ul>")
		end if
	end if
	if bContinueProcessing then
		if bUseSQL then
			adoRS.PageSize = m_iPageSize ' Number of rows per page
			iPageCount = adoRS.PageCount
			'adoRS.AbsolutePage = iPageNo
		end if
%>
		<table COLS="3" WIDTH="100%" CELLPADDING="0" CELLSPACING="0" BORDER="0">
			<tr>
				<td WIDTH="10%"><input TYPE="IMAGE" SRC="images/query_button.gif" ACCESSKEY="S" NAME="NewQuery" VALUE="Start Query" onClick="ValidSend(0)"></td>
				<td WIDTH="65%"><font face="verdana, arial, helvetica" size="2"><b>Enter your criteria, and then click the Start Query button or key (ALT+S).</b></font></td>
				<td ALIGN="RIGHT" NAME="MorePages" ID="MorePages">&nbsp;
<%
				if bUseSQL and iPageCount > 1 then
					sTemp = "<FONT SIZE=2 face='verdana, arial, helvetica'><b>Page " & iPageNo & " of " & iPageCount & "</b></FONT>&nbsp;&nbsp;"
					if iPageNo > 1 Then
						sTemp = sTemp & "<INPUT ALIGN=RIGHT TYPE=BUTTON NAME=ScrollAction VALUE='Page " & iPageNo-1 & "' onClick='ValidSend(" & iPageNo-1 & ")'>"
					else
						sTemp = sTemp & "<INPUT STYLE='visibility:hidden;' ALIGN=RIGHT TYPE=BUTTON VALUE='Page " & iPageNo-1 & "' >"
					end if
					if iPageNo < iPageCount Then
						sTemp = sTemp & "<INPUT ALIGN=RIGHT TYPE=BUTTON NAME=ScrollAction VALUE='Page " & iPageNo+1 & "' onClick='ValidSend(" & iPageNo+1 & ")'>"
					else
						sTemp = sTemp & "<INPUT STYLE='visibility:hidden;' ALIGN=RIGHT TYPE=BUTTON VALUE='Page " & iPageNo+1 & "' >"
					end if
					sTemp = sTemp & "</TD>"
					Response.Write sTemp
				end if
%>
			</tr>
		</table>
		<div STYLE="height=13%; width:100%; overflow=auto;">
			<form ID="Criteria" METHOD="GET" ACTION="CheckSearch.asp" TARGET="Details">
				<table CLASS="CriteriaTable" COLS="4" CELLPADDING="0" CELLSPACING="0">
					<tr>
<!-- These TH's were required in order to get IE to deal properly with the DIV...among other things, without them and with BORDER=1, things were displayed OK, butwith the desired BORDER=0 IE seemed to 'lose' a number of the elements-->
						<th>
						<th>
						<th>
						<th>
					</tr>
					<tr>
						<td class="White">Check Account:</td>
						<td><!--- <input TYPE="TEXT" ID="CheckAcct" NAME="CheckAcct" SIZE="8" MAXLENGTH=6 VALUE="<%= m_sCheckAcct%>"> --->
                <select name="CheckAcct" ID="CheckAcct">
                  <option value="">--Choose One--</option>
                  <option value="005CLM" <% if Request.QueryString("m_sCheckAcct") = "005CLM" then %>Selected<% end if %>>005CLM</option>
                  <option value="008DEN" <% if Request.QueryString("m_sCheckAcct") = "008DEN" then %>Selected<% end if %>>008DEN</option>
                  <option value="021CLM" <% if Request.QueryString("m_sCheckAcct") = "021CLM" then %>Selected<% end if %>>021CLM</option>
                  <option value="021F" <% if Request.QueryString("m_sCheckAcct") = "021F" then %>Selected<% end if %>>021F</option>
                  <option value="021RX" <% if Request.QueryString("m_sCheckAcct") = "021RX" then %>Selected<% end if %>>021RX</option>
                  <option value="027CLM" <% if Request.QueryString("m_sCheckAcct") = "027CLM" then %>Selected<% end if %>>027CLM</option>
                  <option value="028MED" <% if Request.QueryString("m_sCheckAcct") = "028MED" then %>Selected<% end if %>>028MED</option>
                  <option value="028RX" <% if Request.QueryString("m_sCheckAcct") = "028RX" then %>Selected<% end if %>>028RX</option>
                  <option value="055CLM" <% if Request.QueryString("m_sCheckAcct") = "055CLM" then %>Selected<% end if %>>055CLM</option>
                  <option value="055F" <% if Request.QueryString("m_sCheckAcct") = "055F" then %>Selected<% end if %>>055F</option>
                  <option value="21RRX" <% if Request.QueryString("m_sCheckAcct") = "21RRX" then %>Selected<% end if %>>21RRX </option>
                  <option value="39RMED" <% if Request.QueryString("m_sCheckAcct") = "39RMED" then %>Selected<% end if %>>39RMED</option>
                  <option value="39RRX" <% if Request.QueryString("m_sCheckAcct") = "39RRX" then %>Selected<% end if %>>39RRX</option>
                  <option value="FRTALL" <% if Request.QueryString("m_sCheckAcct") = "FRTALL" then %>Selected<% end if %>>FRTALL</option>
                  <option value="FRTCLM" <% if Request.QueryString("m_sCheckAcct") = "FRTCLM" then %>Selected<% end if %>>FRTCLM</option>
                  <option value="FRTOTH" <% if Request.QueryString("m_sCheckAcct") = "FRTOTH" then %>Selected<% end if %>>FRTOTH</option>
                  <option value="FRTRX" <% if Request.QueryString("m_sCheckAcct") = "FRTRX" then %>Selected<% end if %>>FRTRX</option>
                  <option value="OLDFRT" <% if Request.QueryString("m_sCheckAcct") = "OLDFRT" then %>Selected<% end if %>>OLDFRT</option>
                  <option value="NBNVIS" <% if Request.QueryString("m_sCheckAcct") = "NBNVIS" then %>Selected<% end if %>>NBNVIS</option>
                  <option value="NBNEYE" <% if Request.QueryString("m_sCheckAcct") = "NBNEYE" then %>Selected<% end if %>>NBNEYE</option>
                  <option value="CSBCLM" <% if Request.QueryString("m_sCheckAcct") = "CSBCLM" then %>Selected<% end if %>>CSBCLM</option>
                </select></td>
						<td class="White">Payee Name:</td>
						<td><input TYPE="TEXT" ID="Payee" NAME="Payee" SIZE="30" MAXLENGTH=30 VALUE="<%= m_sPayee%>"></td>
					</tr>
					<tr>
						<td class="White">Check Number:</td>
						<td><input TYPE="TEXT" ID="CheckNum" NAME="CheckNum" SIZE="10" MAXLENGTH=10 VALUE="<%= m_iCheckNum%>"></td>
						<td class="White">Participant SSN:</td>
						<td><input TYPE="TEXT" ID="SSN" NAME="SSN" SIZE="10" MAXLENGTH=10 VALUE="<%= m_iSSN%>"></td>
					</tr>
				</table>
				<input TYPE="HIDDEN" ID="ActionType" NAME="ActionType" VALUE>
			</form>
		</div>
		<br CLEAR="LEFT">
<%
		if bUseSQL then
%>
			<p><font size="2" face="verdana, arial, helvetica"><b>To see the details for the check you are searching for please click the appropiate check number.</b></font>
      <div STYLE="height=68%; width=100%; overflow=auto;">
				<table WIDTH="100%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
					<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
            <td ALIGN=CENTER NOWRAP bgcolor='#cccccc'><FONT COLOR=BLUE><B>Account, Check #</B></FONT></TD>
            <td ALIGN=CENTER NOWRAP bgcolor='#cccccc'><FONT COLOR=BLUE><B>Claim Number</B></FONT></TD>
            <td ALIGN=CENTER NOWRAP bgcolor='#cccccc'><FONT COLOR=BLUE><B>SSN</B></FONT></TD>
            <td ALIGN=CENTER NOWRAP bgcolor='#cccccc'><FONT COLOR=BLUE><B>Dep #</B></FONT></TD>
            <td ALIGN=CENTER NOWRAP bgcolor='#cccccc'><FONT COLOR=BLUE><B>Issue Date</B></FONT></TD>
            <td ALIGN=CENTER NOWRAP bgcolor='#cccccc'><FONT COLOR=BLUE><B>Payee</B></FONT></TD>
            <td ALIGN=CENTER NOWRAP bgcolor='#cccccc'><FONT COLOR=BLUE><B>Amount</B></FONT></TD>
          </tr>
<%  
          iRowCount = adoRS.PageSize
					Do While Not adoRS.EOF and iRowCount > 0

%>

					<TR BorderColorDark=FCFCFC BorderColorlight=CCCCCC>
            <TD NOWRAP bgcolor=F0F0F0><A HREF="CheckDetails.asp?CheckAcct=<%= adoRS("Account") %>&CheckNum=<%= adoRS("Check #") %>" onClick="WorkingStatus()"><%= adoRS("Account") %>, <%= adoRS("Check #") %></a></td>
            <TD NOWRAP bgcolor=F0F0F0><%= ConvertToClaimNum(adoRS("Claim #"),adoRS("Sequence"),adoRS("Distr")) %></td>
            <TD NOWRAP bgcolor=F0F0F0><%= adoRS("SSN") %></td>
            <TD NOWRAP bgcolor=F0F0F0><%= adoRS("Dep. #") %></td>
            <TD NOWRAP bgcolor=F0F0F0><%= adoRS("Issue Date") %></td>
            <TD NOWRAP bgcolor=F0F0F0><%= adoRS("Payee") %></td>
            <TD NOWRAP bgcolor=F0F0F0><% if adoRS("Amount") <> "" then %><%= formatcurrency(adoRS("Amount")) %><% else %>&nbsp;<% end if %></td>
          </tr>
<% 

					adoRS.MoveNext
          Loop
					adoRS.Close
					adoConn.Close
					set adoRS = nothing
					set adoConn = nothing
%>
				</table>
			</div>
<%	
		end if
	end if
%>
</body>
<!--#include file="VBFuncs.inc" -->
<script LANGUAGE="VBSCRIPT">
	
	sub ValidSend(iIndex)
	dim bValidData
	dim vbOKCancel, vbOk
	
		vbOkCancel = 1
		vbOk = 1

		Criteria.CheckNum.Value = trim(Criteria.CheckNum.Value)
		Criteria.CheckAcct.Value = trim(Criteria.CheckAcct.Value)
		Criteria.Payee.Value = trim(Criteria.Payee.Value)
		Criteria.SSN.Value = trim(Criteria.SSN.Value)
		bValidData = false
		if iIndex = 0 and (Criteria.CheckNum.Value = "" or Criteria.CheckAcct.Value = "") then
			msgbox "You must enter both a Check Number and Check Account.",48,"Missing Data"
      exit sub
		end if
		if Criteria.CheckNum.Value <> ""  then
			if not isNumeric(Criteria.CheckNum.Value) then
				msgbox "Please enter a numeric Check Number.",,"Invalid Data"
				exit sub
			else
				if Criteria.CheckNum.Value > <%= Application("IntMax")%> or _
						Criteria.CheckNum.Value < <%= Application("IntMin")%> then
					msgbox "Please enter a valid Check Number.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.CheckAcct.Value <> "" then
			if ContainsInvalids(Criteria.CheckAcct.Value) then
				msgbox "Please remove invalid characters from the Check Account field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Payee.Value <> "" then
			if ContainsInvalids(Criteria.Payee.Value) then
				msgbox "Please remove invalid characters from the Payee field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.SSN.Value <> ""  then
			if not isNumeric(Criteria.SSN.Value) then
				msgbox "Please enter a valid SSN.",,"Invalid Data"
				exit sub
			else
				if Criteria.SSN.Value > <%= Application("IntMax")%> or _
						Criteria.SSN.Value < <%= Application("IntMin")%> then
					msgbox "Please enter a valid SSN.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if not bValidData then
			msgbox "You haven't entered anything to search for.",,"No Criteria"
			exit sub
		end if
		
		WorkingStatus
		if "<%= m_iCheckNum%>" <> Criteria.CheckNum.Value or _
				"<%= m_sCheckAcct%>" <> Criteria.CheckAcct.Value or _
				"<%= m_sPayee%>" <> Criteria.Payee.Value or _
				"<%= m_iSSN%>" <> Criteria.SSN.Value then
			Criteria.ActionType.value=""
		else
			Criteria.ActionType.value=iIndex
		end if
		Criteria.submit
	end sub
	
</script>
</html>
