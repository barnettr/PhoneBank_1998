<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim adoConn, adoRS, sSQL, sTemp, sCriteria
Dim iPageNo, iRowCount, i, iPageCount
Dim m_iTaxID, m_iAssocNum, m_sFullName, m_sAddress, m_sCity, m_sState, m_sZip, m_iProvType
dim m_bUseAltTaxID, iNumProvTypes, avProvTypes()
dim bUseSQL, bContinueProcessing
dim adoCmd, adoParam1, adoParam2, adoParam3, adoParam4, adoParam5, adoParam6, adoParam7, adoParam8, adoParam9, adoParam

bContinueProcessing = true

m_sFullName = request.querystring("FullName")
m_iTaxID = request.querystring("TaxID")

set adoConn = Server.CreateObject("ADODB.Connection")
adoConn.ConnectionTimeout = 0
adoConn.CommandTimeout = 0
set adoCmd = Server.CreateObject("ADODB.Command")
adoCmd.CommandTimeout = 0
set adoRS = Server.CreateObject("ADODB.Recordset")
adoConn.Open Application("DataConn")

sSql = "select count(*) NumTypes from ProviderTypes"
adoRS.Open sSql, adoConn, adOpenKeyset, adLockOptimistic
iNumProvTypes = adoRS("NumTypes")

adoRS.Close

redim avProvTypes(1,iNumProvTypes)
sSql = "select ProvType, Description from ProviderTypes order by Description"
adoRS.Open sSql, adoConn, adOpenKeyset, adLockOptimistic
for i = 1 to iNumProvTypes
	avProvTypes(0,i) = adoRS("ProvType")
	avProvTypes(1,i) = adoRS("Description")
	adoRS.MoveNext
next

adoRS.Close

if Request.Querystring("UserCriteria") = "Initial" then
	bUseSQL = false
else
	bUseSQL = true
  
  set adoCmd = Server.CreateObject("ADODB.Command")
  adoCmd.CommandTimeout = 0
  set adoRS = Server.CreateObject("ADODB.Recordset")

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
	if Request.Querystring("UseAltTaxID") <> "" then
		m_bUseAltTaxID = true
	else
		m_bUseAltTaxID = false
	end if
	if Request.Querystring("TAXID") <> "" then
		m_iTaxID = Request.Querystring("TAXID")
  else
    m_iTaxID = 0
	end if
	if Request.Querystring("AssocNum") <> "" then
		m_iAssocNum = Request.Querystring("AssocNum")
	end if
	if Request.Querystring("FullName") <> "" then
		m_sFullName = Request.Querystring("FullName")
  else
    m_sFullName = ""
	end if
	if Request.Querystring("Address") <> "" then
		m_sAddress = Request.Querystring("Address")
	end if
	if Request.Querystring("City") <> "" then
		m_sCity = Request.Querystring("City")
	end if
	if Request.Querystring("State") <> "" then
		m_sState = Request.Querystring("State") 
	end if
	if Request.Querystring("ZIP") <> "" then
		m_sZip = Request.Querystring("ZIP")
	end if
	if Request.Querystring("ProviderType") <> "" and Request.Querystring("ProviderType") <> 0 then
		m_iProvType = Request.Querystring("ProviderType")
	else
		m_iProvType = 0
	end if

  adoCmd.CommandText = "pb_GetProviderInfo"
	adoCmd.CommandType = &H0004
	Set adoCmd.ActiveConnection = adoConn
  Set adoParam1 = adoCmd.CreateParameter("@P1_TaxID", adInteger, adParamInput)
	adoCmd.Parameters.Append adoParam1
  Set adoParam2 = adoCmd.CreateParameter("@P2_Associate", adInteger, adParamInput)
  adoCmd.Parameters.Append adoParam2
  Set adoParam3 = adoCmd.CreateParameter("@P3_FullName", adVarChar, adParamInput, 40)
  adoCmd.Parameters.Append adoParam3
  Set adoParam4 = adoCmd.CreateParameter("@P4_Address", adVarChar, adParamInput, 30)
  adoCmd.Parameters.Append adoParam4
  Set adoParam5 = adoCmd.CreateParameter("@P5_City", adVarChar, adParamInput, 28)
  adoCmd.Parameters.Append adoParam5
  Set adoParam6 = adoCmd.CreateParameter("@P6_State", adVarChar, adParamInput, 2)
  adoCmd.Parameters.Append adoParam6
  Set adoParam7 = adoCmd.CreateParameter("@P7_PostalCode", adVarChar, adParamInput, 12)
  adoCmd.Parameters.Append adoParam7
  Set adoParam8 = adoCmd.CreateParameter("@P8_ProvType", adInteger, adParamInput)
  adoCmd.Parameters.Append adoParam8
  Set adoParam9 = adoCmd.CreateParameter("@P9_UseAlternate", adBoolean, adParamInput)
  adoCmd.Parameters.Append adoParam9
  adoCmd("@P1_TaxID") = m_iTaxID
  adoCmd("@P2_Associate") = m_iAssocNum
  adoCmd("@P3_FullName") = m_sFullName
  adoCmd("@P4_Address") = m_sAddress
  adoCmd("@P5_City") = m_sCity
  adoCmd("@P6_State") = m_sState
  adoCmd("@P7_PostalCode") = m_sZip
  adoCmd("@P8_ProvType") = m_iProvType
  adoCmd("@P9_UseAlternate") = m_bUseAltTaxID

end if
%>
<html>
<head>
<title>Details</title>
</head>
<body TOPMARGIN="2" LEFTMARGIN="2" RIGHTMARGIN="0" LANGUAGE="VBScript" onLoad="UpdateScreen(1)">
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
			<font SIZE="+2" face="verdana, arial, helvetica"><strong>Provider Search</strong></font>
		</td>
		<td ALIGN="CENTER" WIDTH="20%">
			<img SRC="images/bluebar2.gif" onClick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0">
		</td>
	</tr>
</table>
<%

	if bUseSQL then
		'adoRS.Open adoCmd
    Set adoRS = adoCmd.Execute()
    if adoRS.EOF then
			bContinueProcessing = false
			adoRS.Close
			adoConn.Close
			set adoRS = nothing
			set adoConn = nothing
			if m_bUseAltTaxID then
%>
				<br><center><font color="red" face="verdana, arial, helvetica" size="2"><b>
				No matches were found -- try a different search.
				</b></font></center>
<%
			else
%>
				<br><center><font color="red" face="verdana, arial, helvetica" size="2"><b>
				No matches were found.  Try a different search,
				or click the button below to expand the search:
				</b>
				<br>
				<form ID="UseAltTaxID" METHOD="GET" ACTION="ProviderSearch.asp" TARGET="Details">
					<input TYPE="HIDDEN" ID="AltTaxID" NAME="UseAltTaxID" VALUE="TRUE">
					<input TYPE="HIDDEN" ID="TaxID" NAME="TaxID" VALUE="<%= m_iTaxID%>">
					<input TYPE="SUBMIT" NAME="AltIDResub" VALUE="Include Alternate TaxID in Search" onClick="WorkingStatus()">
					<br>(Potentially slow search)
				</form></font></center>
<%
			end if
		end if
	end if
	if bContinueProcessing then
		if bUseSQL then
			adoRS.PageSize = m_iPageSize3 ' Number of rows per page 200
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
			<form ID="Criteria" METHOD="GET" ACTION="ProviderSearch.asp" TARGET="Details">
				<table CLASS="CriteriaTable" COLS="10" CELLPADDING="0" CELLSPACING="0" border="0">
					<tr>
<!-- These TH's were required in order to get IE to deal properly with the DIV...among other things, without them and with BORDER=1, things were displayed OK, butwith the desired BORDER=0 IE seemed to 'lose' a number of the elements-->
						<th>
						<th>
						<th>
						<th>
						<th>
						<th>
						<th>
						<th>
						<th>
						<th>
					</tr>
					<tr>
						<td class="White">Tax ID:</td>
						<td><input TYPE="TEXT" ID="TAXID" NAME="TAXID" SIZE="10" MAXLENGTH=10 VALUE="<%= m_iTaxID %>"></td>
						<td class="White">Assoc #:</td>
						<td><input TYPE="TEXT" ID="ASSOCNUM" NAME="ASSOCNUM" SIZE="5" MAXLENGTH=10 VALUE="<%= m_iAssocNum %>"></td>
						<td class="White">Name:</td>
						<td><input TYPE="TEXT" ID="FullName" NAME="FullName" SIZE="30" MAXLENGTH=40 VALUE="<%= m_sFullName %>"></td>
						<td class="White">Address:</td>
						<td><input TYPE="TEXT" ID="Address" NAME="Address" SIZE="15" MAXLENGTH=30 VALUE="<%= m_sAddress %>"></td>
					</tr>
					<tr>
						<td COLSPAN="6" class="White">
							<input TYPE="CHECKBOX" ID="UseAltTaxID" NAME="UseAltTaxID" <%
							if bUseSQL then
								if m_bUseAltTaxID then
									Response.Write " CHECKED "
								end if
							end if
%>>
							Check here to also search for Alternate Tax ID (slows search)</td>
						<td class="White">City:</td>
						<td><input TYPE="TEXT" ID="City" NAME="City" SIZE="15" MAXLENGTH=28 VALUE="<%= m_sCity%>"></td>
					</tr>
					<tr>
						<td class="White">Provider Type:</td>
						<td COLSPAN="5">
							<select NAME="ProviderType" ID="ProviderType">
								<option VALUE="0">
<%
								for i = 1 to iNumProvTypes
									if avProvTypes(0,i) <> m_iProvType then
%>
										<option VALUE="<%= avProvTypes(0,i)%>"><%= avProvTypes(1,i)%>
<%
									else
%>
										<option VALUE="<%= avProvTypes(0,i)%>" SELECTED><%= avProvTypes(1,i)%>
<%
									end if
								next
%>
							</select></td>
						<td class="White">State:</td>
						<td><input TYPE="TEXT" ID="State" NAME="State" SIZE="3" MAXLENGTH=12 VALUE="<%= m_sState%>"></td>
						<td class="White">Zip Code:</td>
						<td><input TYPE="TEXT" ID="ZIP" NAME="ZIP" SIZE="12" MAXLENGTH=12 VALUE="<%= m_sZip%>"></td>
					</tr>
				</table>
				<input TYPE="HIDDEN" ID="ActionType" NAME="ActionType" VALUE>
			</form>
		</div>	
		<br CLEAR="LEFT">
<%
		if bUseSQL then
%>
			<div STYLE="height=68%; width=100%;">
				<table WIDTH="100%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
					<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
<%
						For i = 0 to adoRS.Fields.Count - 1 
%>
							<td ALIGN="CENTER" NOWRAP bgcolor="#cccccc">
								<font COLOR="BLUE"><b><%=adoRS(i).Name %></b></font>
							</td>
<%
						Next
%>
					</tr>
<% 
					iRowCount = adoRS.PageSize
					Do While Not adoRS.EOF and iRowCount > 0 
%>
						<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
							<td NOWRAP bgcolor="#F0F0F0">
								<a HREF="ProviderDetails.asp?TaxID=<%= adoRS("TAXID")%>&amp;AssocNo=<%= adoRS("Assoc #")%>" onClick="WorkingStatus()">
<% 
								Response.Write adoRS(0)
%>
								</a>
							</td>
<%
							For i = 1 to adoRS.Fields.Count - 1 
								if adoRS(i).type = adCurrency then
									Response.Write "<td ALIGN=RIGHT NOWRAP bgcolor='#F0F0F0'>" & formatcurrency(adoRS(i)) & "</td>"
								else 
									Response.Write "<td NOWRAP bgcolor='#F0F0F0'>" & adoRS(i) & "</td>"
								end if
							Next
%>
						</tr>
<%
						iRowCount = iRowCount - 1
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
	dim bValidData, iTemp

		Criteria.TaxID.Value = trim(Criteria.TaxID.Value)
		Criteria.AssocNum.Value = trim(Criteria.AssocNum.Value)
		Criteria.FullName.Value = trim(Criteria.FullName.Value)
		Criteria.ProviderType.Value = trim(Criteria.ProviderType.Value)
		Criteria.Address.Value = trim(Criteria.Address.Value)
		Criteria.City.Value = trim(Criteria.City.Value)
		Criteria.State.Value = trim(Criteria.State.Value)
		Criteria.Zip.Value = trim(Criteria.Zip.Value)
		bValidData = false
		if Criteria.TaxID.Value <> "" then
			if not isNumeric(Criteria.TaxID.Value) then
				msgbox "Please enter a valid Tax ID.",,"Invalid Data"
				exit sub
			else
				if Criteria.TaxID.Value > <%= Application("IntMax")%> or _
						Criteria.TaxID.Value < <%= Application("IntMin")%> then
					msgbox "Please enter a valid Tax ID.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.AssocNum.Value <> "" then
			if not isNumeric(Criteria.AssocNum.Value) then
				msgbox "Please enter a valid Associate Number.",,"Invalid Data"
				exit sub
			else
				if Criteria.AssocNum.Value > <%= Application("IntMax")%> or _
						Criteria.AssocNum.Value < <%= Application("IntMin")%> then
					msgbox "Please enter a valid Associate Number.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.FullName.Value <> "" then
			if ContainsInvalids(Criteria.FullName.Value) then
				msgbox "Please remove invalid characters from the Provider Name field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.ProviderType.Value <> "0" then
			bValidData = true
		end if
		if Criteria.Address.Value <> "" then
			if ContainsInvalids(Criteria.Address.Value) then
				msgbox "Please remove invalid characters from the Address field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.City.Value <> "" then
			if ContainsInvalids(Criteria.City.Value) then
				msgbox "Please remove invalid characters from the City field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.State.Value <> "" then
			if ContainsInvalids(Criteria.State.Value) then
				msgbox "Please remove invalid characters from the State field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Zip.Value <> "" then
			iTemp = replace(Criteria.Zip.Value,"-","")
			if not isNumeric(iTemp) then
				msgbox "Please enter a valid Zip Code.",,"Invalid Data"
				exit sub
			else
				if iTemp > <%= Application("IntMax")%> or _
						iTemp < <%= Application("IntMin")%> then
					msgbox "Please enter a valid Zip Code.",,"Invalid Data"
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
		if "<%= m_iTaxID%>" <> Criteria.TaxID.Value or _
				"<%= m_iAssocNum%>" <> Criteria.AssocNum.Value or _
				"<%= m_sFullName%>" <> Criteria.FullName.Value or _
				"<%= m_iProvType%>" <> Criteria.ProviderType.Value or _
				"<%= m_sAddress%>" <> Criteria.Address.Value or _
				"<%= m_sCity%>" <> Criteria.City.Value or _
				"<%= m_sState%>" <> Criteria.State.Value or _
				"<%= m_sZip%>" <> Criteria.Zip.Value then
			Criteria.ActionType.value=""
		else
			Criteria.ActionType.value=iIndex
		end if
		Criteria.submit
	end sub
</script>
</html>
