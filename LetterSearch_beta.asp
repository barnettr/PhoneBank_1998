<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim adoConn, adoRS, sSQL, sTemp, sCriteria
Dim iPageNo, iRowCount, i, iPageCount
dim  m_iSSN, m_sFrom, m_sThru, m_sACSID, m_sStatus, m_sStatusFrom, m_sStatusThru, m_sFund, m_sHWType, m_sName1, m_sName2
dim bUseSQL 
'm_sFrom = "NULL"
'm_sThru = "NULL"
'm_sACSID = "NULL"
'm_sStatus = "NULL"
'm_sStatusFrom = "NULL"
'm_sStatusThru = "NULL"
'm_sFund = "NULL"
'm_sHWType = "NULL"
'm_sName1 = "NULL"
'm_sName2 = "NULL"
'm_sName2 = "NULL"
'm_iSSN = "NULL"

'*********************************************************************************************
'* Many changes on the page to get Stored Procedure to run:
'* First: set all form variables to "NULL" Second: all inline SQL code commented out Third:
'* Code added to the if then statements ex. m_sFrom = "'" & m_sFrom + "'"  Fourth : adoRS.AbsolutePage = iPageNo
'* has been commented out to show all results Fifth: Error checking has been commented out to accomodate
'* the Nulls. 
'*********************************************************************************************

if Request.Querystring("UserCriteria") = "Initial" then
	bUseSQL = false
else
	bUseSQL = true
	set adoConn = Server.CreateObject("ADODB.Connection")
	'******************************************************************************************
	' Setting CommandTimeout to 0 waits indefinitely for response; note that per 
	' ADO documentation Command object will not inherit the Connection setting (default
	' is 30 seconds).  This appeared to help with response issues on seasql02.
	'*******************************************************************************************
	adoConn.CommandTimeout = 0
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")

	'sSQL = "select DateCreated Created, Sequence 'Seq #', Status, StatusDate 'Status Date', Regarding SSN, Fund, HWType,"
	'sSQL = sSQL & " LetterKey, ACSUserID, Name_1 Name, Name_2 '2nd Name'"
	'sSQL = sSQL & " from LettersSent where "
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
	'if Request.Querystring("From") = "" then
    'm_sFrom = "NULL"
  'else
		'm_sFrom = Request.Querystring("From")
    'm_sFrom = "'" & m_sFrom + "'"
	'end if
	'if Request.Querystring("Thru") = "" then
    'm_sThru = "NULL"
  'else  
		'm_sThru = Request.Querystring("Thru")
    'm_sThru = "'" & m_sThru + "'"
	'end if
	if Request.Querystring("SSN") = "" then
    m_iSSN = "NULL"
  else
		m_iSSN = Request.Querystring("SSN")
		'sSQL = sSQL & " Regarding =" & m_iSSN & " AND"
	end if
	if Request.Querystring("ACSID") = "" then
    m_sACSID = "NULL"
  else
		m_sACSID = Request.Querystring("ACSID")
		'sSQL = sSQL & " ACSUserID='" & m_sACSID & "' AND"
	end if
	if Request.Querystring("Status") = "" then
    m_sStatus = "NULL"
  else
		m_sStatus = Request.Querystring("Status")
		'sSQL = sSQL & " Status='" & m_sStatus & "' AND"
	end if
	'if Request.Querystring("StatusFrom") = "" then
    'm_sStatusFrom = "NULL"
  'else
		'm_sStatusFrom = Request.Querystring("StatusFrom")
    'm_sStatusFrom = "'" & m_sStatusFrom + "'"
	'end if
	'if Request.Querystring("StatusThru") = "" then
    'm_sStatusThru = "NULL"
  'else
		'm_sStatusThru = Request.Querystring("StatusThru")
    'm_sStatusThru = "'" & m_sStatusThru + "'"
  'end if
	if Request.Querystring("Fund") = "" then
    m_sFund = "NULL"
  else
		m_sFund = Request.Querystring("Fund")
		'sSQL = sSQL & " Fund='" & m_sFund & "' AND"
	end if
	if Request.Querystring("HWType") = "" then
    m_sHWType = "NULL"
  else
		m_sHWType = Request.Querystring("HWType")
		'sSQL = sSQL & " HWType='" & m_sHWType & "' AND"
	end if
	if Request.Querystring("Name1") = "" then
    m_sName1 = "NULL"
  else
		m_sName1 = Request.Querystring("Name1")
		'sSQL = sSQL & " Name_1 LIKE '%" & m_sName1 & "%' AND"
	end if
	if Request.Querystring("Name2") = "" then
    m_sName2 = "NULL"
  else
		m_sName2 = Request.Querystring("Name2")
		'sSQL = sSQL & " Name_2 LIKE '%" & m_sName2 & "%' AND"
	end if
	
	'if m_sFrom <> "" then
		'sSQL = sSQL & " DateCreated >= '" & m_sFrom & "' AND"
	'end if
	'if m_sThru <> "" then
		'sSQL = sSQL & " DateCreated <= '" & dateadd("s",-1,dateadd("d",1,m_sThru)) & "' AND"
	'end if
	'if m_sStatusFrom <> "" then
		'sSQL = sSQL & " StatusDate >= '" & m_sStatusFrom & "' AND"
	'end if
	'if m_sStatusThru <> "" then
		'sSQL = sSQL & " StatusDate <= '" & dateadd("s",-1,dateadd("d",1,m_sStatusThru)) & "' AND"
	'end if
	'sSQL = left(sSQL,len(sSQL)-4)
	'sSQL = sSQL & " order by Created DESC, 'Seq #'"
  
  if request.querystring("From") = "" and request.querystring("Thru") = "" and request.querystring("StatusFrom") = "" and request.querystring("StatusThru") = "" then
    sSQL ="PB_GetLettersSent " & m_iSSN & ", " & m_sACSID & ", " & m_sStatus & ", " & m_sFund & ", " & m_sHWType & ", " & m_sName1 & ", " & m_sName2
  elseif m_sFrom = Request.Querystring("From") and m_sThru = request.querystring("Thru") then
    sSQL ="PB_GetLettersSent '" & m_sFrom & "', '" & m_sThru & "', " & m_iSSN & ", " & m_sACSID & ", " & m_sStatus & ", " & m_sFund & ", " & m_sHWType & ", " & m_sName1 & ", " & m_sName2
  elseif m_sStatusFrom = Request.Querystring("StatusFrom") and m_sStatusThru = Request.Querystring("StatusThru") then
    sSQL ="PB_GetLettersSent '" & m_sStatusFrom & "', '" & m_sStatusThru & "', " & m_iSSN & ", " & m_sACSID & ", " & m_sStatus & ", " & m_sFund & ", " & m_sHWType & ", " & m_sName1 & ", " & m_sName2  
  else
    sSQL ="PB_GetLettersSent '" & m_sFrom & "', '" & m_sThru & "', '" & m_sStatusFrom & "', '" & m_sStatusThru & "', " & m_iSSN & ", " & m_sACSID & ", " & m_sStatus & ", " & m_sFund & ", " & m_sHWType & ", " & m_sName1 & ", " & m_sName2
  end if    
        
    'sSQL ="PB_GetLettersSent '" & m_sFrom & "', '" & m_sThru & "', '" & m_sStatusFrom & "', '" & m_sStatusThru & "', " & m_iSSN & ", " & m_sACSID & ", " & m_sStatus & ", " & m_sFund & ", " & m_sHWType & ", " & m_sName1 & ", " & m_sName2
    response.write sSQL
end if
%>
<html>
<head>
<title>Details</title>
</head>
<body TOPMARGIN="2" LEFTMARGIN="2" RIGHTMARGIN="0" LANGUAGE="VBScript" onload="UpdateScreen(3)">
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
			<font SIZE="+2" face="verdana, arial, helvetica"><strong>Letter Search</strong></font>
		</td>
		<td ALIGN="CENTER" WIDTH="20%">
			<img SRC="images/bluebar2.gif" onClick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0">
		</td>
	</tr>
</table>
<%

	SetPage bUseSQL
	
	Sub SetPage(bOpenRS)

		if bOpenRS then
			Set adoRS = adoConn.execute(sSQL)
      'adoRS.Open sSql, adoConn, adOpenKeyset, adLockOptimistic
			if adoRS.EOF then
				adoRS.Close
				adoConn.Close
				set adoRS = nothing
				set adoConn = nothing
				Response.write ("<ul><p><font size=2 color='red' face='verdana, arial, helvetica'><b>No matches were found -- try a different search.</b></font></ul>")
				exit sub
			end if
			adoRS.PageSize = m_iPageSize3 ' Number of rows per page
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
    <!---  --->
		<div STYLE="height=24%; width:100%; overflow=auto;">
			<form ID="Criteria" METHOD="GET" ACTION="LetterSearch.asp" TARGET="Details">
				<table CLASS="CriteriaTable" COLS="6" CELLPADDING="0" CELLSPACING="0">
					<tr>
<!-- These TH's were required in order to get IE to deal properly with the DIV...among other things, without them and with BORDER=1, things were displayed OK, butwith the desired BORDER=0 IE seemed to 'lose' a number of the elements-->
						<th>
						<th>
						<th>
						<th>
						<th>
						<th>
					</tr>
					<tr>
						<td class="White">SSN: </td>
						<td><input TYPE="TEXT" ID="SSN" NAME="SSN" SIZE="10" MAXLENGTH=10 VALUE="<%= m_iSSN%>"></td>
						<td class="White">Primary Name:</td>
						<td><input TYPE="TEXT" ID="Name1" NAME="Name1" SIZE="15" MAXLENGTH=30 VALUE="<%= m_sName1%>"></td>
						<td class="White">Secondary Name:</td>
						<td><input TYPE="TEXT" ID="Name2" NAME="Name2" SIZE="15" MAXLENGTH=30 VALUE="<%= m_sName2%>"></td>
					</tr>
					<tr>
						<td class="White">ACS User ID:</td>
						<td><input TYPE="TEXT" ID="ACSID" NAME="ACSID" SIZE="7" MAXLENGTH=6 VALUE="<%= m_sACSID%>"></td>
						<td class="White">HWType:</td>
						<td><input TYPE="TEXT" ID="HWType" NAME="HWType" SIZE="3" MAXLENGTH=1 VALUE="<%= m_sHWType%>"></td>
						<td class="White">Fund:</td>
						<td><input TYPE="TEXT" ID="Fund" NAME="Fund" SIZE="5" MAXLENGTH=3 VALUE="<%= m_sFund%>"></td>
					</tr>
					<tr>
						<td class="White">Status:</td>
						<td><input TYPE="TEXT" ID="Status" NAME="Status" SIZE="3" MAXLENGTH=1 VALUE="<%= m_sStatus%>"></td>
						<td class="White">Status From:</td>
						<td><input TYPE="TEXT" ID="StatusFrom" NAME="StatusFrom" SIZE="10" MAXLENGTH=10 VALUE="<%= m_sStatusFrom%>"></td>
						<td class="White">Status Thru:</td>
						<td><input TYPE="TEXT" ID="StatusThru" NAME="StatusThru" SIZE="10" MAXLENGTH=10 VALUE="<%= m_sStatusThru%>"></td>
					</tr>
					<tr>
						<td></td>
						<td></td>
						<td class="White">Creation From:</td>
						<td><input TYPE="TEXT" ID="From" NAME="From" SIZE="10" MAXLENGTH=10 VALUE="<%= m_sFrom%>"></td>
						<td class="White">Creation Thru:</td>
						<td><input TYPE="TEXT" ID="Thru" NAME="Thru" SIZE="10" MAXLENGTH=10 VALUE="<%= m_sThru%>"></td>
					</tr>
				</table>
				<input TYPE="HIDDEN" ID="ActionType" NAME="ActionType" VALUE>
			</form>
		</div>	
		<br CLEAR="LEFT">
<%
		if bOpenRS then
%> 
			 <!--- overflow=auto --->
      <div STYLE="height=65%; width=100%;">
				<table WIDTH="100%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
					<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
<%
						iRowCount = adoRS.PageSize
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
								<a HREF="LetterDetails.asp?DateCreated=<%= adoRS("Created")%>&amp;Seq=<%= adoRS("Seq#")%>"  onClick="WorkingStatus()">
<% 
								Response.Write adoRS(0)
%>
								</a>
							</td>
<%
							For i = 1 to adoRS.Fields.Count - 1 
								if adoRS(i).type = adCurrency then
									Response.Write "<TD ALIGN=RIGHT NOWRAP bgcolor=F0F0F0>" & formatcurrency(adoRS(i)) & "</TD>"
								else 
									Response.Write "<TD NOWRAP bgcolor=F0F0F0>" & adoRS(i) & "</TD>"
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
	end sub
%>
</body>
<!--#include file="VBFuncs.inc" -->
<script LANGUAGE="VBSCRIPT">
	
	sub ValidSend(iIndex)
	dim bValidData

		Criteria.SSN.Value = trim(Criteria.SSN.Value)
		Criteria.ACSID.Value = trim(Criteria.ACSID.Value)
		Criteria.Fund.Value = trim(Criteria.Fund.Value)
		Criteria.HWType.Value = trim(Criteria.HWType.Value)
		Criteria.Name1.Value = trim(Criteria.Name1.Value)
		Criteria.Name2.Value = trim(Criteria.Name2.Value)
		Criteria.From.Value = trim(Criteria.From.Value)
		Criteria.Thru.Value = trim(Criteria.Thru.Value)
		Criteria.Status.Value = trim(Criteria.Status.Value)
		Criteria.StatusFrom.Value = trim(Criteria.StatusFrom.Value)
		Criteria.StatusThru.Value = trim(Criteria.StatusThru.Value)
		
		bValidData = false
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
		if Criteria.ACSID.Value <> "" then
			if ContainsInvalids(Criteria.ACSID.Value) then
				msgbox "Please remove invalid characters from the ACS ID field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Fund.Value <> "" then
			if ContainsInvalids(Criteria.Fund.Value) then
				msgbox "Please remove invalid characters from the Fund field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.HWType.Value <> "" then
			if ContainsInvalids(Criteria.HWType.Value) then
				msgbox "Please remove invalid characters from the HWType field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Name1.Value <> "" then
			if ContainsInvalids(Criteria.Name1.Value) then
				msgbox "Please remove invalid characters from the Primary Name field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Name2.Value <> "" then
			if ContainsInvalids(Criteria.Name2.Value) then
				msgbox "Please remove invalid characters from the Secondary Name field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.From.Value <> "" then
			if not isDate(Criteria.From.Value) then
				msgbox "Please enter a valid date for the Creation From field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Thru.Value <> "" then
			if not isDate(Criteria.Thru.Value) then
				msgbox "Please enter a valid date for the Creation Thru field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Status.Value <> "" then
			if ContainsInvalids(Criteria.Status.Value) then
				msgbox "Please remove invalid characters from the Status field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.StatusFrom.Value <> "" then
			if not isDate(Criteria.StatusFrom.Value) then
				msgbox "Please enter a valid date for the Status From field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.StatusThru.Value <> "" then
			if not isDate(Criteria.StatusThru.Value) then
				msgbox "Please enter a valid date for the Status Thru field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if not bValidData then
			msgbox "You haven't entered anything to search for.",,"No Criteria"
			exit sub
		end if
		
		WorkingStatus
		if "<%= m_iSSN%>" <> Criteria.SSN.Value or _
			"<%= m_sACSID%>" <> Criteria.ACSID.Value or _
			"<%= m_sFund%>" <> Criteria.Fund.Value or _
			"<%= m_sHWType%>" <> Criteria.HWType.Value or _
			"<%= m_sName1%>" <> Criteria.Name1.Value or _
			"<%= m_sName2%>" <> Criteria.Name2.Value or _
      "<%= m_sFrom%>" <> Criteria.From.Value or _
			"<%= m_sThru%>" <> Criteria.Thru.Value or _
			"<%= m_sStatus%>" <> Criteria.Status.Value or _
			"<%= m_sStatusFrom%>" <> Criteria.StatusFrom.Value or _
			"<%= m_sStatusThru%>" <> Criteria.StatusThru.Value then
			Criteria.ActionType.value=""
		else
			Criteria.ActionType.value=iIndex
		end if
		Criteria.submit
	end sub
	
</script>
</html>
