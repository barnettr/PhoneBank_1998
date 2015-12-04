<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim adoConn, adoRS, sSQL, sTemp, sCriteria
Dim iPageNo, iRowCount, i, iPageCount
dim iNumNames, avLettNames()
Dim m_iSSN, m_iTaxID, m_sLogonID, m_sStatus, m_sFromDate, m_sThruDate, m_sFromTime, m_sThruTime
dim m_sLetterNum, m_sLetterName  
dim bUseSQL
dim Manager, Supervisor, Auditor, locktype, Admin 

set adoConn = Server.CreateObject("ADODB.Connection")
adoConn.CommandTimeout = 0
set adoRS = Server.CreateObject("ADODB.Recordset")
adoConn.Open Application("DataConn")

  sSQL = "select * from UserInformation where LogonID='" & Session("User") & "'"
  adoRS.Open sSql, adoConn, adOpenForwardOnly, adLockOptimistic

  if adoRS.state = 1 then
	  if adoRS.EOF then
		  norecords = 1
	  end if
  end if

  if not adoRS.EOF then
    Manager = adoRS("IsManager")
    Supervisor = adoRS("IsSupervisor")
    Auditor = adoRS("IsAuditor")
  end if
  adoRS.Close
  
  if Request.Querystring("SSN") <> "" then
		m_iSSN = Request.Querystring("SSN")
	end if
  sSQL = "select locktype, SupervisorAccessOnly from Names where SSN=" & m_iSSN & " And DepNumber > 0 and locktype is not null"
  adoRS.Open sSql, adoConn, adOpenForwardOnly, adLockOptimistic
  
  if not adoRS.EOF then
    locktype = adoRS("locktype")
    Admin = adoRS("SupervisorAccessOnly")
  end if
  
  adoRS.Close


sSql = "select count(*) NumNames from FormLetters"
adoRS.Open sSql, adoConn, adOpenKeyset, adLockOptimistic
iNumNames = adoRS("NumNames")
adoRS.Close
redim avLettNames(iNumNames)
sSql = "select LetterName from FormLetters order by LetterName"
adoRS.Open sSql, adoConn, adOpenKeyset, adLockOptimistic
for i = 1 to iNumNames
	avLettNames(i) = adoRS("LetterName")
	adoRS.MoveNext
next
adoRS.Close

if Request.Querystring("UserCriteria") = "Initial" then
	bUseSQL = false
else
	bUseSQL = true

	sSQL = "select f.LetterNumber, f.LetterName, f.SSN, f.LogonID, f.TaxID, f.Associate, f.Status,"
	sSQL = sSQL & " f.DateCreated, f.StatusDate"
	sSQL = sSQL & " from FormLettersSent f "
	sSQL = sSQL & " where f.Status='F' AND "

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
	if Request.Querystring("LetterNum") <> "" then
		m_sLetterNum = Request.Querystring("LetterNum")
		sSQL = sSQL & " f.LetterNumber ='" & m_sLetterNum & "' AND"
	end if
	if Request.Querystring("LetterName") <> "" and Request.Querystring("LetterName") <> "0" then
		m_sLetterName = Request.Querystring("LetterName")
		sSQL = sSQL & " f.LetterName ='" & m_sLetterName & "' AND"
	else
		m_sLetterName = 0
	end if
	if Request.Querystring("SSN") <> "" then
		m_iSSN = Request.Querystring("SSN")
		sSQL = sSQL & " f.SSN =" & m_iSSN & " AND"
	end if
	if Request.Querystring("LogonID") <> "" then
		m_sLogonID = Request.Querystring("LogonID")
		sSQL = sSQL & " f.LogonID='" & m_sLogonID & "' AND"
	end if
	'if Request.Querystring("Status") <> "" then
		'm_sStatus = Request.Querystring("Status")
		'sSQL = sSQL & " f.Status='" & m_sStatus & "' AND"
	'end if
	if Request.Querystring("TaxID") <> "" then
		m_iTaxID = Request.Querystring("TaxID")
		sSQL = sSQL & " f.TaxID=" & m_iTaxID & " AND"
	end if
	if Request.Querystring("FromDate") <> "" then
		m_sFromDate = Request.Querystring("FromDate")
	end if
	if Request.Querystring("ThruDate") <> "" then
		m_sThruDate = Request.Querystring("ThruDate")
	end if
	
	if m_sFromDate <> "" then
		sSQL = sSQL & " f.DateCreated >= '" & m_sFromDate & "' AND"
	end if
	if m_sThruDate <> "" then
		sSQL = sSQL & " f.DateCreated <= '" & dateadd("s",-1,dateadd("d",1,m_sThruDate)) & "' AND"
	end if
	
	sSQL = left(sSQL,len(sSQL)-4)
	sSQL = sSQL & " order by f.DateCreated DESC" ' or order by DateCreated DESC?
end if
%>
<html>
<head>
<title>Details</title>
<script language="javascript" src="function.js"></script>
</head>
<body TOPMARGIN="2" LEFTMARGIN="2" RIGHTMARGIN="0" LANGUAGE="VBScript" onLoad="UpdateScreen(3)">
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
			<font SIZE="+2" face="verdana, arial, helvetica"><strong>ACE Letter Search</strong></font>
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
			adoRS.Open sSql, adoConn, adOpenKeyset, adLockOptimistic
			if adoRS.EOF then
				adoRS.Close
				adoConn.Close
				set adoRS = nothing
				set adoConn = nothing
				Response.write ("<font size=2 face='verdana, arial, helvetica' color='red'><b>No matches were found -- try a different search.</b></font>")
				exit sub
			end if
			adoRS.PageSize = m_iPageSize ' Number of rows per page
			iPageCount = adoRS.PageCount
			adoRS.AbsolutePage = iPageNo
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
			<form ID="Criteria" METHOD="GET" ACTION="FormLetterSearch.asp" TARGET="Details">
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
						<td class="White">SSN:</td>
						<td><input TYPE="TEXT" ID="SSN" NAME="SSN" SIZE="10" MAXLENGTH=10 VALUE="<%= m_iSSN%>"></td>
						<td class="White">Letter Number:</td>
						<td><input TYPE="TEXT" ID="LetterNum" NAME="LetterNum" MAXLENGTH=13 SIZE="15" VALUE="<%= m_sLetterNum%>"></td>
						<td class="White">Creation From Date:</td>
						<td><input TYPE="TEXT" ID="FromDate" NAME="FromDate" SIZE="9" MAXLENGTH=10 VALUE="<%= m_sFromDate%>"></td>
					</tr>
					<tr>
						<td class="White">Provider Tax ID:</td>
						<td><input TYPE="TEXT" ID="TaxID" NAME="TaxID" SIZE="10" MAXLENGTH=10 VALUE="<%= m_iTaxID%>"></td>
						<td class="White">Letter Status:</td>
						<!--- <td><input TYPE="TEXT" ID="Status" NAME="Status" SIZE="3" MAXLENGTH=1 VALUE="<%= m_sStatus%>"></td> --->
						<td class="White">Creation Thru Date:</td>
						<td><input TYPE="TEXT" ID="ThruDate" NAME="ThruDate" SIZE="9" MAXLENGTH=10 VALUE="<%= m_sThruDate%>"></td>
					</tr>
					<tr>
						<td class="White">Letter Name:</td>
						<td COLSPAN="3">
							<select NAME="LetterName">
								<option VALUE="0">
<%
								for i = 1 to iNumNames
									if avLettNames(i) <> m_sLetterName then
%>
										<option VALUE="<%= avLettNames(i)%>"><%= avLettNames(i)%>
<%
									else
%>
										<option VALUE="<%= avLettNames(i)%>" SELECTED><%= avLettNames(i)%>
<%
									end if
								next
%>
							</select>
						</td>
						<td class="White">Logon ID:</td>
						<td><input TYPE="TEXT" ID="LogonID" NAME="LogonID" SIZE="15" MAXLENGTH=30 VALUE="<%= m_sLogonID%>"></td>
					</tr>
				</table>
				<input TYPE="HIDDEN" ID="ActionType" NAME="ActionType" VALUE>
			</form>
		</div>
		<br CLEAR="LEFT">
<%
		if bOpenRS then
      if locktype <> "" then
        If Manager = "True" or Supervisor = "True" or Auditor = "True" Then
%>		
			<table width="100%" border="0">
        <tr>
          <td><font size="2" face="verdana, arial, helvetica"><b><% if locktype <> "" then %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="images/lock1.gif" width="32" height="32" alt="LOCKED FILE!!"><font size="2" face="verdana, arial, helvetica" color="red"><b>&nbsp;&nbsp;&nbsp;LOCKED "<%= locktype %>"  RECORDS!!!</b></font><% end if %></td>
        </tr>
      </table>
      
      <div STYLE="height=68%; width=100%; overflow=auto">
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
								<a HREF="FormLetterDetails.asp?LetterNum=<%= adoRS("LetterNumber")%>" onClick="WorkingStatus()">
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
		    else 'Manager is not true
%>
      <div STYLE="height=65%; width=100%; overflow=auto">
				<table WIDTH="100%" BORDER="1" bordercolor="white" bgcolor="white" style="font:10pt verdana, arial, helvetica, sans-serif">
					<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
            <td ALIGN="CENTER" NOWRAP bgcolor="#F0F0F0"><font COLOR="Red"><b>The Participant or one of the Dependents has a Locked File. Letters are not available. Please see your Supervisor.</b></font></td>
          </tr>
        </table>
      </div>
<%  
        end if
      else 'locktype is not true
%>
      <div STYLE="height=68%; width=100%; overflow=auto">
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
								<a HREF="FormLetterDetails.asp?LetterNum=<%= adoRS("LetterNumber")%>" onClick="WorkingStatus()">
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
    end if
	end sub
%>
</body>
<!--#include file="VBFuncs.inc" -->
<script LANGUAGE="VBSCRIPT">
	
	sub ValidSend(iIndex)
	dim bValidData

		Criteria.SSN.Value = trim(Criteria.SSN.Value)
		Criteria.TaxID.Value = trim(Criteria.TaxID.Value)
		Criteria.LogonID.Value = trim(Criteria.LogonID.Value)
		Criteria.LetterNum.Value = trim(Criteria.LetterNum.Value)
		Criteria.LetterName.Value = trim(Criteria.LetterName.Value)
		'Criteria.Status.Value = trim(Criteria.Status.Value)
		Criteria.FromDate.Value = trim(Criteria.FromDate.Value)
		Criteria.ThruDate.Value = trim(Criteria.ThruDate.Value)
		
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
		if Criteria.LetterNum.Value <> "" then
			if ContainsInvalids(Criteria.LetterNum.Value) then
				msgbox "Please remove invalid characters from the Letter Number field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.TaxID.Value <> ""  then
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
		if Criteria.LogonID.Value <> "" then
			if ContainsInvalids(Criteria.LogonID.Value) then
				msgbox "Please remove invalid characters from the Logon ID field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		'if Criteria.Status.Value <> "" then
			'if ContainsInvalids(Criteria.Status.Value) then
				'msgbox "Please remove invalid characters from the Letter Status field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
		if Criteria.FromDate.Value <> "" then
			if not isDate(Criteria.FromDate.Value) then
				msgbox "Please enter a valid date for the From Date field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.LetterName.Value <> "0" then
			bValidData = true
		end if
		if Criteria.ThruDate.Value <> "" then
			if not isDate(Criteria.ThruDate.Value) then
				msgbox "Please enter a valid date for the Thru Date field.",,"Invalid Data"
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
			"<%= m_iTaxID%>" <> Criteria.TaxID.Value or _
			"<%= m_sLogonID%>" <> Criteria.LogonID.Value or _
			"<%= m_sLetterNum%>" <> Criteria.LetterNum.Value or _
			"<%= m_sLetterName%>" <> Criteria.LetterName.Value or _
			"<%= m_sFromDate%>" <> Criteria.FromDate.Value or _
			"<%= m_sThruDate%>" <> Criteria.ThruDate.Value then
			Criteria.ActionType.value=""
		else
			Criteria.ActionType.value=iIndex
		end if
		Criteria.submit
	end sub
	
</script>
</html>
