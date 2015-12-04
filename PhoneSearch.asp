<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim adoConn, adoRS, sSQL, sTemp, sCriteria
Dim iPageNo, iRowCount, i, iPageCount, m_bIsSupervisor, m_sCallBack, m_sSummCode
Dim m_iSSN, m_sLogonID, m_sACSUserID, m_sFromDate, m_sThruDate, m_sFromTime, m_sThruTime 
dim bUseSQL, iRemainder, ClaimNum, LogType
dim Manager, Supervisor, Auditor

ClaimNum = request.querystring("ClaimNum")
m_iSSN = request.querystring("SSN")

if Request.Querystring("UserCriteria") = "Initial" then
	bUseSQL = false
else
	bUseSQL = true
	set adoConn = Server.CreateObject("ADODB.Connection")
	'************************************************************************************
	' Setting CommandTimeout to 0 waits indefinitely for response; note that per 
	' ADO documentation Command object will not inherit the Connection setting (default
	' is 30 seconds).  This appeared to help with response issues on seasql02.
	'************************************************************************************
	adoConn.CommandTimeout = 0
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
	
	sSQL = "select IsSupervisor from UserInformation where LogonID='" & Session("User") & "'"
	adoRS.Open sSql, adoConn, adOpenForwardOnly, adLockOptimistic
	if adoRS.EOF then
		m_bIsSupervisor = false
	else
		m_bIsSupervisor = adoRS("IsSupervisor")
	end if
  
	adoRS.Close
  
  sSQL = "select IsSupervisor, IsManager, IsAuditor from UserInformation where LogonID='" & Session("User") & "'"
	adoRS.Open sSql, adoConn, adOpenForwardOnly, adLockOptimistic
  if not adoRS.EOF then
    Manager = adoRS("IsManager")
    Supervisor = adoRS("IsSupervisor")
    Auditor = adoRS("IsAuditor")
  end if
  
  adoRS.Close

	sSQL = "select p.SSN, p.TimeOfCall, convert(varchar(26),p.TimeOfCall,9) fulltime,"
	sSQL = sSQL & " p.CallBack, p.CallSummaryCode, p.NameOfCaller,"
	sSQL = sSQL & " p.PhoneNumber, p.Extension, p.NoteText, p.WhoTookCall,u.LogonID,"
	sSQL = sSQL & " u.ACSUserID, u.FirstName + ' ' + u.MiddleName + ' ' + u.LastName UserName,"
	sSQL = sSQL & " n.FirstName + ' ' + n.MiddleName + ' ' + n.LastName ParticName, n.locktype, n.SupervisorAccessOnly, " 
	sSQL = sSQL & " n.SSN 'dbSSN' from phonecalls p inner join"
'	sSQL = sSQL & " userinformation u on p.whotookcall= u.acsuserid left join names n on"
	sSQL = sSQL & " userinformation u on convert(int,p.whotookcall)=u.ACSUserTrackNumber left join names n on"
	sSQL = sSQL & " p.ssn=n.ssn where"
	if Request.QueryString("AfterChange") <> "" then
		if ucase(Request.QueryString("SummCode")) <> "OVR" then
			sSQL = sSQL & " (p.CallSummaryCode <> 'OVR' OR p.CallSummaryCode IS NULL) AND" 
		end if
	else
		if ucase(Request.Querystring("SummCode")) <> "OVR" then
			sSQL = sSQL & " (p.CallSummaryCode <> 'OVR' OR p.CallSummaryCode IS NULL) AND" 
		end if
	end if
	sSQL = sSQL & " (n.DepNumber=0 or n.DepNumber IS NULL) AND " 

	if Request.QueryString("AfterChange") <> "" then
		if Trim(Request.QueryString("ActionType")) = "0" then
			iPageNo = 1
		else
			if Trim(Request.QueryString("ActionType")) = "" then
				iPageNo = 1
			else
				iPageNo = Cint(Request.QueryString("ActionType"))
				if iPageNo < 1 Then 
					iPageNo = 1
				end if
			end if
		end if
		if Request.QueryString("SSN") <> "" then
			m_iSSN = Request.QueryString("SSN")
			sSQL = sSQL & " p.SSN =" & m_iSSN & " AND"
		end if
		if Request.QueryString("CallBack") <> "" then
			m_sCallBack = Request.QueryString("CallBack")
			sSQL = sSQL & " p.CallBack ='" & m_sCallBack & "' AND"
		end if
		if Request.QueryString("LogonID") <> "" then
			m_sLogonID = Request.QueryString("LogonID")
			sSQL = sSQL & " u.LogonID='" & m_sLogonID & "' AND"
		end if
		if Request.QueryString("SummCode") <> "" then
			m_sSummCode = Request.QueryString("SummCode")
			sSQL = sSQL & " p.CallSummaryCode='" & m_sSummCode & "' AND"
		end if
		if Request.QueryString("ACSUserID") <> "" then
			m_sACSUserID = Request.QueryString("ACSUserID")
			sSQL = sSQL & " u.ACSUserID='" & m_sACSUserID & "' AND"
		end if
		if Request.QueryString("FromDate") <> "" then
			m_sFromDate = Request.QueryString("FromDate")
		end if
		if Request.QueryString("FromTime") <> "" then
			m_sFromTime = Request.QueryString("FromTime")
		end if
		if Request.QueryString("ThruDate") <> "" then
			m_sThruDate = Request.QueryString("ThruDate")
		end if
		if Request.QueryString("ThruTime") <> "" then
			m_sThruTime = Request.QueryString("ThruTime")
		end if
	else
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
		if Request.Querystring("SSN") <> "" then
			m_iSSN = Request.Querystring("SSN")
			sSQL = sSQL & " p.SSN =" & m_iSSN & " AND"
		end if
		if Request.Querystring("CallBack") <> "" then
			m_sCallBack = Request.Querystring("CallBack")
			sSQL = sSQL & " p.CallBack ='" & m_sCallBack & "' AND"
		end if
		if Request.Querystring("LogonID") <> "" then
			m_sLogonID = Request.Querystring("LogonID")
			sSQL = sSQL & " u.LogonID='" & m_sLogonID & "' AND"
		end if
		if Request.Querystring("SummCode") <> "" then
			m_sSummCode = Request.Querystring("SummCode")
			sSQL = sSQL & " p.CallSummaryCode='" & m_sSummCode & "' AND"
		end if
		if Request.Querystring("ACSUserID") <> "" then
			m_sACSUserID = Request.Querystring("ACSUserID")
			sSQL = sSQL & " u.ACSUserID='" & m_sACSUserID & "' AND"
		end if
		if Request.Querystring("FromDate") <> "" then
			m_sFromDate = Request.Querystring("FromDate")
		end if
		if Request.Querystring("FromTime") <> "" then
			m_sFromTime = Request.Querystring("FromTime")
		end if
		if Request.Querystring("ThruDate") <> "" then
			m_sThruDate = Request.Querystring("ThruDate")
		end if
		if Request.Querystring("ThruTime") <> "" then
			m_sThruTime = Request.Querystring("ThruTime")
		end if
	end if
	
	if m_sFromDate <> "" or m_sFromTime <> "" then
		sSQL = sSQL & " p.TimeOfCall >= '" & m_sFromDate & " " & m_sFromTime & "' AND"
	end if
	if m_sThruDate <> "" or m_sThruTime <> "" then
		if m_sThruTime = "" then
			sSQL = sSQL & " p.TimeOfCall <= '" & dateadd("s",-1,dateadd("d",1,m_sThruDate)) & "' AND"
		else
			sSQL = sSQL & " p.TimeOfCall <= '" & m_sThruDate & " " & m_sThruTime & "' AND"
		end if
	end if
	sSQL = left(sSQL,len(sSQL)-4)
	sSQL = sSQL & " order by p.TimeOfCall DESC"
end if
%>
<html>
<head>
<title>Details</title>
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
			<font SIZE="+2" face="verdana, arial, helvetica"><strong>Phone Call Search</strong></font>
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
				Response.write ("<p><ul><font color='red' face='verdana, arial, helvetica' size='2'><b>No matches were found -- try a different search.</b></font></ul>")
				exit sub
			end if
			adoRS.PageSize = m_iPageSize2 ' Number of rows per page
			iPageCount = adoRS.PageCount
			adoRS.AbsolutePage = iPageNo
		end if
%>
		<table COLS="3" WIDTH="100%" CELLPADDING="0" CELLSPACING="0" BORDER="0">
			<tr>
				<td WIDTH="10%"><input TYPE="button" SRC="images/query_button.gif" ACCESSKEY="S" NAME="NewQuery" VALUE="Start Query" onClick="ValidSend(0)">&nbsp;<input type="button" name="Phone" value="Log Call" onClick="LogCall()"></td>
				<td WIDTH="65%"><font face="verdana, arial, helvetica" size="2"><b>&nbsp;Enter your criteria, and then click the Start Query button or key (ALT+S).</b></font></td>
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
		<div STYLE="height=13%; width:100%; overflow=auto">
			<form ID="Criteria" METHOD="GET" ACTION="PhoneSearch.asp" TARGET="Details">
				<table CLASS="CriteriaTable" COLS="8" CELLPADDING="0" CELLSPACING="0">
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
					</tr>
					<tr>
						<td class="White">SSN:</td>
						<td><input SIZE="15" TYPE="TEXT" ID="SSN" NAME="SSN" VALUE="<%= m_iSSN%>"></td>
						<td class="White">Call Back:</td>
						<td><input SIZE="2" TYPE="TEXT" ID="CallBack" NAME="CallBack" VALUE="<%= m_sCallBack%>"></td>
						<td ALIGN="RIGHT" class="White">From Date&nbsp;&nbsp;&nbsp;</td>
						<td ALIGN="LEFT" class="White">From Time</td>
						<td ALIGN="RIGHT" class="White">Thru Date&nbsp;&nbsp;&nbsp;</td>
						<td ALIGN="LEFT" class="White">Thru Time</td>
					</tr>
					<tr>
						<td class="White">Logon ID:</td>
						<td><input TYPE="TEXT" ID="LogonID" NAME="LogonID" VALUE="<%= m_sLogonID%>"></td>
						<td class="White">Call Summary Code:</td>
						<td><input SIZE="4" TYPE="TEXT" ID="SummCode" NAME="SummCode" VALUE="<%= m_sSummCode%>"></td>
						<td ALIGN="RIGHT"><input SIZE="10" TYPE="TEXT" ID="FromDate" NAME="FromDate" VALUE="<%= m_sFromDate%>"></td>
						<td ALIGN="LEFT"><input SIZE="8" TYPE="TEXT" ID="FromTime" NAME="FromTime" VALUE="<%= m_sFromTime%>"></td>
						<td ALIGN="RIGHT"><input SIZE="10" TYPE="TEXT" ID="ThruDate" NAME="ThruDate" VALUE="<%= m_sThruDate%>"></td>
						<td ALIGN="LEFT"><input SIZE="8" TYPE="TEXT" ID="ThruTime" NAME="ThruTime" VALUE="<%= m_sThruTime%>"></td>
					</tr>
					<tr>
						<td class="White">ACS User ID:</td>
						<td><input SIZE="6" TYPE="TEXT" ID="ACSUserID" NAME="ACSUserID" VALUE="<%= m_sACSUserID%>"></td>
					</tr>
				</table>
				<input TYPE="HIDDEN" ID="ActionType" NAME="ActionType" VALUE>
			</form>
		</div>

		<br CLEAR="LEFT">
<%
		if bOpenRS then
    '**********************************************************************************************************************************************
    '** Removed overflow=auto in the div style tag below in order to see the complete recordset without scrolling also to print the whole recordset.
    '**********************************************************************************************************************************************
%> 
      
      <div STYLE="height=68%; width=100%;">
      <p><font size=2 face='verdana, arial, helvetica'><b>As of <%= NOW %>, this Recordset currently has <font color="<%= Session("EmphColor")%>"><%= adoRS.RecordCount %></font> phone records in it.</b></font>&nbsp;&nbsp;<input type="button" name="Print" id="Print" value="Print Records" onClick="window.print()">
      <% if request.querystring("LogType") = "Claim" then %><p><font face="verdana, arial, helvetica" size="2"><b><%= ClaimNum %></b></font><% end if %>
<%
      iRowCount = adoRS.PageSize
      iRemainder = iRowCount - adoRS.RecordCount
			i = 0
			Do While Not adoRS.EOF and iRowCount > 0
      
      if bUseSQL and iRowCount > 1 then
        sTemp = "<table width='100%' border='0'><tr><td align='right'><font size=1 face='verdana, arial, helvetica'><b>This is phone record <font color='sienna'>" & iRowCount - iRemainder & "</font> out of a possible " & adoRS.RecordCount & ".</b></font></td></tr></table>"
        Response.Write sTemp
      end if

%> 
        
        <table WIDTH="100%" FRAME="BOX" RULES="NONE" BORDER="1" bordercolor="black" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
					<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
						<td WIDTH="20%" bgcolor="#cccccc">Date & Time:</td>
						<td STYLE='color:<%= Session("EmphColor")%>;' WIDTH="30%" bgcolor="#cccccc"><%= adoRS("TimeOfCall")%><input TYPE="HIDDEN" ID="CallTime<%= i%>" NAME="CallTime<%= i%>" VALUE="<%= adoRS("FullTime")%>"></td>
						<td WIDTH="20%" bgcolor="#cccccc">Call Taker:</td>
						<td STYLE='color:<%= Session("EmphColor")%>;' WIDTH="30%" bgcolor="#cccccc"><%= adoRS("UserName")%></td>
					</tr>
					<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
						<td bgcolor="#F0F0F0">SSN:</td>
						<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("SSN")%></td>
						<td bgcolor="#F0F0F0">Participant Name:</td>
<%
						if adoRS("locktype") <> "" then
              if Manager = "True" or Supervisor = "True" or Auditor = "True" then
                sTemp = "<TD ALIGN=LEFT bgcolor='#F0F0F0'><A HREF='PersonDetails.asp?SSN=" & adoRS("SSN") & "&DepNum=0&locktype=" & adoRS("locktype") & "&Admin=" & adoRS("SupervisorAccessOnly")
                sTemp = sTemp & "' onClick='javascript:return LogCheck(" & adoRS("SSN") & ")'>"
                sTemp = sTemp & adoRS("ParticName") & "</A></TD>"
              else
                sTemp = "<TD ALIGN=LEFT bgcolor='#F0F0F0'><A HREF='PersonLockDetail.asp?SSN=" & adoRS("SSN") & "&locktype=" & adoRS("locktype") & "&DepNum=0&Admin=" & adoRS("SupervisorAccessOnly")
                sTemp = sTemp & "' onClick='WorkingStatus()'>"
                sTemp = sTemp & adoRS("ParticName") & "</A></TD>"
              end if
            else
              sTemp = "<TD ALIGN=LEFT bgcolor='#F0F0F0'><A HREF='PersonDetails.asp?SSN=" & adoRS("SSN") & "&DepNum=0" 
							sTemp = sTemp & "' onClick='javascript:return LogCheck(" & adoRS("SSN") & ")'>"
							sTemp = sTemp & adoRS("ParticName") & "</A></TD>"
            end if
            response.write sTemp  
            
            'if not isnull(adoRS("dbSSN")) then
							'sTemp = "<TD ALIGN=LEFT bgcolor='#F0F0F0'><A HREF='PersonDetails.asp?SSN=" & adoRS("SSN") & "&DepNo=0" 
							'sTemp = sTemp & "' onClick='javascript:return LogCheck(" & adoRS("SSN") & ")'>"
							'sTemp = sTemp & adoRS("ParticName") & "</A></TD>"
						'else
							'sTemp = "<TD bgcolor='#F0F0F0'></TD>"
						'end if
						'Response.Write sTemp
%>					    
					</tr>
					<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
						<td COLSPAN="4">
<%
            
              if adoRS("TimeOfCall") > date then
%>
								<table COLS="7" ALIGN="CENTER" BORDER="0" RULES="NONE" WIDTH="100%" bgcolor="#F0F0F0" style="font:9pt verdana, arial, helvetica,sans-serif">
									<tr>
										<td>Caller:</td>
										<td><input SIZE="40" TYPE="TEXT" ID="Caller<%= i%>" NAME="Caller<%= i%>" VALUE="<%= adoRS("NameOfCaller")%>"></td>
										<td>Phone Number:<input SIZE="10" TYPE="TEXT" ID="PhoneNumber<%= i%>" NAME="PhoneNumber<%= i%>" VALUE="<%= adoRS("PhoneNumber")%>">
											  Ext:<input SIZE="4" TYPE="TEXT" ID="Ext<%= i%>" NAME="Ext<%= i%>" VALUE="<%= adoRS("Extension")%>">
										</td>
										<!--- <td>Call Summ. Code:</td> --->
										<td><input SIZE="4" TYPE="HIDDEN" ID="SummCode<%= i%>" NAME="SummCode<%= i%>" VALUE="<%= adoRS("CallSummaryCode")%>"></td>
										<td>Call Back:</td>
										<td><input SIZE="2" TYPE="TEXT" ID="CallBack<%= i%>" NAME="CallBack<%= i%>" VALUE="<%= adoRS("CallBack")%>"></td>
									</tr>
									<tr>
										<td>Details:</td>
										<td COLSPAN="4"><textarea WRAP="VIRTUAL" COLS="90" ROWS="3" ID="NOTETEXT<%= i%>" NAME="NOTETEXT<%= i%>"><%= adoRS("NoteText")%></textarea></td>
										<td COLSPAN="2" ALIGN="CENTER"><% if LCASE(MID(adoRS("UserName"), INSTRREV(adoRS("UserName"), " ") +1)) = LCASE(MID(Session("User"), 2)) then %><input TYPE="BUTTON" ID="SaveCall" VALUE="Save Changes" onClick="UpdateCall <%= i%>,0,&quot;<%= adoRS("SSN")%>&quot;,&quot;<%= adoRS("TimeofCall")%>&quot;"><% else %><%= Session("NWAName") %> is not authorized to make changes to this phone record.<% end if %></td>
									</tr>
								</table>
<%
							else
%>
								<table COLS="8" ALIGN="CENTER" BORDER="1" RULES="NONE" WIDTH="100%" style="font:9pt verdana, arial, helvetica,sans-serif">
									<tr>
										<td>Caller:</td>
										<td STYLE='color:<%= Session("EmphColor")%>;'><%= adoRS("NameOfCaller")%></td>
										<td>Phone Number:</td>
<%
											sTemp=adoRS("PhoneNumber")
											if trim(adoRS("Extension") & " ") <> "" then
												sTemp = sTemp & " Ext:" & adoRS("Extension")
											end if
											Response.Write ("<td STYLE='color:sienna;'>" & sTemp & "</td>")
%>
										
										<!--- <td>Call Summ. Code:</td>
										<td STYLE='color:<%= Session("EmphColor")%>;'><%= adoRS("CallSummaryCode")%></td> --->
										<td>Call Back:</td>
										<td STYLE='color:<%= Session("EmphColor")%>;'><%= adoRS("CallBack")%></td>
										<td align="right"><input TYPE="BUTTON" ID="CancelCallBack<%= i%>" CBVALUE="<%= adoRS("CallBack")%>" VALUE="Cancel Call Back" onClick="UpdateCall <%= i%>,2,&quot;<%= adoRS("SSN")%>&quot;,&quot;<%= adoRS("TimeofCall")%>&quot;"><% if m_bIsSupervisor then %>&nbsp;&nbsp;<input TYPE="BUTTON" ID="DelCall" VALUE="Delete Call" onClick="UpdateCall <%= i%>,1,&quot;<%= adoRS("SSN")%>&quot;,&quot;<%= adoRS("TimeofCall")%>&quot;"><% end if %></td>
									</tr>
									<tr>
										<td COLSPAN="1">Details:</td>
                    <td colspan="7" STYLE='color:<%= Session("EmphColor")%>;'><%= adoRS("NoteText")%></td>
									</tr>
								</table>
<%
							end if
%>
						</td>
					</tr>
        </table>
				<br>
<%
				iRowCount = iRowCount - 1
				i = i + 1
				adoRS.MoveNext
			Loop
%>
      </div>
<%
			adoRS.Close
			adoConn.Close
			set adoRS = nothing
			set adoConn = nothing
		end if
	end sub
%>

</body>
<!--#include file="VBFuncs.inc" -->
<script LANGUAGE="VBSCRIPT">
	
	sub ValidSend(iIndex)
	dim bValidData

		Criteria.SSN.Value = trim(Criteria.SSN.Value)
		Criteria.CallBack.Value = trim(Criteria.CallBack.Value)
		Criteria.LogonID.Value = trim(Criteria.LogonID.Value)
		Criteria.SummCode.Value = trim(Criteria.SummCode.Value)
		Criteria.ACSUserID.Value = trim(Criteria.ACSUserID.Value)
		Criteria.FromDate.Value = trim(Criteria.FromDate.Value)
		Criteria.FromTime.Value = trim(Criteria.FromTime.Value)
		Criteria.ThruDate.Value = trim(Criteria.ThruDate.Value)
		Criteria.ThruTime.Value = trim(Criteria.ThruTime.Value)
		
		bValidData = false
		if Criteria.SSN.Value <> "" then
			if not isNumeric(Criteria.SSN.Value) then
				msgbox "Please enter a valid SSN.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.CallBack.Value <> "" then
			if ContainsInvalids(Criteria.CallBack.Value) then
				msgbox "Please remove invalid characters from the Call Back field.",,"Invalid Data"
				exit sub
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
		if Criteria.SummCode.Value <> "" then
			if ContainsInvalids(Criteria.SummCode.Value) then
				msgbox "Please remove invalid characters from the Call Summary Code field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.ACSUserID.Value <> "" then
			if ContainsInvalids(Criteria.ACSUserID.Value) then
				msgbox "Please remove invalid characters from the ACS User field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.FromDate.Value <> "" then
			if not isDate(Criteria.FromDate.Value) then
				msgbox "Please enter a valid date for the From Date field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.FromTime.Value <> "" then
			if not isDate(Criteria.FromTime.Value) then
				msgbox "Please enter a valid time for the From Time field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.ThruDate.Value <> "" then
			if not isDate(Criteria.ThruDate.Value) then
				msgbox "Please enter a valid date for the Thru Date field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.ThruTime.Value <> "" then
			if not isDate(Criteria.ThruTime.Value) then
				msgbox "Please enter a valid time for the Thru Time field.",,"Invalid Data"
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
			"<%= m_sCallBack%>" <> Criteria.CallBack.Value or _
			"<%= m_sLogonID%>" <> Criteria.LogonID.Value or _
			"<%= m_sSummCode%>" <> Criteria.SummCode.Value or _
			"<%= m_sACSUserID%>" <> Criteria.ACSUserID.Value or _
			"<%= m_sFromDate%>" <> Criteria.FromDate.Value or _
			"<%= m_sThruDate%>" <> Criteria.ThruDate.Value or _
			"<%= m_sFromTime%>" <> Criteria.FromTime.Value or _
			"<%= m_sThruTime%>" <> Criteria.ThruTime.Value then
			Criteria.ActionType.value=""
		else
			Criteria.ActionType.value=iIndex
		end if
		Criteria.submit
	end sub
	
	sub UpdateCall(iIndex, iType, sSSN, sDate)
'
'	iType = 0: Update
'	iType = 1: Delete
'	iType = 0: Cancel Call Back
'
	dim sHREF, sTemp
	dim vbOKCancel, vbOk

		vbOkCancel = 1
		vbOk = 1

		select case iType
			case 0
				'if 'msgbox("Are you sure you want to save your changes for this phone call?",vbOkCancel,"Save Changes?") <> vbOk then
					'exit sub
				'end if

				if trim(document.all.item("caller" & iIndex).value) <> "" then
					if ContainsInvalids(document.all.item("caller" & iIndex).value) then
						msgbox "Please remove invalid characters from the Caller field.",,"Invalid Data"
						exit sub
					end if
				end if
				if trim(document.all.item("phonenumber" & iIndex).value) <> "" then
					if ContainsInvalids(document.all.item("phonenumber" & iIndex).value) then
						msgbox "Please remove invalid characters from the Phone Number field.",,"Invalid Data"
						exit sub
					end if
				end if
				if trim(document.all.item("ext" & iIndex).value) <> "" then
					if ContainsInvalids(document.all.item("ext" & iIndex).value) then
						msgbox "Please remove invalid characters from the Ext. field.",,"Invalid Data"
						exit sub
					end if
				end if
				if trim(document.all.item("callback" & iIndex).value) <> "" then
					if ContainsInvalids(document.all.item("callback" & iIndex).value) then
						msgbox "Please remove invalid characters from the Call Back field.",,"Invalid Data"
						exit sub
					end if
				end if
				if trim(document.all.item("summcode" & iIndex).value) <> "" then
					if ContainsInvalids(document.all.item("summcode" & iIndex).value) then
						msgbox "Please remove invalid characters from the Call Summary Code field.",,"Invalid Data"
						exit sub
					end if
				end if
				if trim(document.all.item("notetext" & iIndex).value) <> "" then
					if ContainsInvalids(document.all.item("notetext" & iIndex).value) then
						if msgbox("There are reserved characters in the Notes field.  These characters will be removed before the data is written to the database (except for single quotes).",vbOkCancel,"Notes") <> vbOk then
							exit sub
						end if
					end if
				end if
				
			case 1
				if msgbox("Are you sure you want to delete this phone call?",vbOkCancel,"Delete Call?") <> vbOk then
					exit sub
				end if
				
			case 2
				if lcase(document.all.item("CancelCallBack" & iIndex).CBVALUE) <> "y" then
					if msgbox("This phone call does not appear to have the Call Back field set to Yes.  Set it to No anyway?",vbOkCancel,"Cancel Call Back?") <> vbOk then
						exit sub
					end if
				else					
					if msgbox("Are you sure you want to set the Call Back for this phone call to No?",vbOkCancel,"Cancel Call Back?") <> vbOk then
						exit sub
					end if
				end if
				
		end select

		sHREF = "UpdatePhoneLog.asp?SSN=" & sSSN & "&CallTime=" & document.all.item("calltime" & iIndex).value & "&"
		sHREF = sHREF & "CriteriaSSN=<%= m_iSSN%>&CallBack=<%= m_sCallBack%>&SummCode=<%= m_sSummCode%>"
		sHREF = sHREF & "&LogonID=<%= m_sLogonID%>&ACSID=<%= m_sACSUserID%>&FromDate=<%= m_sFromDate%>&FromTime="
		sHREF = sHREF & "<%= m_sFromTime%>&ThruDate=<%= m_sThruDate%>&ThruTime=<%= m_sThruTime%>"

		select case iType
			case 0
				sHREF = sHREF & "&Caller=" & document.all.item("caller" & iIndex).value
				sHREF = sHREF & "&Phone=" & document.all.item("phonenumber" & iIndex).value
				sHREF = sHREF & "&Ext=" & document.all.item("ext" & iIndex).value
				sHREF = sHREF & "&NewCallBack=" & document.all.item("CallBack" & iIndex).value
				sHREF = sHREF & "&NewSummCode=" & document.all.item("SummCode" & iIndex).value
				sTemp = replace(document.all.item("notetext" & iIndex).value,chr(34),"")
				sTemp = replace(sTemp,"&","")
				sTemp = FixQuote(sTemp)
				sHREF = sHREF & "&Notes=" & sTemp
			case 1
				sHREF = sHREF & "&Delete=true"
			case 2
				sHREF = sHREF & "&CancelCallback=true"
		end select

		WorkingStatus
		top.Details.location.replace(sHREF)
	end sub
	
</script>
</html>
