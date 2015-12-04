<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
dim i, j, adoConn, adoRS, adoCmd, adoParam, sSQL, sTemp, sShownSSN
dim m_iSSN, m_iDepNo, m_bDependent, m_sParticName, m_bUseEAS, m_sBirthDate
dim avFundPlan(), avEligData(), avSchools(), avSchoolData(), iNumSchools, dep_FName, dep_LName
dim iNumFundPlans, iOldestYear, iNewestYear, iNumYears, iNewestMonth, iOldestMonth
dim iCurrRow, iCurrCol, sSectionTitle, avRowData(), iCurrYear
dim avMonthName(12)
dim sAttend, iStartCol, iEndCol, sCurrSchool, locktype, DepNum, Admin, LastName, SSN, Reason
dim M, D, T, C, P, V
dim ParticAdmin, ParticLocktype, sTemp2

m_iDepNo = request.querystring("DepNo")
m_iSSN = request.querystring("SSN")
m_iDepNo = request.querystring("DepNum")
locktype = request.querystring("locktype")
Admin = request.querystring("Admin")
LastName = request.querystring("LastName")

set adoConn = Server.CreateObject("ADODB.Connection")
adoConn.CommandTimeout = 0
set adoRS = Server.CreateObject("ADODB.Recordset")
adoConn.Open Application("DataConn")
sSQL = "select locktype, SupervisorAccessOnly from Names where SSN=" & m_iSSN & " and DepNumber=" & m_iDepNo
adoRS.Open sSQL, adoConn, adOpenKeyset, adLockOptimistic
if not adoRS.EOF then
  if adoRS("locktype") <> "" then
  else
    response.redirect "PersonDetails.asp?SSN=" & m_iSSN & "&DepNum=0&locktype=" & adoRS("locktype") & "&Admin=" & adoRS("SupervisorAccessOnly")
  end if
end if
  'if locktype <> "" then
    'locktype = true
  'else
    'locktype = false
  'end if
'end if
        'response.redirect "PersonDetails.asp?SSN=" & m_iSSN & "&DepNum=0&locktype=" & adoRS("locktype") & "&Admin=" & adoRS("SupervisorAccessOnly")
        'end if

adoRS.Close
adoConn.Close
set adoRS = nothing
set adoConn = nothing
'response.write "SSN = " & m_iSSN & "<br>"
'response.write "Last Name = " & LastName & "<br>"
'response.write "Dependent Number = " & m_iDepNo & "<br>"
'response.write "LockType = " & locktype & "<br>"
'response.write "Admin = " & Admin
%>

		<script LANGUAGE="VBSCRIPT">			
			sub PickButton(iIndex)
				select case iIndex
					case 0
						top.AppStatus.location.replace("Working.htm")
						ClaimPost.SSN.value = <%= m_iSSN%>
						ClaimPost.DepNum.value = <%= m_iDepNo%>
						ClaimPost.submit
					case 1 
						top.AppStatus.location.replace("Working.htm")
						FormLetterPost.SSN.value = <%= m_iSSN%>
						FormLetterPost.submit
					case 2 
						top.AppStatus.location.replace("Working.htm")
						LetterPost.SSN.value = <%= m_iSSN%>
						LetterPost.submit
					case 3 
						top.AppStatus.location.replace("Working.htm")
						CheckPost.SSN.value = <%= m_iSSN%>
						CheckPost.DepNum.value = <%= m_iDepNo%>
						CheckPost.submit
					case 4 
						top.AppStatus.location.replace("Working.htm")
						PhoneCallPost.Crit.value = <%= m_iSSN%>
						PhoneCallPost.submit
          case 5 
						top.AppStatus.location.replace("Working.htm")
						PrescriptionPost.SSN.value = <%= m_iSSN%>
            PrescriptionPost.DepNum.value = <%= m_iDepNo %>
						PrescriptionPost.submit
          case 6 
						top.AppStatus.location.replace("Working.htm")
						TrackingPost.SSN.value = <%= m_iSSN%>
            TrackingPost.DepNum.value = <%= m_iDepNo %>
            TrackingPost.Years.value = <%= Year(Date) %>
						TrackingPost.submit
          case 7 
						top.AppStatus.location.replace("Working.htm")
						VisionClaimPost.SSN.value = <%= m_iSSN%>
            VisionClaimPost.DepNum.value = <%= m_iDepNo %>
            VisionClaimPost.HWType.value = "V"
						VisionClaimPost.submit
          case 8
						top.AppStatus.location.replace("Working.htm")
						BlankClaimPost.UserCriteria.value = "Initial"
						BlankClaimPost.submit
					case else
						msgbox "none"
				end select
			end sub
			
			function ShowPlanDesc(Fund,Plan,Rate)
			dim sResult

				sResult=showModalDialog("ACSInfo.asp?EASFund=" & Fund & "&EASPlan=" & Plan & "&EASRate=" & Rate, "ACS Information", "dialogWidth:500px; dialogHeight:295px; help:no;")
				
			end function

		</script>
<%  
%>
<!--#include file="VBFuncs.inc" -->
</head>
<body LANGUAGE="VBScript" onload="UpdateScreen(0)">
<link REL="STYLESHEET" HREF="styles/CritTable.css">
<%
if Admin then
  set adoConn = Server.CreateObject("ADODB.Connection")
  adoConn.CommandTimeout = 0
  set adoRS = Server.CreateObject("ADODB.Recordset")
  adoConn.Open Application("DataConn")
  if m_iSSN <> "" then
    sSQL = "select LockedReason from Names where SSN=" & m_iSSN 
  else
    sSQL = "select LockedReason from Names where LastName='" & LastName & "'"
  end if 
  'response.write "<br>" & sSQL
  adoRS.Open sSQL, adoConn, adOpenKeyset, adLockOptimistic
  'response.write sSQL
  Reason = adoRS("LockedReason")
%>

  <font size="2" face="verdana, arial, helvetica">
  <p>&nbsp;
  <p>&nbsp;
  <center><p><font size="4" face="arial, helvetica" color="red"><b>Access Denied !!</b></font>
  <p><b>Administrative Lock On PDF. Contact Your Supervisor.</b>
  <p><b>To return to a new search please click the Participants/Dependent Tab.</b>
  <p>This file is locked because:
  <p><b><% if Reason <> "" then %><font color="#983300"><%= Reason %><% else %><font color="#983300">No Reason in Database !<% end if %></b></font>
  <p>The Lock Types are: <b><% if locktype <> "" then %><font color="#983300"><%= locktype %><% else %><font color="#983300">No Lock Types in Database !<% end if %></b></font>
  <p>This participants SSN is <b><font color="#983300"><%= m_iSSN %></b></font>.
  <!--- <p><img SRC="images/log.gif" onClick="LogCall()"> --->
  
  
  </center>
  </font>
  
<%  
  adoRS.Close
  adoConn.Close
  set adoRS = nothing
  set adoConn = nothing
else

if locktype <> "" then

  if m_iSSN = "" or m_iDepNo = "" then
	  Response.Write ("<p><ul><font color='red' face='verdana, arial, helvetica' size='2'><b>SSN or DepNumber missing; contact your network administrator.</b></font></ul>")
  else
	
  set adoConn = Server.CreateObject("ADODB.Connection")
	adoConn.CommandTimeout = 0
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
  
  sSQL = "select locktype from Names where SSN =" & m_iSSN & " AND DepNumber > 0 and locktype is null"
  adoRS.Open sSql, adoConn, adOpenForwardOnly, adLockOptimistic
  
  if adoRS.EOF then
    sTemp2 = " the <font color='red'>Full Family</font> of "
  end if
  
  adoRS.Close
	
	if m_iDepNo = 0 then
		m_bDependent = false

			sSQL = "select n.FirstName + ' ' + n.MiddleName + ' ' + n.LastName 'Name'," 
		sSQL = sSQL & " n.ssn, n.LockedReason, n.BirthDate, n.Gender, o.Fund, "
		if not m_bUseEAS then
			sSQL = sSQL & " h.StreetAddress Address, "
		else
			sSQL = sSQL & " h.Addr1 Address, "
		end if
		sSQL = sSQL & " h.City, h.State, h.PostalCode,"
		sSQL = sSQL & " i.PhoneNumber HomePhone, i.WorkPhone, m.Description 'MaritalStatus',"
		if not m_bUseEAS then
			sSQL = sSQL & " e.RetireDate, e.DeathDate, n.locktype, n.LockedBy, n.LockedDate, n.SupervisorAccessOnly FROM Names n LEFT JOIN Homes"
		else
			sSQL = sSQL & " n.RetireDate, n.DeathDate FROM EASNames n LEFT JOIN EASHomes"
		end if
		sSQL = sSQL & " h on (h.SSN = n.SSN) LEFT JOIN Insureds i on (i.SSN = n.SSN) LEFT JOIN"
		sSQL = sSQL & " MaritalStatuses m on (i.MaritalStatus = m.MaritalStatus) LEFT JOIN OtherAreaNames o on (n.SSN=o.SSN) "
		if not m_bUseEAS then
			sSQL = sSQL & " LEFT JOIN EASNames e on (n.SSN = e.SSN)"
		end if
		sSQL = sSQL & " WHERE n.SSN = " & m_iSSN & " AND n.DepNumber = " & m_iDepNo
		if not m_bUseEAS then
			sSQL = sSQL & " AND h.datecreated = (select max(DateCreated) from Homes where ssn=" & m_iSSN & ")"
		end if
	else
		m_bDependent = true
		sSQL = "select n.BirthDate, n.locktype, n.SupervisorAccessOnly, n.FirstName + ' ' + n.MiddleName + ' ' + n.LastName 'Name'" 
		sSQL = sSQL & " FROM" 
		if not m_bUseEAS then
			sSQL = sSQL & " Names"
		else
			sSQL = sSQL & " EASNames"
		end if
		sSQL = sSQL & " n WHERE n.SSN = " & m_iSSN & " AND n.DepNumber = 0"  
		adoRS.Open sSql, adoConn, adOpenForwardOnly, adLockOptimistic
		ParticAdmin = adoRS("SupervisorAccessOnly")
    ParticLocktype = adoRS("locktype")
    m_sParticName = adoRS("Name")
    m_sBirthDate = adoRS("BirthDate")
		adoRS.Close
			
		sSQL = "select n.FirstName + ' ' + n.MiddleName + ' ' + n.LastName 'Name'," 
		sSQL = sSQL & " n.ssn, n.DepSSN, n.locktype, n.LockedReason, n.LockedBy, n.LockedDate, n.BirthDate, n.Gender, n.RelationCode,"
		if not m_bUseEAS then
			sSQL = sSQL & "h.StreetAddress Address,"
		else
			sSQL = sSQL & "h.Addr1 Address,"
		end if
		sSQL = sSQL & " h.City, h.State, h.PostalCode, r.Description Relationship"
		if not m_bUseEAS then
			sSQL = sSQL & " FROM Names n LEFT JOIN Homes"
		else
			sSQL = sSQL & " FROM EASNames n LEFT JOIN EASHomes"
		end if
		sSQL = sSQL & " h ON (h.SSN = n.SSN)"
		sSQL = sSQL & " LEFT JOIN RelationDescription r ON (n.RelationCode = r.RelationCode)"
		sSQL = sSQL & " WHERE n.SSN = " & m_iSSN & " AND n.DepNumber = " & m_iDepNo
		if not m_bUseEAS then
			sSQL = sSQL & " AND h.datecreated = (select max(DateCreated) from Homes where ssn=" & m_iSSN & ")"
		end if
	end if
  end if	
	if m_bDependent then
		sTemp = "Dependent Information"
	else
		sTemp = "Participant Information"
	end if
%>
	<table WIDTH="100%" COLS="3" border="0">
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
				<font SIZE="+2" face="verdana, arial, helvetica" color="red"><strong><%= sTemp%></strong></font>
			</td>
			<td ALIGN="CENTER" WIDTH="20%">
				<img SRC="images/bluebar2.gif" onClick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0">
			</td>
		</tr>
	</table>
<%
	adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
	if adoRS.EOF then
		Response.Write ("<p><ul><font color='red' face='verdana, arial, helvetica' size='2'><b>Unable to find all necessary information; please contact your network administrator.</b></font></ul>")
	else
    
		if m_bDependent then
			sShownSSN = adoRS("DepSSN")
		else
			sShownSSN = adoRS("SSN")
			m_sParticName = adoRS("Name")
      m_sBirthDate = adoRS("BirthDate")
		end if
    if m_bDependent then
		  sTemp = "Dependent"
	  else
		  sTemp = "Participant"
	  end if
%>
    <table width="70%" ALIGN="CENTER" BORDER="0">
			<tr>
				<td width="5%">&nbsp;</td>
        <td ALIGN="CENTER">
					<!--- <input TYPE="BUTTON" VALUE="Claims" onClick="PickButton(0)"> --->
					<input TYPE="BUTTON" VALUE="ACELetters" onClick="PickButton(1)">
					<input TYPE="BUTTON" VALUE="ACSLetters" onClick="PickButton(2)">
					<input TYPE="BUTTON" VALUE="Checks" onClick="PickButton(3)">
					<input TYPE="BUTTON" VALUE="Phone Calls" onClick="PickButton(4)">
          <input TYPE="BUTTON" VALUE="Prescription" onClick="PickButton(5)">
          <input TYPE="BUTTON" VALUE="Tracking" onClick="PickButton(6)">
          <!--- <input TYPE="BUTTON" VALUE="VisionClaims" onClick="PickButton(7)">
          <input TYPE="BUTTON" VALUE="BlankClaims" onClick="PickButton(8)"> --->
				</td>
        <td width="20%">&nbsp;</td>
			</tr>
      <tr>
				<td width="5%">&nbsp;</td>
        <td ALIGN="CENTER">
					<input TYPE="BUTTON" VALUE="AllClaims" onClick="PickButton(0)">
          <input TYPE="BUTTON" VALUE="VisionClaims" onClick="PickButton(7)">
          <!--- <input TYPE="BUTTON" VALUE="BlankClaims" onClick="PickButton(8)"> --->
				</td>
        <td width="20%">&nbsp;</td>
			</tr>
		</table>
		<form ID="ClaimPost" TARGET="Details" ACTION="ClaimSearch.asp" METHOD="GET">
			<input TYPE="HIDDEN" ID="SSN" NAME="SSN" VALUE>
			<input TYPE="HIDDEN" ID="DepNum" NAME="DepNum" VALUE>
		</form>
		<form ID="FormLetterPost" TARGET="Details" ACTION="FormLetterSearch.asp" METHOD="GET">
			<input TYPE="HIDDEN" ID="SSN" NAME="SSN" VALUE>
		</form>
		<form ID="LetterPost" TARGET="Details" ACTION="LetterSearch.asp" METHOD="GET">
			<input TYPE="HIDDEN" ID="SSN" NAME="SSN" VALUE>
		</form>
		<form ID="CheckPost" TARGET="Details" ACTION="CheckSearch.asp" METHOD="GET">
			<input TYPE="HIDDEN" ID="SSN" NAME="SSN" VALUE>
			<input TYPE="HIDDEN" ID="DepNum" NAME="DepNum" VALUE>
		</form>
		<form ID="PhoneCallPost" TARGET="Details" ACTION="PhoneSearch.asp" METHOD="GET">
			<input TYPE="HIDDEN" ID="Crit" NAME="SSN" VALUE>
		</form>
    <form ID="PrescriptionPost" TARGET="Details" ACTION="Prescription.asp" METHOD="GET">
			<input TYPE="HIDDEN" ID="SSN" NAME="SSN" VALUE>
      <input TYPE="HIDDEN" ID="DepNum" NAME="DepNum" VALUE>
		</form>
    <form ID="TrackingPost" TARGET="Details" ACTION="Tracking.asp" METHOD="GET">
			<input TYPE="HIDDEN" ID="SSN" NAME="SSN" VALUE>
      <input TYPE="HIDDEN" ID="DepNum" NAME="DepNum" VALUE>
      <input TYPE="HIDDEN" ID="Years" NAME="Years" VALUE>
		</form>
    <form ID="VisionClaimPost" TARGET="Details" ACTION="ClaimSearch.asp" METHOD="GET">
			<input TYPE="HIDDEN" ID="SSN" NAME="SSN" VALUE>
			<input TYPE="HIDDEN" ID="DepNum" NAME="DepNum" VALUE>
      <input TYPE="HIDDEN" ID="HWType" NAME="HWType" VALUE>
		</form>
    <form ID="BlankClaimPost" TARGET="Details" ACTION="ClaimSearch.asp" METHOD="GET">
			<input TYPE="HIDDEN" ID="UserCriteria" NAME="UserCriteria" VALUE>
		</form>
		<br CLEAR="LEFT">
    <table width="70%" ALIGN="CENTER" BORDER="0">
      <tr>
        <td width="5%">&nbsp;</td>
        <td align="center"><font face="verdana, arial, helvetica" size="2"><img src="images/lock4.gif" width="32" height="32" border="0" alt="Locked File">&nbsp;<b>This is a locked file for <% if m_bDependent then %>dependent&nbsp;<% else %><%= sTemp2 %>&nbsp;<% end if %> <%= adoRS("Name") %>
        <br>Reason: <% if adoRS("LockedReason") <> "" then %><%= adoRS("LockedReason") %><% else %>NO REASON IN DATABASE!!<% end if %>
        <br>Locked by: <% if adoRS("LockedBy") <> "" then %><%= adoRS("LockedBy") %><% else %>NO NAME IN DATABASE!!<% end if %>
        <br>Date Locked: <% if adoRS("LockedDate") <> "" then %><%= formatdatetime(adoRS("LockedDate"),vbShortDate) %><% else %>NO DATE IN DATABASE!!<% end if %>
        </b></font></td>
        <td width="20%">&nbsp;</td>
      </tr>
		</table>

<% 
    if INSTR(locktype, "M") <> "0" then
      M = "Medical"
    end if
    if INSTR(locktype, "D") <> "0" then
      D = "Dental"
    end if
    if INSTR(locktype, "T") <> "0" then
      T = "Timeloss"
    end if
    if INSTR(locktype, "C") <> "0" then
      C = "CSB"
    end if
    if INSTR(locktype, "P") <> "0" then
      P = "Prescription"
    end if
    if INSTR(locktype, "V") <> "0" then
      V = "Vision"
    end if

%>
    <table border="1" width="100%" COLS=4 RULES=GROUPS bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<COLGROUP SPAN=2>
			<COLGROUP SPAN=2>
<%
			sTemp = "<tr BorderColorDark='#F0F0F0' BorderColorlight='#999999'><td width=25% bgcolor='#cccccc'><strong>"
			if m_bDependent then
				sTemp = sTemp & "Dependent Code:"
			else
				sTemp = sTemp & "SSN:"
			end if
			sTemp = sTemp & "</strong></TD><TD width=25% STYLE='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'>"
			if m_bDependent then
				sTemp = sTemp &"<b>" & m_iDepNo
			else
				sTemp = sTemp &"<b>" & sShownSSN
			end if
			sTemp = sTemp & "</b></TD><td width=25% bgcolor='#cccccc'><strong>Name:</strong>"
			sTemp = sTemp & "<TD width=25% STYLE='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'><b>" & adoRS("Name") & "</b></td></tr>"
			Response.Write sTemp			
			if m_bDependent then
				sTemp = "<TR BorderColorDark='#F0F0F0' BorderColorlight='#999999'><td bgcolor='#cccccc'><strong>Participant SSN:</strong></TD>"
				sTemp = sTemp & "<TD STYLE='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'><b>" & adoRS("SSN") & "</b></td>"
				sTemp = sTemp & "<td bgcolor='#cccccc'><strong>Participant Name:</strong></TD>"
				sTemp = sTemp & "<TD STYLE='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'>"
        if ParticLocktype <> "" then
          sTemp = sTemp & "<A HREF='PersonLockDetail.asp?SSN=" & m_iSSN & "&locktype=" & ParticLocktype & "&DepNum=0&Admin=" & ParticAdmin
          sTemp = sTemp & "' onClick='WorkingStatus()'><b>" & m_sParticName & "</b></a>"
				  sTemp = sTemp & "</TD></tr>"
        else
				  sTemp = sTemp & "<A HREF='PersonDetails.asp?SSN=" & m_iSSN & "&DepNum=0' onClick='WorkingStatus()'><b>" & m_sParticName & "</b></A>"
				  sTemp = sTemp & "</TD></tr>"
        end if
				Response.Write sTemp
			end if			
%>
		</table>

    <br>
		<table border="1" cellpadding="1" cellspacing="1" width="100%" COLS=4 bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td WIDTH="25%" bgcolor="#cccccc"><strong>Address:</strong></td>
		    <td WIDTH="25%" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><b><%= adoRS("Address")%></b></td>
		    <td WIDTH="25%" bgcolor="#cccccc"><strong>Birthdate:</strong></td>
		    <td WIDTH="25%" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><b><%= adoRS("BirthDate")%></b></td>
			</tr>
		  <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td bgcolor="#F0F0F0">&nbsp;</td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor=F0F0F0><b><%= adoRS("City")%> ,  <%=adoRS("State") + " " +adoRS("PostalCode") %></b></td>
		    <td bgcolor=F0F0F0><strong>Age:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor=F0F0F0><b><%= DateDiff("d", adoRS("BirthDate"), Now) \ 365%></b></td>
		  </tr>
		  <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
<%
				if not m_bDependent then
%>
				<td bgcolor="#cccccc"><strong>Home Phone:</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><b><%= adoRS("HomePhone")%></b></td>
<%
				else
%>
				<td bgcolor="#cccccc">&nbsp;</td>
				<td bgcolor="#cccccc">&nbsp;</td>
<%
				end if
%>
			  <td bgcolor="#cccccc"><strong>Gender:</strong></td>
			  <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><b><%= adoRS("Gender")%></b></td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
<%
				if not m_bDependent then
%>
				<td bgcolor="#F0F0F0"><strong>Work Phone:</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><%= adoRS("WorkPhone")%></b></td>
<%
				else
%>
				<td bgcolor="#F0F0F0"><strong>Dependent SSN:</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><%= sShownSSN %></b></td>
<%
				end if
				if not m_bDependent then
%>
				<td bgcolor="#F0F0F0"><strong>Marital Status:</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><%= adoRS("MaritalStatus")%></b></td>
<%
				else
%>
				<td bgcolor="#F0F0F0"><strong>Relationship:</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><%= adoRS("Relationship")%></b></td>
<%
				end if
%>
			</tr>
<%
			if not m_bDependent then
%>
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td bgcolor="#CCCCCC"><strong>Locked File:</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#CCCCCC"><% if adoRS("locktype") <> "" then %><img SRC="images/lock1.gif" width="32" height="32" alt="LOCKED FILE!!">&nbsp;&nbsp;<b>Yes</b><% else %><b>No</b><% end if %>&nbsp;&nbsp;<b><%= adoRS("locktype") %></b></td>
				<td bgcolor="#CCCCCC"><strong>Retirement Date:</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#CCCCCC"><b><%= adoRS("RetireDate")%></b></td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td bgcolor="#F0F0F0">&nbsp;</td>
				<td bgcolor="#F0F0F0">&nbsp;</td>
				<td bgcolor="#F0F0F0"><strong>Date of Death:</strong></td>
				<td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= adoRS("DeathDate")%></b></td>
			</tr>
<%  
      else
%>
    <tr BorderColorDark="#FCFCFC" BorderColorlight="#999999">  
      <td bgcolor="#F0F0F0"><strong>Locked Files:</strong></td>
		  <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><img SRC="images/lock1.gif" width="32" height="32" alt="LOCKED FILE!!">&nbsp;&nbsp;<b><%= adoRS("locktype") %></b></td>
      <td bgcolor="#F0F0F0" colspan="2"><strong>&nbsp;</strong></td>
    </tr>
<%
			end if
%>
		</table>
    <br>
		<hr SIZE="3" NOSHADE>
<%
		adoRS.Close
'**************
' Dependents
'**************
		if not m_bDependent then
%>
			<br>	
			<font SIZE="3" face="arial, helvetica"><center><b>Dependents</b></center></font>
			<table border="1" cellpadding="2" cellspacing="2" width="100%" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
				<tr ALIGN="CENTER" BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				    <td bgcolor="#CCCCCC"><strong>Dep. #</strong></td>
				    <td bgcolor="#CCCCCC" colspan="2"><strong>Name</strong></td>
				    <td bgcolor="#CCCCCC"><strong>SSN</strong></td>
				    <td bgcolor="#CCCCCC"><strong>Gender</strong></td>
				    <td bgcolor="#CCCCCC"><strong>Relationship</strong></td>
				    <td bgcolor="#CCCCCC"><strong>Birthdate</strong></td>
            <td bgcolor="#CCCCCC" colspan="2"><strong>Locked File</strong></td>
				</tr>
<%
				sSQL = "select SSN = n.DepSSN, n.DepNumber, n.LastName + ', ' + "
				sSQL = sSQL & " n.FirstName + ' ' + n.MiddleName 'Name', n.BirthDate,"
				sSQL = sSQL & " n.Gender, n.locktype, n.LockedReason, n.LockedBy, n.LockedDate, n.SupervisorAccessOnly, r.Description FROM "
				sSQL = sSQL & " Names n LEFT JOIN RelationDescription r "
				sSQL = sSQL & " on (r.RelationCode = n.RelationCode) WHERE n.SSN = " & m_iSSN  
				sSQL = sSQL & " AND n.DepNumber !=0 ORDER BY n.DepNumber"			
				adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
				
        if not adoRS.EOF then
					Do While Not adoRS.EOF
%>
						<tr ALIGN="CENTER" BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
						    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><%=adoRS("DepNumber")%></b></td>
                <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><% if adoRS("locktype") <> "" then %><img SRC="images/lock1.gif" width="32" height="32" alt="LOCKED FILE!!"><% else %>&nbsp;<% end if %></td>
                <td align="left" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><a href="PersonDetails.asp?SSN=<%= m_iSSN %>&DepNum=<%= adoRS("DepNumber") %>&locktype=<%= adoRS("locktype") %>&Admin=<%= Admin %>" onClick="WorkingStatus()"><b><%= adoRS("Name") %></b></a></td>
<%
							sTemp = "<TD ALIGN=LEFT bgcolor='#F0F0F0'><A HREF='PersonLockDetail.asp?SSN=" & m_iSSN & "&DepNum=" & adoRS("DepNumber") & "' onClick='WorkingStatus()'><b>" & adoRS("Name") & "</b></A></TD>"
							'Response.Write sTemp
%>					    
						    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b>&nbsp;<%=adoRS("SSN")%></b></td>
						    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><%=adoRS("Gender")%></b></td>
						    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><%=adoRS("Description")%></b></td>
						    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><%=adoRS("BirthDate")%></b></td>
                <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><% if adoRS("locktype") <> "" then %><b>Yes</b><% else %><b>No</b><% end if %></td>
                <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><%=adoRS("locktype")%></b></td>
						</tr>
            
<% 
						adoRS.MoveNext
					Loop 
				else
%>
					<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC" align="center">
					    <td COLSPAN="7" bgcolor="#F0F0F0"><font color="red" face="verdana, arial, helvetica" size="2"><b>No dependents found.</b></font></td>
					</tr>
<%
				end if  
%>
      
      </table>
<%
			adoRS.Close
		end if
  end if
%>






<% end if %>
<% end if %>
<% 
%>
</body>
</html>