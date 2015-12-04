<%@ language=vbscript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim i, j, adoConn, adoRS, adoCmd, adoParam, sSQL, sTemp, sShownSSN
Dim m_iSSN, m_iDepNo, m_bDependent, m_sParticName, m_bUseEAS, m_sBirthDate
Dim avFundPlan(), avEligData(), avSchools(), avSchoolData(), iNumSchools, dep_FName, dep_LName
Dim iNumFundPlans, iOldestYear, iNewestYear, iNumYears, iNewestMonth, iOldestMonth
Dim iCurrRow, iCurrCol, sSectionTitle, avRowData(), iCurrYear
Dim avMonthName(12)
Dim sAttend, iStartCol, iEndCol, sCurrSchool, locktype, Manager, Supervisor, Auditor, OtherArea
Dim m_sFund, Carrier, Reason, ParticAdmin, ParticLocktype, Admin, sTemp2, sTemp3


Set adoConn = Server.CreateObject("ADODB.Connection")
adoConn.CommandTimeout = 0

Set adoRS = Server.CreateObject("ADODB.Recordset")
adoConn.Open Application("DataConn")

sSQL = "select * from UserInformation where Logonid='" & Session("User") & "'"
adoRS.Open sSQL, adoConn, adOpenKeyset, adLockOptimistic

Manager = adoRS("IsManager")
Supervisor = adoRS("IsSupervisor")
Auditor = adoRS("IsAuditor")

adoRS.Close
				
Function TranslateCoverages(sCOBString)
    '***************************************************************************************************
    ' The defined coverages (where i=character position within the sCOBString) are:
    '	i	Coverage Type
    '----	--------------
    '	1	Medical
    '	2	Dental
    '	3	Orthodontics
    '	4	Vision
    '	5	Prescription
    '	6	Other
    '	7	Chiropractic
    '***************************************************************************************************
    Dim sCoverages, sCovTemplate
    Dim i

	sCovTemplate = "MDOVP C"
	For i=1 to len(sCovTemplate)
		If i <> 6 Then
			If mid(sCOBString,i,1) = "Y" Then
				sCoverages= sCoverages & mid(sCovTemplate,i,1) & " "
			End If
		End If
	Next
	If mid(sCOBString,6,1) = "Y" Then
		sCoverages= sCoverages & "Other"
	End If
	
	TranslateCoverages = sCoverages
End Function

If Request.QueryString("UseEAS") = "" Then
	m_bUseEAS = false
Else
	m_bUseEAS = Request.QueryString("UseEAS")
End If

If request.querystring("OtherArea") = "" Then
  OtherArea = false
Else
  OtherArea = request.querystring("OtherArea")
End If

locktype = request.querystring("locktype")
m_iSSN = Request.QueryString("SSN")
m_iDepNo = Request.QueryString("DepNum")
OtherArea = request.querystring("OtherArea")
Admin = request.querystring("Admin")  

'response.write Manager & Supervisor & Auditor & "<br />"
'response.write "OtherArea= " & OtherArea & "<br />"
'response.write "m-bUseEAS= " & m_bUseEAS
'if locktype <> "" then
  'locktype = True
'else
  'locktype = False
'end if

If locktype <> "" Then
    If Manager = "true" or Supervisor = "true" or Auditor = "true" Then
    Else
        response.redirect "PersonLockDetail.asp?SSN=" & m_iSSN & "&DepNum=" & m_iDepNo & "&locktype=" & locktype & "&Admin=" & Admin
    End if
End If

If Admin Then
    If Manager = "true" or Supervisor = "true" or Auditor = "true" Then
    Else
        response.redirect "PersonLockDetail.asp?SSN=" & m_iSSN & "&DepNum=" & m_iDepNo & "&Admin=" & Admin
    End if
End If
%>

<html>
<head>

<style type="text/css">
    /* Sliding Menu Style Sheet */
    .sliderView {position: absolute; left: 0pt}
    .sliderView div.slider {float: left; cursor: default; background: lightgray; border-top: 1pt #EEEEEE solid; border-left: 1pt #EEEEEE solid;  border-bottom: 1pt gray solid; border-right: 1pt gray solid; width: 1.3em; text-align: center; font-size: 10pt}
    .sliderView span.arrow {font: 12pt webdings}
    .sliderView .contents {width: 90%; float: left; background: lightgray; border-top: 1pt #EEEEEE solid; border-left: 1pt #EEEEEE solid;  border-bottom: 1pt gray solid; border-right: 1pt gray solid;}
</style>
<script type="text/javascript">

    function slide(dir, index) {
        var elSlide = animArray[index];
        var elMenu = elSlide.parentElement;
        if (!dir) {
            elMenu.style.pixelLeft -= elMenu._offset;
            if (elSlide.offsetLeft <= (-elMenu.style.pixelLeft)) {
                elMenu.style.pixelLeft = -elSlide.offsetLeft;
                elMenu._arrow.innerText = 4;
            } else {
                setTimeout("slide(" + dir + "," + index + ")", 15);
            }
        } else {
            elMenu.style.pixelLeft += elMenu._offset;
            if (elMenu.style.pixelLeft >= 0) {
                elMenu.style.pixelLeft = 0;
                elMenu._arrow.innerText = 3;
            } else {
                setTimeout("slide(" + dir + "," + index + ")", 15);
            }
        }
    }

    // Used to cache animated element
    var animArray = new Array();

    function doSlide(src) {
        el = src.parentElement;
        el._offset = src.offsetLeft / 10
        el._arrow = src.children.tags("span")[0];
        if (el._index == null) {
            el._index = animArray.length;
            animArray[animArray.length] = src;
        }
        if (el.style.pixelLeft != -src.offsetLeft) {
            slide(false, el._index); // Slide in
        } else {
            slide(true, el._index);   // Slide out
        }
    }
</script>

<% 			
	Sub PickButton(iIndex)
		Select Case iIndex
			Case 0
				top.AppStatus.location.replace("Working.htm")
				ClaimPost.SSN.value = <%= m_iSSN %>
				ClaimPost.DepNum.value = <%= m_iDepNo %>
				ClaimPost.submit
			Case 1 
				top.AppStatus.location.replace("Working.htm")
				FormLetterPost.SSN.value = <%= m_iSSN %>
				FormLetterPost.submit
			Case 2 
				top.AppStatus.location.replace("Working.htm")
				LetterPost.SSN.value = <%= m_iSSN %>
				LetterPost.submit
			Case 3 
				top.AppStatus.location.replace("Working.htm")
				CheckPost.SSN.value = <%= m_iSSN %>
				CheckPost.DepNum.value = <%= m_iDepNo %>
				CheckPost.submit
			Case 4 
				top.AppStatus.location.replace("Working.htm")
				PhoneCallPost.Crit.value = <%= m_iSSN %>
				PhoneCallPost.submit
			Case Else
				msgbox "none"
		End Select
	End Sub
			
	Function ShowPlanDesc(Fund,Plan,Rate)
		Dim sResult

		sResult=showModalDialog("ACSInfo.asp?EASFund=" & Fund & "&EASPlan=" & Plan & "&EASRate=" & Rate, "ACS Information", "dialogWidth:500px; dialogHeight:295px; help:no;")
				
	End Function

%>
<!--#include file="VBFuncs.inc" -->
</head>
<body  onload="UpdateScreen(0)">
<link rel="stylesheet" href="styles/CritTable.css" />

    
<%

if m_iSSN = "" or m_iDepNo = "" then
	Response.Write ("<p><ul><font color='red' face='verdana, arial, helvetica' size='2'><b>SSN or DepNumber missing; contact your network administrator.</b></font></ul>")
else
	set adoConn = Server.CreateObject("ADODB.Connection")
'**********************************************************************************
' Setting CommandTimeout to 0 waits indefinitely for response; note that per 
' ADO documentation Command object will not inherit the Connection setting (default
' is 30 seconds).  This appeared to help with response issues on seasql02.
'**********************************************************************************
    adoConn.CommandTimeout = 0
	'set adoCmd = Server.CreateObject("ADODB.Command")
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
    'response.write OtherArea
	
	if OtherArea = "true" then
        if m_iDepNo = 0 then
	        m_bDependent = false
    
            sSQL = "select n.FirstName + ' ' + n.MiddleName + ' ' + n.LastName 'Name'," 
		    sSQL = sSQL & " n.ssn, n.BirthDate, n.Gender, n.Fund, "
            sSQL = sSQL & " h.StreetAddress Address, "
            sSQL = sSQL & " h.City, h.State, h.PostalCode,"
		    sSQL = sSQL & " i.PhoneNumber HomePhone, i.WorkPhone, m.Description 'MaritalStatus' "
            sSQL = sSQL & "  FROM OtherAreaNames n LEFT JOIN OtherAreaHomes"
            sSQL = sSQL & " h on (h.SSN = n.SSN) LEFT JOIN Insureds i on (i.SSN = n.SSN) LEFT JOIN"
		    sSQL = sSQL & " MaritalStatuses m on (i.MaritalStatus = m.MaritalStatus) "
            sSQL = sSQL & " WHERE n.SSN = " & m_iSSN & " AND n.DepNumber = " & m_iDepNo
      
        else
            m_bDependent = true
            sSQL = "select n.BirthDate, n.locktype, n.FirstName + ' ' + n.MiddleName + ' ' + n.LastName 'Name'" 
	        sSQL = sSQL & " FROM" 
		    sSQL = sSQL & " OtherAreaNames"
		    sSQL = sSQL & " n WHERE n.SSN = " & m_iSSN & " AND n.DepNumber = 0"  
		
            adoRS.Open sSql, adoConn, adOpenForwardOnly, adLockOptimistic
		    m_sParticName = adoRS("Name")
            m_sBirthDate = adoRS("BirthDate")
		
            adoRS.Close
      
            sSQL = "select n.FirstName + ' ' + n.MiddleName + ' ' + n.LastName 'Name'," 
		    sSQL = sSQL & " n.ssn, n.DepSSN, n.BirthDate, n.Gender, n.RelationCode,"
            sSQL = sSQL & "h.StreetAddress Address,"
            sSQL = sSQL & " h.City, h.State, h.PostalCode, r.Description Relationship"
            sSQL = sSQL & " FROM OtherAreaNames n LEFT JOIN OtherAreaHomes"
            sSQL = sSQL & " h ON (h.SSN = n.SSN)"
		    sSQL = sSQL & " LEFT JOIN RelationDescription r ON (n.RelationCode = r.RelationCode)"
		    sSQL = sSQL & " WHERE n.SSN = " & m_iSSN & " AND n.DepNumber = " & m_iDepNo
      
        end if
        if m_bDependent then
		    sTemp = "Dependent Information"
	    else
		    sTemp = "Participant Information"
	    end if
        if OtherArea = "true" then
		    sTemp = sTemp & " [Other Area Data]"
	    end if
    end if    
%>
	<table width="100%" cols="3" border="0">
		<tr>
			<td align="center" width="10%">
<%
				if not Session("IsClerk") then
%>
					<img SRC="images/log.gif" onclick="LogCall()">
<%
				end if
%>
			</td>
			<td align="center">
				<font size="+2" face="verdana, arial, helvetica" color="red"><strong><%= sTemp%></strong></font>
			</td>
			<td align="center" width="20%">
				<img SRC="images/bluebar2.gif" onclick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif" onclick="history.go(+1)" border="0">
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
    end if
%>
    <table width="70%" align="center" border="0">
		<tr>
			<td width="5%">&nbsp;</td>
            <td align="center">
				<input type="button" value="Claims" onclick="PickButton(0)" />
				<input type="button" value="ACELetters" onclick="PickButton(1)" />
				<input type="button" value="ACSLetters" onclick="PickButton(2)" />
				<input type="button" value="Checks" onclick="PickButton(3)" />
				<input type="button" value="Phone Calls" onclick="PickButton(4)" />
			</td>
            <td width="20%">&nbsp;</td>
		</tr>
	</table>

	<form id="ClaimPost" target="Details" action="ClaimSearch.asp" method="get">
		<input type="hidden" id="SSN" name="SSN" value>
		<input type="hidden" id="DepNum" name="DepNum" value>
	</form>
	<form id="FormLetterPost" target="Details" action="FormLetterSearch.asp" method="get">
		<input type="hidden" id="SSN" name="SSN" value>
	</form>
	<form id="LetterPost" target="Details" action="LetterSearch.asp" method="get">
		<input type="hidden" id="SSN" name="SSN" value>
	</form>
	<form id="CheckPost" target="Details" action="CheckSearch.asp" method="get">
		<input type="hidden" id="SSN" name="SSN" value>
		<input type="hidden" id="DepNum" name="DepNum" value>
	</form>
	<form id="PhoneCallPost" target="Details" action="PhoneSearch.asp" method="get">
		<input type="hidden" id="Crit" name="SSN" value>
	</form>
	<br clear="left">    
		<% if locktype <> "" then %>
      <p><font face="verdana, arial, helvetica" size="2"><b>This is a locked file for the participant <%= adoRS("Name") %></b></font></p>
    <% end if %>
     
    
    <!--- <p><a href="PersonDetails.asp onclick=LogCall(<%= adoRS("SSN")%>)">Log Call</a> --->
    <table border="1" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif;width:100%;">
			<colgroup span="2" />
			<colgroup span="2" />
<%
			sTemp = "<tr borderColorDark='#f0f0f0' borderColorlight='#999999'><td width=25% bgcolor='#cccccc'><strong>"
			if m_bDependent then
				sTemp = sTemp & "Dependent Code:"
			else
				sTemp = sTemp & "SSN:"
			end if
			sTemp = sTemp & "</strong></td><td width=25% style='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'>"
			if m_bDependent then
				sTemp = sTemp &"<b>" & m_iDepNo
			else
        if len(sShownSSN) = 7 then
          sTemp = sTemp &"<b>00" & sShownSSN
        else
				  sTemp = sTemp &"<b>" & sShownSSN
        end if
			end if
			sTemp = sTemp & "</b></td><td width=25% bgcolor='#cccccc'><strong>Name:</strong>"
			sTemp = sTemp & "<td width=25% style='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'><b>" & adoRS("Name") & "</b></td></tr>"
			Response.Write sTemp			
			if m_bDependent then
				sTemp = "<TR borderColorDark='#f0f0f0' borderColorlight='#999999'><td bgcolor='#cccccc'><strong>Participant SSN:</strong></td>"
				sTemp = sTemp & "<td style='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'><b>" & adoRS("SSN") & "</b></td>"
				sTemp = sTemp & "<td bgcolor='#cccccc'><strong>Participant Name:</strong></td>"
				sTemp = sTemp & "<td style='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'>"
        sTemp = sTemp & "<A href='PersonDetails.asp?SSN=" & m_iSSN & "&DepNo=0' onclick='WorkingStatus()'><b>" & m_sParticName & "</b></A>"
				sTemp = sTemp & "</td></tr>"
				Response.Write sTemp
			end if			
%>
		</table>
    <br />
		<table border="1" cellpadding="1" cellspacing="1" width="100%" cols=4 bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr borderColorDark="#f0f0f0" borderColorlight="#999999">
				<td width="25%" bgcolor="#cccccc"><strong>Address:</strong></td>
		    <td width="25%" style='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><b><%= adoRS("Address")%></b></td>
		    <td width="25%" bgcolor="#cccccc"><strong>Birthdate:</strong></td>
		    <td width="25%" style='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><b><%= adoRS("BirthDate")%></b></td>
			</tr>
		  <tr borderColorDark="#fcfcfc" borderColorlight="#cccccc">
				<td bgcolor="#f0f0f0">&nbsp;</td>
		    <td style='color:<%= Session("EmphColor")%>;' bgcolor=F0F0F0><b><%= adoRS("City")%> ,  <%=adoRS("State") + " " +adoRS("PostalCode") %></b></td>
		    <td bgcolor=F0F0F0><strong>Age:</strong></td>
		    <td style='color:<%= Session("EmphColor")%>;' bgcolor=F0F0F0><b><%= DateDiff("d", adoRS("BirthDate"), Now) \ 365%></b></td>
		  </tr>
		  <tr borderColorDark="#f0f0f0" borderColorlight="#999999">
<%
				if not m_bDependent then
%>
				<td bgcolor="#cccccc"><strong>Home Phone:</strong></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><b><%= adoRS("HomePhone")%></b></td>
<%
				else
%>
				<td bgcolor="#cccccc">&nbsp;</td>
				<td bgcolor="#cccccc">&nbsp;</td>
<%
				end if
%>
			  <td bgcolor="#cccccc"><strong>Gender:</strong></td>
			  <td style='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><b><%= adoRS("Gender")%></b></td>
			</tr>
			<tr borderColorDark="#fcfcfc" borderColorlight="#cccccc">
<%
				if not m_bDependent then
%>
				<td bgcolor="#f0f0f0"><strong>Work Phone:</strong></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><b><%= adoRS("WorkPhone")%></b></td>
<%
				else
%>
				<td bgcolor="#f0f0f0"><strong>Dependent SSN:</strong></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><b><%= sShownSSN%></b></td>
<%
				end if
				if not m_bDependent then
%>
				<td bgcolor="#f0f0f0"><strong>Marital Status:</strong></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><b><%= adoRS("MaritalStatus")%></b></td>
<%
				else
%>
				<td bgcolor="#f0f0f0"><strong>Relationship:</strong></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><b><%= adoRS("Relationship")%></b></td>
<%
				end if
%>
			</tr>
<%
			if not m_bDependent then
%>
			<tr borderColorDark="#f0f0f0" borderColorlight="#999999">
				<td bgcolor="#cccccc"><strong>Fund:</strong></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><b><%= adoRs("Fund") %></b></td>
				<td bgcolor="#cccccc"><strong>Retirement Date:</strong></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><b></b></td>
			</tr>
			<tr borderColorDark="#fcfcfc" borderColorlight="#cccccc">
				<td bgcolor="#f0f0f0">&nbsp;</td>
				<td bgcolor="#f0f0f0">&nbsp;</td>
				<td bgcolor="#f0f0f0"><strong>Date of Death:</strong></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b></b></td>
			</tr>
<%
			end if
%>
		</table>
    <br />
		<hr size="3" noshade>
<%
		adoRS.Close      
end if

else 
  sSQL = "select Fund from Tracking where SSN=" & m_iSSN 
  adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
    if adoRS.EOF then
      m_sFund = 028
    else
      m_sFund = adoRS("Fund")
    end if
  adoRS.Close
  
  
  sSQL = "select locktype from Names where SSN =" & m_iSSN & " AND DepNumber > 0 and locktype is null"
  adoRS.Open sSql, adoConn, adOpenForwardOnly, adLockOptimistic
  
  if adoRS.EOF then
    sTemp2 = " the <font color='red'>Full Family</font> of "
  end if
  
  adoRS.Close
  
  sSQL = "select SupervisorAccessOnly from Names where SSN =" & m_iSSN & " AND DepNumber > 0 and SupervisorAccessOnly = 0"
  adoRS.Open sSql, adoConn, adOpenForwardOnly, adLockOptimistic
  
  if adoRS.EOF then
    sTemp3 = " the <font color='red'>Full Family</font> of "
  end if
  
  adoRS.Close 
  
  
  if m_iDepNo = 0 then
		m_bDependent = false
    'response.write m_bUseEAS

		sSQL = "select n.FirstName + ' ' + n.MiddleName + ' ' + n.LastName 'Name'," 
		sSQL = sSQL & " n.ssn, n.BirthDate, n.Gender, o.Fund, "
		if not m_bUseEAS then
			sSQL = sSQL & " h.StreetAddress Address, "
		else
			sSQL = sSQL & " h.Addr1 Address, "
		end if
		sSQL = sSQL & " h.City, h.State, h.PostalCode,"
		sSQL = sSQL & " i.PhoneNumber HomePhone, i.WorkPhone, m.Description 'MaritalStatus',"
		if not m_bUseEAS then
			sSQL = sSQL & " e.RetireDate, e.DeathDate, n.locktype, n.LockedBy, n.LockedDate, n.LockedReason, n.SupervisorAccessOnly FROM Names n LEFT JOIN Homes"
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
		sSQL = sSQL & " n.ssn, n.DepSSN, n.BirthDate, n.Gender, n.RelationCode, n.locktype, n.LockedBy, n.LockedDate, n.LockedReason, n.SupervisorAccessOnly, "
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
	if m_bDependent then
		sTemp = "Dependent Information"
	else
		sTemp = "Participant Information"
	end if
	if m_bUseEAS then
		sTemp = sTemp & " [EAS Data]"
	end if
%>
	<table width="100%" cols="3" border="0">
		<tr>
			<td align="center" width="10%">
<%
				if not Session("IsClerk") then
%>
					<img SRC="images/log.gif" onclick="LogCall()">
<%
				end if
%>
			</td>
			<td align="center">
				<font size="+2" face="verdana, arial, helvetica" color="red"><strong><%= sTemp%></strong></font>
			</td>
			<td align="center" width="20%">
				<img SRC="images/bluebar2.gif" onclick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif" onclick="history.go(+1)" border="0">
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
%>
    <table width="70%" align="center" border="0">
			<tr>
				<td width="5%">&nbsp;</td>
        <td align="center">
					<!--- <input type="button" value="Claims" onclick="PickButton(0)"> --->
					<input type="button" value="ACELetters" onclick="PickButton(1)">
					<input type="button" value="ACSLetters" onclick="PickButton(2)">
					<input type="button" value="Checks" onclick="PickButton(3)">
					<input type="button" value="Phone Calls" onclick="PickButton(4)">
          <input type="button" value="Prescription" onclick="PickButton(5)">
          <input type="button" value="Tracking" onclick="PickButton(6)">
          <!--- <input type="button" value="VisionClaims" onclick="PickButton(7)">
          <input type="button" value="BlankClaims" onclick="PickButton(8)"> --->
				</td>
        <td width="20%">&nbsp;</td>
			</tr>
      <tr>
				<td width="5%">&nbsp;</td>
        <td align="center">
					<input type="button" value="Claims" onclick="PickButton(0)">
          <input type="button" value="VisionClaims" onclick="PickButton(7)">
          <!--- <input type="button" value="BlankClaims" onclick="PickButton(8)"> --->
				</td>
        <td width="20%">&nbsp;</td>
			</tr>
      <tr>
				<td width="5%">&nbsp;</td>
        <td align="center">
	        
          <% if Admin then %>
            <font face="verdana, arial, helvetica" size="2">
            <p><img src="images/lock4.gif" width="32" height="32" border="0" alt="Locked File">&nbsp;<b>This file has an <font color="red">ADMINISTRATIVE LOCK</font> for <% if m_bDependent then %>dependent&nbsp;<% else %><%= sTemp3 %>&nbsp;<% end if %> <%= adoRS("Name") %></b>
            <br /><b>Reason: <% if adoRS("LockedReason") <> "" then %><%= adoRS("LockedReason") %><% else %>NO REASON IN DATABASE!!<% end if %></b>
            <br /><b>Locked By: <% if adoRS("LockedBy") <> "" then %><%= adoRS("LockedBy") %><% else %>NO NAME IN DATABASE!!<% end if %></b>
            <br /><b>Date Locked: <% if adoRS("LockedDate") <> "" then %><%= formatdatetime(adoRS("LockedDate"),vbShortDate) %><% else %>NO DATE IN DATABASE!!<% end if %></b>
            </font>
          <% else %>
          <% if adoRS("locktype") <> "" then %>
            <font face="verdana, arial, helvetica" size="2">
            <p><img src="images/lock4.gif" width="32" height="32" border="0" alt="Locked File">&nbsp;<b>This is a locked file for <% if m_bDependent then %>dependent&nbsp;<% else %><%= sTemp2 %>&nbsp;<% end if %> <%= adoRS("Name") %></b>
            <br /><b>Reason: <% if adoRS("LockedReason") <> "" then %><%= adoRS("LockedReason") %><% else %>NO REASON IN DATABASE!!<% end if %></b>
            <br /><b>Locked By: <% if adoRS("LockedBy") <> "" then %><%= adoRS("LockedBy") %><% else %>NO NAME IN DATABASE!!<% end if %></b>
            <br /><b>Date Locked: <% if adoRS("LockedDate") <> "" then %><%= formatdatetime(adoRS("LockedDate"),vbShortDate) %><% else %>NO DATE IN DATABASE!!<% end if %></b>
            </font>
          <% end if 
          end if%>
          		    
				</td>
        <td width="20%">&nbsp;</td>
			</tr>
		</table>
		<form id="ClaimPost" target="Details" action="ClaimSearch.asp" method="get">
			<input type="hidden" id="SSN" name="SSN" value>
			<input type="hidden" id="DepNum" name="DepNum" value>
		</form>
		<form id="FormLetterPost" target="Details" action="FormLetterSearch.asp" method="get">
			<input type="hidden" id="SSN" name="SSN" value>
		</form>
		<form id="LetterPost" target="Details" action="LetterSearch.asp" method="get">
			<input type="hidden" id="SSN" name="SSN" value>
		</form>
		<form id="CheckPost" target="Details" action="CheckSearch.asp" method="get">
			<input type="hidden" id="SSN" name="SSN" value>
			<input type="hidden" id="DepNum" name="DepNum" value>
		</form>
		<form id="PhoneCallPost" target="Details" action="PhoneSearch.asp" method="get">
			<input type="hidden" id="Crit" name="SSN" value>
		</form>
    <form id="PrescriptionPost" target="Details" action="Prescription.asp" method="get">
			<input type="hidden" id="SSN" name="SSN" value>
      <input type="hidden" id="DepNum" name="DepNum" value>
		</form>
    <form id="TrackingPost" target="Details" action="Tracking.asp" method="get">
			<input type="hidden" id="SSN" name="SSN" value>
      <input type="hidden" id="DepNum" name="DepNum" value>
      <input type="hidden" id="Years" name="Years" value>
		</form>
    <form id="VisionClaimPost" target="Details" action="ClaimSearch.asp" method="get">
			<input type="hidden" id="SSN" name="SSN" value>
			<input type="hidden" id="DepNum" name="DepNum" value>
      <input type="hidden" id="HWType" name="HWType" value>
		</form>
    <form id="BlankClaimPost" target="Details" action="ClaimSearch.asp" method="get">
			<input type="hidden" id="UserCriteria" name="UserCriteria" value>
		</form>
		<br clear="left">
     
    
    <!--- <p><a href="PersonDetails.asp onclick=LogCall(<%= adoRS("SSN")%>)">Log Call</a> --->
    <table border="1" width="100%" cols=4 rules=groups bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<colgroup span=2>
			<colgroup span=2>
<%
			sTemp = "<tr borderColorDark='#f0f0f0' borderColorlight='#999999'><td width=25% bgcolor='#cccccc'><strong>"
			if m_bDependent then
				sTemp = sTemp & "Dependent Code:"
			else
				sTemp = sTemp & "SSN:"
			end if
			sTemp = sTemp & "</strong></td><td width=25% style='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'>"
			if m_bDependent then
				sTemp = sTemp &"<b>" & m_iDepNo
			else
        if len(sShownSSN) = 8 then
          sTemp = sTemp &"<b>0" & sShownSSN
        elseif len(sShownSSN) = 7 then
          sTemp = sTemp &"<b>00" & sShownSSN
        else
				  sTemp = sTemp &"<b>" & sShownSSN
        end if
			end if
			sTemp = sTemp & "</b></td><td width=25% bgcolor='#cccccc'><strong>Name:</strong>"
			sTemp = sTemp & "<td width=25% style='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'><b>" & adoRS("Name") & "</b></td></tr>"
			Response.Write sTemp			
			if m_bDependent then
				sTemp = "<TR borderColorDark='#f0f0f0' borderColorlight='#999999'><td bgcolor='#cccccc'><strong>Participant SSN:</strong></td>"
				sTemp = sTemp & "<td style='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'><b>" & adoRS("SSN") & "</b></td>"
				sTemp = sTemp & "<td bgcolor='#cccccc'><strong>Participant Name:</strong></td>"
				sTemp = sTemp & "<td style='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'>"
        if ParticLocktype <> "" then
          if Manager = "true" or Supervisor = "true" or Auditor = "true" then
				    sTemp = sTemp & "<A href='PersonDetails.asp?SSN=" & m_iSSN & "&Admin=" & ParticAdmin & "&locktype=" & ParticLocktype & "&DepNum=0' onclick='WorkingStatus()'><b>" & m_sParticName & "</b></A>"
				    sTemp = sTemp & "</td></tr>"
          else
            sTemp = sTemp & "<A href='PersonLockDetail.asp?SSN=" & m_iSSN & "&locktype=" & ParticLocktype & "&DepNum=0&Admin=" & ParticAdmin
            sTemp = sTemp & "' onclick='WorkingStatus()'><b>" & m_sParticName & "</b></a>"
				    sTemp = sTemp & "</td></tr>"
          end if
        else
          sTemp = sTemp & "<A href='PersonDetails.asp?SSN=" & m_iSSN & "&DepNum=0' onclick='WorkingStatus()'><b>" & m_sParticName & "</b></A>"
				  sTemp = sTemp & "</td></tr>"
          'sTemp = sTemp & "<b>" & m_sParticName & "</b></td></tr>"
        end if
				Response.Write sTemp
			end if			
%>
		</table>
    
    <div class=sliderView style="top: 800pt; Left: -1000px">
      <div class=contents style="background: lightgrey; width: 1000; padding-left: 5pt"><font face="verdana, arial, helvetica" size="2">
        <P class=start><center><b>Participant/Dependent Information</b></center></font> 
        <table border="1" width="100%" cols=4 rules=groups bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			  <colgroup span=2>
			  <colgroup span=2>
<%
			sTemp = "<tr borderColorDark='#f0f0f0' borderColorlight='#999999'><td width=25% bgcolor='#cccccc'><strong>"
			if m_bDependent then
				sTemp = sTemp & "Dependent Code:"
			else
				sTemp = sTemp & "SSN:"
			end if
			sTemp = sTemp & "</strong></td><td width=25% style='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'>"
			if m_bDependent then
				sTemp = sTemp &"<b>" & m_iDepNo
			else
				sTemp = sTemp &"<b>" & sShownSSN
			end if
			sTemp = sTemp & "</b></td><td width=25% bgcolor='#cccccc'><strong>Name:</strong>"
			sTemp = sTemp & "<td width=25% style='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'><b>" & adoRS("Name") & "</b></td></tr>"
			Response.Write sTemp			
			if m_bDependent then
				sTemp = "<TR borderColorDark='#f0f0f0' borderColorlight='#999999'><td bgcolor='#cccccc'><strong>Participant SSN:</strong></td>"
				sTemp = sTemp & "<td style='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'><b>" & adoRS("SSN") & "</b></td>"
				sTemp = sTemp & "<td bgcolor='#cccccc'><strong>Participant Name:</strong></td>"
				sTemp = sTemp & "<td style='color:" & Session("EmphColor") & ";' bgcolor='#cccccc'>"
				sTemp = sTemp & "<A href='PersonDetails.asp?SSN=" & m_iSSN & "&DepNo=0' onclick='WorkingStatus()'><b>" & m_sParticName & "</b></A>"
				sTemp = sTemp & "</td></tr>"
				Response.Write sTemp
			end if			
%>
		    </table>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
      </div>
      <div style="background: #63659C; color: white;" ONSELECTSTART="return false" onclick="doSlide(this)" class="slider"><span class=arrow>3</span><font color="white" face="verdana, arial, helvetica" size="2"><b><br />P<br />A<br />R<br />T<br />I<br />C<br />I<br />P<br />A<br />N<br />T</b></font></div>
    </div>
    
    
    <br />
		<table border="1" cellpadding="1" cellspacing="1" width="100%" cols=4 bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr borderColorDark="#f0f0f0" borderColorlight="#999999">
				<td width="25%" bgcolor="#cccccc"><strong>Address:</strong></td>
		    <td width="25%" style='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc" colspan="2"><b><%= adoRS("Address")%></b></td>
		    <td width="25%" bgcolor="#cccccc"><strong>Birthdate:</strong></td>
		    <td width="25%" style='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><b><%= adoRS("BirthDate")%></b></td>
			</tr>
		  <tr borderColorDark="#fcfcfc" borderColorlight="#cccccc">
				<td bgcolor="#f0f0f0">&nbsp;</td>
		    <td style='color:<%= Session("EmphColor")%>;' bgcolor=F0F0F0 colspan="2"><b><%= adoRS("City")%> ,  <%=adoRS("State") + " " +adoRS("PostalCode") %></b></td>
		    <td bgcolor=F0F0F0><strong>Age:</strong></td>
		    <td style='color:<%= Session("EmphColor")%>;' bgcolor=F0F0F0><b><%= DateDiff("d", adoRS("BirthDate"), Now) \ 365%></b></td>
		  </tr>
		  <tr borderColorDark="#f0f0f0" borderColorlight="#999999">
<%
				if not m_bDependent then
%>
				<td bgcolor="#cccccc"><strong>Home Phone:</strong></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc" colspan="2"><b><%= adoRS("HomePhone")%></b></td>
<%
				else
%>
				<td bgcolor="#cccccc">&nbsp;</td>
				<td bgcolor="#cccccc" colspan="2">&nbsp;</td>
<%
				end if
%>
			  <td bgcolor="#cccccc"><strong>Gender:</strong></td>
			  <td style='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><b><%= adoRS("Gender")%></b></td>
			</tr>
			<tr borderColorDark="#fcfcfc" borderColorlight="#cccccc">
<%
				if not m_bDependent then
%>
				<td bgcolor="#f0f0f0"><strong>Work Phone:</strong></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0" colspan="2"><b><%= adoRS("WorkPhone")%></b></td>
<%
				else
%>
				<td bgcolor="#f0f0f0"><strong>Dependent SSN:</strong></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0" colspan="2"><b><%= sShownSSN%></b></td>
<%
				end if
				if not m_bDependent then
%>
				<td bgcolor="#f0f0f0"><strong>Marital Status:</strong></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><b><%= adoRS("MaritalStatus")%></b></td>
<%
				else
%>
				<td bgcolor="#f0f0f0"><strong>Relationship:</strong></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><b><%= adoRS("Relationship")%></b></td>
<%
				end if
%>
			</tr>
<%
			if not m_bDependent then
%>
			<tr borderColorDark="#f0f0f0" borderColorlight="#999999">
				<td bgcolor="#cccccc"><b>Fund:</b></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc" colspan="2">&nbsp;<b><%= adoRS("Fund")%></b></td>
				<td bgcolor="#cccccc"><strong>Retirement Date:</strong></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><b><%= adoRS("RetireDate")%></b></td>
			</tr>
			<tr borderColorDark="#fcfcfc" borderColorlight="#cccccc">
				<td bgcolor="#f0f0f0"><b>Locked File:</b></td>
        <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><% if adoRS("locktype") <> "" then %><img SRC="images/lock1.gif" width="32" height="32" alt="LOCKED FILE!!">&nbsp;&nbsp;<b>Yes</b><% else %><b>No</b><% end if %></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><% if adoRS("locktype") <> "" then %><b><%=adoRS("locktype")%></b><% else %>&nbsp;<% end if %></td>
				<td bgcolor="#f0f0f0"><strong>Date of Death:</strong></td>
				<td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%= adoRS("DeathDate")%></b></td>
			</tr>
<%
			end if
%>
		</table>
    <br />
		<hr size="3" noshade>
<%
		adoRS.Close
'**************
' Dependents
'**************
		if not m_bDependent then
%>
			<br />	
			<font size="3" face="arial, helvetica"><center><b>Dependents</b></center></font>
			<table border="1" cellpadding="2" cellspacing="2" width="100%" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
				<tr align="center" borderColorDark="#f0f0f0" borderColorlight="#999999">
				    <td bgcolor="#cccccc"><strong>Dep. #</strong></td>
				    <td bgcolor="#cccccc" colspan="2"><strong>Name</strong></td>
				    <td bgcolor="#cccccc"><strong>SSN</strong></td>
				    <td bgcolor="#cccccc"><strong>Gender</strong></td>
				    <td bgcolor="#cccccc"><strong>Relationship</strong></td>
				    <td bgcolor="#cccccc"><strong>Birthdate</strong></td>
            <td bgcolor="#cccccc" colspan="2"><strong>Locked</strong></td>
            <td bgcolor="#cccccc"><strong>Admin Lock</strong></td>
          </tr>
<%
				sSQL = "select SSN = n.DepSSN, n.DepNumber, n.LastName + ', ' + "
				sSQL = sSQL & " n.FirstName + ' ' + n.MiddleName 'Name', n.BirthDate,"
				sSQL = sSQL & " n.Gender, "
				if not m_bUseEAS then
					sSQL = sSQL & " r.Description, n.SupervisorAccessOnly, n.locktype From Names n LEFT JOIN RelationDescription r "
				else
					sSQL = sSQL & " r.Description from EASNames n LEFT JOIN RelationDescription r "
				end if
				sSQL = sSQL & " on (r.RelationCode = n.RelationCode) WHERE n.SSN = " & m_iSSN  
				sSQL = sSQL & " AND n.DepNumber !=0 ORDER BY n.DepNumber"			
        adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
        
	'		SQLQuery = "WEBDependentDemoGraphics " + SSN
	'		Set rs = Conn.Execute(SQLQuery)
				if not adoRS.EOF then
					Do While Not adoRS.EOF
%>
						<tr align="center" borderColorDark="#fcfcfc" borderColorlight="#cccccc">
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><b><%=adoRS("DepNumber")%></b></td>
                <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><% if adoRS("locktype") <> "" then %><img SRC="images/lock1.gif" width="32" height="32" alt="LOCKED FILE!!"><% else %>&nbsp;<% end if %></td>
                <td align="left" style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">
<%
						    if adoRS("locktype") <> "" then 'or adoRS("SupervisorAccessOnly") then
                  if Manager = "true" or Supervisor = "true" or Auditor = "true" then
                    sTemp = "<A href='PersonDetails.asp?SSN=" & m_iSSN & "&DepNum=" & adoRS("DepNumber") & "&locktype=" & adoRS("locktype") & "&Admin=" & adoRS("SupervisorAccessOnly")
                    sTemp = sTemp & "' onclick='WorkingStatus()'>"
                    sTemp = sTemp & "<b><font color='red'>" & adoRS("Name") & "</font></b></A></td>"
                  else
                    if adoRS("SupervisorAccessOnly") then
                      sTemp = "<A href='LockedFile.asp?SSN=" & m_iSSN & "&locktype=" & adoRS("locktype") & "&Admin=" & adoRS("SupervisorAccessOnly") & "&DepNum=" & adoRS("DepNumber")
                      sTemp = sTemp & "' onclick='WorkingStatus()'>"
                      sTemp = sTemp & "<b>" & adoRS("Name") & "</b></A></td>"
                    else
                      sTemp = "<A href='PersonLockDetail.asp?SSN=" & m_iSSN & "&locktype=" & adoRS("locktype") & "&Admin=" & adoRS("SupervisorAccessOnly") & "&DepNum=" & adoRS("DepNumber")
                      sTemp = sTemp & "' onclick='WorkingStatus()'>"
                      sTemp = sTemp & "<b>" & adoRS("Name") & "</b></A></td>"
                    end if
                  end if
                else
                  if adoRS("SupervisorAccessOnly") then
                    if Manager = "true" or Supervisor = "true" or Auditor = "true" then
                      sTemp = "<A href='PersonDetails.asp?SSN=" & m_iSSN & "&DepNum=" & adoRS("DepNumber") & "&Admin=" & adoRS("SupervisorAccessOnly")
                      sTemp = sTemp & "' onclick='WorkingStatus()'>"
                      sTemp = sTemp & "<b>" & adoRS("Name") & "</b></A></td>"
                    else
                      sTemp = "<A href='LockedFile.asp?SSN=" & m_iSSN & "&locktype=" & adoRS("locktype") & "&Admin=" & adoRS("SupervisorAccessOnly") & "&DepNum=" & adoRS("DepNumber")
                      sTemp = sTemp & "' onclick='WorkingStatus()'>"
                      sTemp = sTemp & "<b>" & adoRS("Name") & "</b></A></td>"
                    end if
                  else
                    sTemp = "<A href='PersonDetails.asp?SSN=" & m_iSSN & "&DepNum=" & adoRS("DepNumber")
                    sTemp = sTemp & "' onclick='WorkingStatus()'>"
                    sTemp = sTemp & "<b>" & adoRS("Name") & "</b></A></td>"
                  end if
                end if
                response.write sTemp
%>
                <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><b>&nbsp;
<% 
                if len(adoRS("SSN")) = 8 then
                  sTemp = "0" & adoRS("SSN")
                elseif len(adoRS("SSN")) = 7 then
                  sTemp = "00" & adoRS("SSN")
                else
				          sTemp = adoRS("SSN")
                end if
                response.write sTemp & "</b></td>"
%>                
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><b><%=adoRS("Gender")%></b></td>
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><b><%=adoRS("Description")%></b></td>
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><b><%=adoRS("BirthDate")%></b></td>
                <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><% if adoRS("locktype") <> "" then %><b>Yes</b><% else %><b>No</b><% end if %></td>
                <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><% if adoRS("locktype") <> "" then %><b><%=adoRS("locktype")%></b><% else %>&nbsp;<% end if %></td>
                <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><% if adoRS("SupervisorAccessOnly") then %><b>Yes</b><% else %><b>No</b><% end if %></td>
              </tr>
            
<% 
						adoRS.MoveNext
					Loop 
				else
%>
					<tr borderColorDark="#fcfcfc" borderColorlight="#cccccc" align="center">
					    <td colspan="10" bgcolor="#f0f0f0"><font color="red" face="verdana, arial, helvetica" size="2"><b>No dependents found.</b></font></td>
					</tr>
<%
				end if  
%>
      
          <tr>
            <td colspan="10" align="right"><img SRC="images/bluebar2.gif" onclick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif" onclick="history.go(+1)" border="0"></td>
          </tr>
      
      </table>
<%
			adoRS.Close
		end if
'****************
' Enrollments
'****************
%>
		<br />	
		<font size="3" face="arial, helvetica"><center><b>Fund Enrollment</b></center></font>
		<table border="1" cellpadding="2" cellspacing="2" width="100%" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr align="center" borderColorDark="#f0f0f0" borderColorlight="#999999">
				<td bgcolor="#cccccc"><strong>Fund</strong></td>
				<td bgcolor="#cccccc"><strong>Fund Name</strong></td>
				<td bgcolor="#cccccc"><strong>Date Enrolled</strong></td>
			</tr>
<%
			'sSQL = "select e.Enrolledfund, CONVERT(VARCHAR(12),e.dateenrolled,101) EDate, "
			sSQL = "select e.Enrolledfund, CONVERT(VARCHAR(12),e.dateenrolled,101) EDate, "
      sSQL = sSQL & " f.name1 FundName from enrolled e"
			'sSQL = sSQL & " inner join funds f on e.enrolledfund = f.enrolledfund"
      sSQL = sSQL & " inner join funds f on e.enrolledfund = f.fund"
			sSQL = sSQL & " WHERE e.SSN = " & m_iSSN  & " AND e.DepNumber = " & m_iDepno		
			sSQL = sSQL & " order by FundName, EDate"					
      
			adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
      
			if not adoRS.EOF then
				Do While Not adoRS.EOF
%>
					<tr align="center" borderColorDark="#fcfcfc" borderColorlight="#cccccc">
					    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><b><%=adoRS("Enrolledfund")%></b></td>
					    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><b><%=adoRS("FundName")%></b></td>
					    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0"><b><%=adoRS("EDate")%></b></td>
					</tr>
<% 
					adoRS.MoveNext
				Loop
			else
%>
				<tr borderColorDark="#fcfcfc" borderColorlight="#cccccc" align="center">
				    <td colspan="3" bgcolor="#f0f0f0"><font color="red" face="verdana, arial, helvetica" size="2"><b>Not enrolled in any funds.</b></font></td>
				</tr>
<%
			end if  
%>
		</table>
<%
		adoRS.Close
'***************
' Eligibility
'***************
    if not m_bDependent then
      
%>		
			<br />	
			<font size="3" face="arial, helvetica"><center><b>Eligibility</b></center></font>
<%
			sSQL = "select shortdescription from monthdescription order by month ASC"
			adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
			for i = 1 to 12
				avMonthName(i) = adoRS("shortdescription")
				adoRS.MoveNext
			next
			adoRS.Close
			sSQL = "select count(distinct easfund+' '+easplan) FUND_PLAN_COUNT from elig where ssn=" & m_iSSN
			adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
			iNumFundPlans = adoRS("FUND_PLAN_COUNT")
			adoRS.Close

			sSQL = "select distinct easfund+' '+easplan FUND_PLAN from elig where ssn=" & m_iSSN & " order by FUND_PLAN"
			adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic

			If adoRS.EOF Then
			    Response.Write "<TABLE bgcolor='white' bordercolor='white'<tr borderColorDark='#FCFCFC' borderColorlight='#CCCCCC' style='font:9pt verdana, arial, helvetica,sans-serif'><td bgcolor='#f0f0f0' colspan='6'><font color='red'><b>No Eligibility Information found.</b></font></td></tr>"
				adoRS.Close
			Else
			    ReDim avFundPlan(iNumFundPlans)
			    For i = 1 To iNumFundPlans
			        avFundPlan(i) = adoRS("FUND_PLAN")
			        adoRS.MoveNext
			    Next 
				adoRS.Close

				sSQL = "select max(year) MaxYear, min(year) MinYear from elig where ssn=" & m_iSSN
				adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
				iNewestYear = adoRS("MaxYear")
				iOldestYear = adoRS("MinYear")
				iNumYears = iNewestYear - iOldestYear + 1
				adoRS.Close

				sSQL = "select max(month) month from elig where ssn=" & m_iSSN & " and year=" & iNewestYear
				adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
				iNewestMonth = adoRS("month")
				adoRS.Close

				sSQL = "select min(month) month from elig where ssn=" & m_iSSN & " and year=" & iOldestYear
				adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
				iOldestMonth = adoRS("month")
				adoRS.Close

				If iOldestMonth > iNewestMonth Then
				    iNumYears = iNumYears - 1
				End If

				ReDim avEligData(iNumFundPlans,iNumYears * 12)
        ReDim Carrier(iNumFundPlans,iNumYears * 12)
        ReDim Reason(iNumFundPlans,iNumYears * 12)

				'sSQL = "select e.easfund, e.easplan, e.Year, e.Month, s.NewChoice, e.StatusReason + ' ' + e.EASRate Data from elig e left join EASPriorCBS s on (e.SSN=s.SSN and s.EASFund = e.EASFund and s.EASPlan = e.EASPlan and s.EffDate <= convert( datetime, convert( varchar(5), e.Month ) + '-01-' + convert( varchar(5), e.Year ) ) and s.TranDate = ( Select max( trandate ) From EASPriorCBS Where SSN = s.SSN and EASFund = s.EASFund and EASPlan = s.EASPlan and EffDate <= convert( datetime, convert( varchar(5), e.Month ) + '-01-' + convert( varchar(5), e.Year ) ) )) where e.ssn=" & m_iSSN  & " order by e.easfund,e.easplan,e.year DESC, e.month DESC"
        sSQL = "ListEligibility " & m_iSSN
				'response.write sSQL
        Set adoRS = adoConn.execute(sSQL)
        'adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
    
				Do While Not adoRS.EOF
				    iCurrRow = 0
				    For i = 1 To UBound(avFundPlan)
				        If avFundPlan(i) = adoRS("Fund") & " " & adoRS("PlanCode") Then
				            Exit For
				        End If
				    Next 
				    If i > UBound(avFundPlan) Then
				        MsgBox "error finding fund_plan"
				        Exit Do
				    End If
				    iCurrRow = i
				    iCurrCol = ((iNewestYear - adoRS("Year")) * 12) + (iNewestMonth - adoRS("Month") + 1)
				    avEligData(iCurrRow, iCurrCol) = adoRS("RateCode") 
            Carrier(iCurrRow, iCurrCol) = adoRS("Carrier")
            Reason(iCurrRow, iCurrCol) = adoRS("StatusReason")
				    adoRS.MoveNext
				Loop
				adoRS.Close

				ReDim avRowData(UBound(avEligData, 1))
				
				For i = 1 To UBound(avEligData, 1)
					avRowData(i) = "<td ALIGN=CENTER bgcolor='#CCCCCC'>" & mid(avFundPlan(i),1,instr(avFundPlan(i)," ")-1) & "</td>" & "<td ALIGN=CENTER bgcolor='#CCCCCC'>" & mid(avFundPlan(i),instr(avFundPlan(i)," ")+1,len(avFundPlan(i))) & "</td>"
				next
				sSectionTitle = "<td ALIGN=CENTER bgcolor='#CCCCCC'>Fund</td><td ALIGN=CENTER bgcolor='#CCCCCC'>Plan</td>"
				iCurrYear = iNewestYear
				if iNewestMonth = 12 then
					iCurrYear = iCurrYear + 1
				end if
				Response.Write "<TABLE border=1 cellpadding=2 cellspacing=2 width=100% bgcolor='white' bordercolor='white' style='font:9pt verdana, arial, helvetica,sans-serif'>"
				For j = 1 To UBound(avEligData, 2)
					if (((11 - ((j - 1) mod 12) + iNewestMonth) mod 12) + 1) = 12 then
						iCurrYear =  iCurrYear - 1
					end if 
					sSectionTitle = sSectionTitle & "<td ALIGN=CENTER bgcolor='#f0f0f0'>" & avMonthName(((11 - ((j - 1) mod 12) + iNewestMonth) mod 12) + 1) & " " & right(iCurrYear,2) & "</td>"
				    For i = 1 To UBound(avEligData, 1)
						if trim(avEligData(i, j)) <> "" then
              'avEligData needs to replace Reason in Function ShowPlanDesc for it to work correctly 
							sTemp = "<td><TABLE cellpadding=1 cellspacing=0><TR borderColorDark='#FCFCFC' borderColorlight='#CCCCCC' style='font:9pt verdana, arial, helvetica,sans-serif'><td ALIGN=CENTER bgcolor='#f0f0f0'>" & avEligData(i, j) & " " & Reason(i, j) & "<br />" & Carrier(i, j) & "</td></TR><TR><td ALIGN=CENTER bgcolor='#f0f0f0'>"
							sTemp = sTemp & "<IMG SRC='images/ACSPlan.gif' onclick='javascript:ShowPlanDesc(" & chr(34) & mid(avFundPlan(i),1,instr(avFundPlan(i)," ")-1) & chr(34) & "," & chr(34) & mid(avFundPlan(i),instr(avFundPlan(i)," ")+1,len(avFundPlan(i))) & chr(34) & "," & chr(34) & right(Reason(i, j),2) & chr(34) & ")'>"
							sTemp = sTemp & "</td></TR></TABLE></td>"
						else
							sTemp = "<td bgcolor='#f0f0f0'></td>"
						end if
				        avRowData(i) = avRowData(i) & sTemp
				    Next 
				    If (j Mod 12) = 0 Then
						Response.Write "<TR borderColorDark='#FCFCFC' borderColorlight='#CCCCCC'>" & sSectionTitle &  "</TR>"
				        For i = 1 To UBound(avEligData, 1)
							Response.Write "<TR borderColorDark='#FCFCFC' borderColorlight='#CCCCCC'>" &  avRowData(i) & "</TR>"
							avRowData(i) = "<td ALIGN=CENTER bgcolor='#CCCCCC'>" & mid(avFundPlan(i),1,instr(avFundPlan(i)," ")-1) & "</td>" & "<td ALIGN=CENTER bgcolor='#CCCCCC'>" & mid(avFundPlan(i),instr(avFundPlan(i)," ")+1,len(avFundPlan(i))) & "</td>"
							sSectionTitle = "<td ALIGN=CENTER bgcolor='#CCCCCC'>Fund</td><td ALIGN=CENTER bgcolor='#CCCCCC'>Plan</td>"
				        Next
				        Response.Write "</TABLE><br /><TABLE border=1 cellpadding=2 cellspacing=2 width=100% bgcolor='white' bordercolor='white' style='font:9pt verdana, arial, helvetica,sans-serif'>"
				    End If
				Next 
			End If
%>
			</table>
<%
		end if		
'***********
' Notes
'***********
%>
		<br />
		<font size="3" face="arial, helvetica"><center><b>Notes</b></center></font>
		<table border="1" cellpadding="2" cellspacing="2" width="100%" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr align="center" borderColorDark="#f0f0f0" borderColorlight="#999999">
			    <td bgcolor="#cccccc"><strong>Date</strong></td>
			    <td bgcolor="#cccccc"><strong>Note Type</strong></td>
			    <td bgcolor="#cccccc"><strong>Description</strong></td>
          <!--- <td bgcolor="#cccccc"><strong>Priority</strong></td> --->
			</tr>
<%
	  if not m_bDependent then
       sSQL = "select a = n.DateEntered, b = n.NoteType, c = n.NoteText, d = COALESCE( o.Priority, 0 ), e = n.UpdatedBy, n.NoteID"
       sSQL = sSQL & " From NameNotes n Left Join NotePriorityOrder o on ( o.NoteType = n.NoteType )"
       sSQL = sSQL & " WHERE SSN = " & m_iSSN & " AND DepNumber in (99,0)"
       sSQL = sSQL & " order by d DESC, a DESC" 
      'sSQL = "select DateEntered, CONVERT(VARCHAR(12),DateEntered,101) DisplayDate, NoteType, NoteText"
			'sSQL = sSQL & " FROM NameNotes"
			'sSQL = sSQL & " WHERE SSN = " & m_iSSN & " AND DepNumber in (99,0) "
			'sSQL = sSQL & " order by DateEntered DESC"
    else					
       sSQL = "select a = n.DateEntered, b = n.NoteType, c = n.NoteText, d = COALESCE( o.Priority, 0 ), e = n.UpdatedBy, n.NoteID"
       sSQL = sSQL & " From NameNotes n Left Join NotePriorityOrder o on ( o.NoteType = n.NoteType )"
       sSQL = sSQL & " WHERE SSN = " & m_iSSN & " AND (DepNumber = 99 or DepNumber=" & m_iDepNo & ")"
       sSQL = sSQL & " order by d DESC, a DESC" 
      'sSQL = "select DateEntered, CONVERT(VARCHAR(12),DateEntered,101) DisplayDate, NoteType, NoteText"
			'sSQL = sSQL & " FROM NameNotes"
			'sSQL = sSQL & " WHERE SSN = " & m_iSSN & " AND (DepNumber = 99 or DepNumber=" & m_iDepNo & ") "
			'sSQL = sSQL & " order by DateEntered DESC"
    end if
			
      adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
'		SQLQuery = "WEBNoteHistory " + SSN
'		Set rs = Conn.Execute(SQLQuery)
			if not adoRS.EOF then
				Do While Not adoRS.EOF
%>
					<tr align="center" borderColorDark="#fcfcfc" borderColorlight="#cccccc">
					    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%= FORMATDATETIME(adoRS("a"),vbshortdate)%></b></td>
					    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%= adoRS("b")%></b></td>
					    <td align="left" style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%= adoRS("c")%></b></td>
              <!--- <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%= adoRS("d")%></b></td> --->
					</tr>
<% 
					adoRS.MoveNext
				Loop 
			else
%>
				<tr borderColorDark="#fcfcfc" borderColorlight="#cccccc" align="center">
				    <td colspan="4" bgcolor="#f0f0f0"><font color="red" face="verdana, arial, helvetica" size="2"><b>No notes found.</b></font></td>
				</tr>
<%
			end if  
%>
		    <tr>
          <td colspan="6" align="right"><img SRC="images/bluebar2.gif" onclick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif" onclick="history.go(+1)" border="0"></td>
        </tr>
    
    </table>
<%
		adoRS.Close

'**********
' COB
'**********
%>
		<br />
		<font size="3" face="arial, helvetica"><center><b>Coordination of Benefits Information</b></center></font>
		<table width="100%" border="1" cellpadding="2" cellspacing="2" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr align="center" borderColorDark="#f0f0f0" borderColorlight="#999999">
			    <td bgcolor="#cccccc"><strong>OIC Participant</strong></td>
			    <td bgcolor="#cccccc"><strong>Birthdate</strong></td>
			    <td bgcolor="#cccccc"><strong>Dep. #</strong></td>
          <td bgcolor="#cccccc"><strong>Dependent Name</strong></td>
			    <td bgcolor="#cccccc"><strong>NWA/COB Order</strong></td>
<!--			    <td><strong>COB Key</strong></td>-->
			    <td bgcolor="#cccccc"><strong>Carr. Code</strong></td>
			    <td bgcolor="#cccccc"><strong>Cov. Type</strong></td>
			    <td bgcolor="#cccccc"><strong>Eff. Date</strong></td>
			    <td bgcolor="#cccccc"><strong>Term. Date</strong></td>
			</tr>
<%
			sSQL = "select OtherFirstName + ' ' + OtherInitial + ' ' + OtherLastName 'Name'," 
			sSQL = sSQL & " c.BirthDate, c.DepNumber, c.Coveragemedical, c.Coveragedental, c.CoverageOrtho, c.CoverageVision, c.CoveragePrescription, c.Coveragechiro, c.CoverageOther,  EffectiveDate, c.TermDate, NWAOrder,"
			sSQL = sSQL & " COBKey, CarrierCode , n.FirstName, n.LastName FROM NEWCOB c LEFT JOIN Names n ON c.SSN = n.SSN and c.depnumber = n.depnumber"
			sSQL = sSQL & " WHERE c.SSN = " & m_iSSN 
			sSQL = sSQL & " order by c.DepNumber"
      adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
			if not adoRS.EOF then
				Do While Not adoRS.EOF
%>
					<tr borderColorDark="#fcfcfc" borderColorlight="#cccccc">
						  <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%=adoRS("Name")%></b></td>
					    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%= adoRS("Birthdate")%></b></td>
					    <td bgcolor="#f0f0f0">&nbsp;<a href="COBDetails.asp?COBKey=<%= adoRS("COBKey")%>&amp;DepNumber=<%= adoRS("DepNumber") %>&amp;dFname=<%= adoRS("FirstName") %>&amp;dLname=<%= adoRS("LastName") %>&amp;BirthDay=<%= m_sBirthDate %>&amp;Particname=<%= m_sParticName%>&amp;UseEAS=<%= m_bUseEAS%>" onclick="WorkingStatus()"><b><%= adoRS("DepNumber")%></b></a></td>
              <td bgcolor="#f0f0f0">&nbsp;<a href="COBDetails.asp?COBKey=<%= adoRS("COBKey")%>&SSN=<%= m_iSSN %>&amp;DepNumber=<%= adoRS("DepNumber") %>&amp;dFname=<%= adoRS("FirstName") %>&amp;dLname=<%= adoRS("LastName") %>&amp;BirthDay=<%= m_sBirthDate %>&amp;Particname=<%= m_sParticName%>&amp;UseEAS=<%= m_bUseEAS%>" onclick="WorkingStatus()"><b><%= adoRS("FirstName")%>&nbsp;<%= adoRS("LastName")%></b></a></td>
					    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><% if adoRS("NWAOrder") = 0 then %>Not On File<% else %><%= adoRS("NWAOrder")%><% end if %></b></td>
<!--					    <td><%= adoRS("COBKey")%></td>-->
					    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%= adoRS("CarrierCode")%></b></td>
					    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<% if adoRS("Coveragemedical") = "true" then %><b>M </b><% end if %><% if adoRS("Coveragedental") = "true" then %><b>D </b><% end if %><% if adoRS("CoverageOrtho") = "true" then %><b>O </b><% end if %><% if adoRS("CoverageVision") = "true" then %><b>V </b><% end if %><% if adoRS("CoveragePrescription") = "true" then %><b>P </b><% end if %><% if adoRS("Coveragechiro") = "true" then %><b>C </b><% end if %><% if adoRS("CoverageOther") = "true" then %><b>Other</b><% end if %></td>
					    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%= adoRS("EffectiveDate")%></b></td>
					    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%= adoRS("TermDate")%></b></td>
					</tr>
<% 
					adoRS.MoveNext
          
				Loop
         
			else
%>
				<tr borderColorDark="#fcfcfc" borderColorlight="#cccccc" align="center">
				    <td colspan="9" bgcolor="#f0f0f0"><font color="red" face="verdana, arial, helvetica" size="2"><b>No COB coverage found.</b></font></td>
				</tr>
<%
			end if  
			adoRS.Close
      
%>
		</table>
<%
'************
' Schools
'************
		if m_bDependent then
%>		
			<br />	
			<font size="3" face="arial, helvetica"><center><b>School Attendance</b></center></font>
<%
			sSQL = "select count(distinct schoolname) SCHOOL_COUNT from school where ssn="
			sSQL = sSQL & m_iSSN & " and depnumber=" & m_iDepNo 
			adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
			iNumSchools = adoRS("SCHOOL_COUNT")
			adoRS.Close
			If iNumSchools = 0 Then
			    Response.Write "<TABLE style='font:9pt verdana, arial, helvetica,sans-serif'><td><font color='red'><b>No School Attendance Information found.</b></font></td>"
			Else
				sSQL = "select shortdescription from monthdescription order by month ASC"
				adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
				for i = 1 to 12
					avMonthName(i) = adoRS("shortdescription")
					adoRS.MoveNext
				next
				adoRS.Close
				reDim avSchools(iNumSchools)
				sSQL = "select DISTINCT SchoolName from School where ssn=" & m_iSSN & " and depnumber=" & m_iDepNo & "order by SchoolName"
				adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
				'**Changed i = 1 to i = 0 10/28/99
        i = 0
				Do While Not adoRS.EOF
					avSchools(i) = adoRS("SchoolName")
					i = i + 1
					adoRS.MoveNext
				loop
				adoRS.Close
				sSQL = "select datepart(year,min(FromDate)) START_YEAR,"
				sSQL = sSQL & " datepart(year,max(ThruDate)) END_YEAR,"
				sSQL = sSQL & " datepart(month,min(FromDate)) START_MONTH,"
				sSQL = sSQL & " datepart(month,max(ThruDate)) END_MONTH"
				sSQL = sSQL & " from school where ssn=" & m_iSSN & " and depnumber=" & m_iDepNo
				adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
				iNewestYear = adoRS("END_YEAR")
				iOldestYear = adoRS("START_YEAR")
				iNewestMonth = adoRS("END_MONTH")
				iOldestMonth = adoRS("START_MONTH")
				iNumYears = iNewestYear - iOldestYear + 1
				adoRS.Close

				If iOldestMonth > iNewestMonth Then
				    iNumYears = iNumYears - 1
				End If

				ReDim avSchoolData(iNumSchools,iNumYears * 12)

				sSQL = "select FullTime, datepart(year,FromDate) START_YEAR,"
				sSQL = sSQL & " datepart(year,ThruDate) END_YEAR,"
				sSQL = sSQL & " datepart(month,FromDate) START_MONTH,"
				sSQL = sSQL & " datepart(month,ThruDate) END_MONTH, SchoolName"
				sSQL = sSQL & " from school where ssn=" & m_iSSN & " and depnumber=" & m_iDepNo
				sSQL = sSQL & " order by SchoolName, FromDate"
				adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
				iCurrRow = 1
				sCurrSchool = adoRS("SchoolName")
				Do While Not adoRS.EOF
					if adoRS("FullTime") then
						sAttend = "FT"
					else
						sAttend = "PT"
					end if	
				    iEndCol = ((iNewestYear - adoRS("START_YEAR")) * 12) + (iNewestMonth - adoRS("START_MONTH") + 1)
				    iStartCol = ((iNewestYear - adoRS("END_YEAR")) * 12) + (iNewestMonth - adoRS("END_MONTH") + 1)
				    for i = iStartCol to iEndCol
						avSchoolData(iCurrRow, i) = sAttend
					next
				    adoRS.MoveNext
				    if not adoRS.EOF then
						if sCurrSchool <> adoRS("SchoolName") then
							iCurrRow = iCurrRow +1
							sCurrSchool = adoRS("SchoolName")
						end if
					end if
				Loop
				adoRS.Close
					
				ReDim avRowData(UBound(avSchools, 1))
						    
				For i = 1 To UBound(avSchools, 1)
					avRowData(i) = "<td>" & avSchools(i) & "</td>"
				next
				sSectionTitle = "<td ALIGN=CENTER>School</td>"
				iCurrYear = iNewestYear
				if iNewestMonth = 12 then
					iCurrYear = iCurrYear + 1
				end if
				Response.Write "<table border=1 cellpadding=2 cellspacing=2 width=100% bgcolor='white' bordercolor='white' style='font:9pt verdana, arial, helvetica,sans-serif'>"
				For j = 1 To UBound(avSchoolData, 2)
					if (((11 - ((j - 1) mod 12) + iNewestMonth) mod 12) + 1) = 12 then
						iCurrYear =  iCurrYear - 1
					end if 
					sSectionTitle = sSectionTitle & "<td NOWRAP>" & avMonthName(((11 - ((j - 1) mod 12) + iNewestMonth) mod 12) + 1) & " " & right(iCurrYear,2) & "</td>"
				    For i = 1 To UBound(avSchoolData, 1)
				        avRowData(i) = avRowData(i) & "<td NOWRAP>" & avSchoolData(i, j) & "</td>"
				    Next 
				    If (j Mod 12) = 0 Then
						Response.Write "<TR>" & sSectionTitle &  "</TR>"
				        For i = 1 To UBound(avSchoolData, 1)
							Response.Write "<TR>" &  avRowData(i) & "</TR>"
							avRowData(i) = "<td>" & avSchools(i) & "</td>"
							sSectionTitle = "<td ALIGN=CENTER>School</td>"
				        Next
				        Response.Write "</TABLE><br /><TABLE border=1 cellpadding=2 cellspacing=2 width=100% >"
				    End If
				Next 
			End If
%>
			</table>
<%
		end if
'**************
' Employment
'**************
		if not m_bDependent then
			sSQL = "select distinct t.local, t.employer, e.Name, t.Year from TransmittalDetail t " 
			sSQL = sSQL & " left join employers e on (e.Employer = t.Employer) WHERE"
			sSQL = sSQL & " t.ssn = " & m_iSSN & " order by  t.year DESC, t.local"
%>
			<br />	
			<p><font size="3" face="arial, helvetica"><center><b>Employment History</b></center></font>
			<table border="1" cellpadding="2" cellspacing="2" width="100%" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
				<tr align="center" borderColorDark="#f0f0f0" borderColorlight="#999999">
					<td bgcolor="#cccccc"><strong>Employer Name</strong></td>
					<td bgcolor="#cccccc"><strong>Local</strong></td>
					<td bgcolor="#cccccc"><strong>Year</strong></td>
				</tr>
<%
'***********************************************************************************************************************************************
' See comments regarding using the ADO Command object below, relative to this Close
' statement.  Per use of just an ADO recordset, this line is currenly un-commented
'
'adoRS.Close
''
'' When moved to seasql02, using an SP failed with an ADO 'login failed' error -- perhaps
'' a permission problem, although the SP settings seemed normal.  Converted to use just
'' the ADO recordset (setting SQL here), and things worked...
''
'				adoCmd.CommandText = "dbo.listemployment"
'				adoCmd.CommandType = 4 'adCmdStoredProc
'				Set adoCmd.ActiveConnection = adoConn
'				'
'				' For CreateParameter method, parameters are:
'				'	[Name As String], [Type As DataTypeEnum = adEmpty], [Direction As ParameterDirectionEnum = adParamInput], [Size As Long], [value])
'				' and the value for the SP are:
'				'	user, adVarChar, adParamInput, 50 (field is varChar 50), username (Session("User"))
'				'
'				Set adoParam = adoCmd.CreateParameter("@SSN", 3, 1, 4, m_iSSN)
'				adoCmd.Parameters.Append adoParam
''
'' For some indeterminate reason, if adoRS.Close is performed earlier in this script
'' (for example, just in the commented-out line just before the Employment section
'' above), later recordset operations will at some point produce errors -- either
'' 'TDS Proctol Error' or 'Communication Link Failure'.  However, it seems to occur
'' after several successful accesses (that is, Benefit Periods succeeds, but then
'' Eligibility fails.  This error occurred when this section was the only one using
'' an ADO Command object, so that might have some impact.
''
'				adoRS.Open adoCmd
'***********************************************************************************************************************************************
				adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
				if not adoRS.EOF then
					Do While Not adoRS.EOF
%>
						<tr borderColorDark="#fcfcfc" borderColorlight="#cccccc">
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%=adoRS("name")%></b></td>
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%=adoRS("local")%></b></td>
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%=adoRS("year")%></b></td>
						</tr>
<% 
						adoRS.MoveNext
					Loop
				else
%>
					<tr borderColorDark="#fcfcfc" borderColorlight="#cccccc" align="center">
					    <td colspan="3" bgcolor="#f0f0f0"><font color="red" face="verdana, arial, helvetica" size="2"><b>No employment information found.</b></font></td>
					</tr>
<%
				end if  
%>
			</table>
<%
			adoRS.Close
		end if
'*****************
' Benefit Periods
'*****************
		if not m_bDependent then
%>
			<br />	
			<font size="3" face="arial, helvetica"><center><b>Benefit Periods</b></center></font>
			<table border="1" cellpadding="2" cellspacing="2" width="100%" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
				<tr align="center" borderColorDark="#f0f0f0" borderColorlight="#999999">
					<td bgcolor="#cccccc"><strong>Dep. #</strong></td>
					<td bgcolor="#cccccc"><strong>From Date</strong></td>
					<td bgcolor="#cccccc"><strong>Thru Date</strong></td>
					<td bgcolor="#cccccc"><strong>Diag. Code</strong></td>
					<td bgcolor="#cccccc"><strong>Diagnosis</strong></td>
					<td bgcolor="#cccccc"><strong>Notes</strong></td>
				</tr>
<%
				'sSQL = "select depNumber, FromDate, ThruDate, b.Diagnosis, Description,"
				'sSQL = sSQL & " Notes FROM BenefitPeriod b LEFT JOIN Diagnosis d on"
				'sSQL = sSQL & " (b.Diagnosis = d.DiagnosisCode) WHERE b.SSN = " & m_iSSN			
				'sSQL = sSQL & " order by FromDate"
        sSQL = "PB_BenefitPeriod " & m_iSSN					

				Set adoRS = adoConn.execute(sSQL)
        'adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
				if not adoRS.EOF then
					Do While Not adoRS.EOF
%>
						<tr borderColorDark="#fcfcfc" borderColorlight="#cccccc">
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%=adoRS("depNumber")%></b></td>
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%=adoRS("FromDate")%></b></td>
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%=adoRS("ThruDate")%></b></td>
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%=adoRS("Diagnosis")%></b></td>
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%=adoRS("Description")%></b></td>
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%=adoRS("Notes")%></b></td>
						</tr>
<% 
						adoRS.MoveNext
					Loop
				else
%>
					<tr borderColorDark="#fcfcfc" borderColorlight="#cccccc" align="center">
					    <td colspan="6" bgcolor="#f0f0f0"><font color="red" face="verdana, arial, helvetica" size="2"><b>No Benefit Periods found.</b></font></td>
					</tr>
<%
				end if  
%>
			</table>
<% 
			adoRS.Close
		end if
'*******************************************************************************************************
' Address History
'****************
' Note: Nominally Address changes would be maintained in HomeChanges.  However, that table is
' not currently completely implemented, and so Homes is used.  Additionally, the field 
' UpdatedOn might be expected to be the controlling value for most recent address, but for
' the vast majority of records it is blank, and when HomeChanges is implemented it may not be
' applicable anyway.  DateCreated is used to sort values.  Note that this does not apply
' when using EAS Data.
'*******************************************************************************************************
		if not m_bDependent and not m_bUseEAS then
%>
			<br />	
			<font size="3" face="arial, helvetica"><center><b>Address History</b></center></font>
			<table border="1" cellpadding="2" cellspacing="2" width="100%" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
				<tr align="center" borderColorDark="#f0f0f0" borderColorlight="#999999">
					<td bgcolor="#cccccc"><strong>Address</strong></td>
					<td bgcolor="#cccccc"><strong>Updated On</strong></td>
					<td bgcolor="#cccccc"><strong>By</strong></td>
				</tr>
<%
				'sSQL = "select UpdatedBy, CONVERT(VARCHAR(12),DateCreated,101) DisplayDate, "
				'sSQL = sSQL & " StreetAddress, City, State, PostalCode FROM Homes"
				'sSQL = sSQL & " WHERE SSN = " & m_iSSN			
				'sSQL = sSQL & " order by DateCreated DESC"
        sSQL = "PB_HomesInfo " & m_iSSN					

				Set adoRS = adoConn.execute(sSQL)
        'adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
'				SQLQuery = "WEBAddressHistory " + SSN
'				Set rs = Conn.Execute(SQLQuery)
				if not adoRS.EOF then
					Do While Not adoRS.EOF
%>
						<tr borderColorDark="#fcfcfc" borderColorlight="#cccccc">
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%=adoRS("StreetAddress") + ", " & adoRS("City") + ", " & adoRS("State") + " " & adoRS("PostalCode")%></b></td>
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%=adoRS("DisplayDate")%></b></td>
						    <td style='color:<%= Session("EmphColor")%>;' bgcolor="#f0f0f0">&nbsp;<b><%=adoRS("UpdatedBy")%></b></td>
						</tr>
<% 
						adoRS.MoveNext
					Loop
				else
%>
					<tr borderColorDark="#fcfcfc" borderColorlight="#cccccc" align="center">
					    <td colspan="3" bgcolor="#f0f0f0"><font color="red" face="verdana, arial, helvetica" size="2"><b>No address changes found.</b></font></td>
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


		<script language="vbscript">			
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
			Dim sResult

				sResult=showModalDialog("ACSInfo.asp?EASFund=" & Fund & "&EASPlan=" & Plan & "&EASRate=" & Rate, "ACS Information", "dialogWidth:500px; dialogHeight:295px; help:no;")
				
			end function

		</script>
<%
	end if
	set adoRS = nothing
	adoConn.Close
	set adoConn = nothing
end if
%>	
</body>
</html>
