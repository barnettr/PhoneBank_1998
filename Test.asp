<%
' DEFINE VARIABLES

Set dataConn = Server.CreateObject("ADODB.Connection")
dataConn.Open "","SQLGuest","SQLGuest"

  Set adoCmd = Server.CreateObject("ADODB.Command")
  adoCmd.ActiveConnection = adoConn
  adoCmd.CommandType = &H0004
  adoCmd.CommandText = "pb_GetProviderInfo"
	
		Set adoParam			= adoCmd.CreateParameter("@P1_TaxID")
		adoParam.Type			= 3
		adoParam.Direction		= &H0001
		adoParam.Size			= 9
		adoParam.Value			= m_iTaxID
		adoCmd.Parameters.Append adoParam
    
    Set adoParam			= adoCmd.CreateParameter("@P2_Associate")
		adoParam.Type			= 3
		adoParam.Direction		= &H0001
		adoParam.Size			= 4
		adoParam.Value			= m_iAssocNum
		adoCmd.Parameters.Append adoParam
    
    Set adoParam			= adoCmd.CreateParameter("@P3_FullName")
		adoParam.Type			= 200
		adoParam.Direction		= &H0001
		adoParam.Size			= 40
		adoParam.Value			= m_sFullName
		adoCmd.Parameters.Append adoParam
    
    Set adoParam			= adoCmd.CreateParameter("@P4_Address")
		adoParam.Type			= 200
		adoParam.Direction		= &H0001
		adoParam.Size			= 30
		adoParam.Value			= m_sAddress
		adoCmd.Parameters.Append adoParam
    
    Set adoParam			= adoCmd.CreateParameter("@P5_City")
		adoParam.Type			= 200
		adoParam.Direction		= &H0001
		adoParam.Size			= 28
		adoParam.Value			= m_sCity
		adoCmd.Parameters.Append adoParam
    
    Set adoParam			= adoCmd.CreateParameter("@P6_State")
		adoParam.Type			= 200
		adoParam.Direction		= &H0001
		adoParam.Size			= 2
		adoParam.Value			= m_sState
		adoCmd.Parameters.Append adoParam
    
    Set adoParam			= adoCmd.CreateParameter("@P7_PostalCode")
		adoParam.Type			= 200
		adoParam.Direction		= &H0001
		adoParam.Size			= 12
		adoParam.Value			= m_sZip
		adoCmd.Parameters.Append adoParam
    
    Set adoParam			= adoCmd.CreateParameter("@P8_ProvType")
		adoParam.Type			= 3
		adoParam.Direction		= &H0001
		adoParam.Size			= 4
		adoParam.Value			= m_iProvType
		adoCmd.Parameters.Append adoParam
    
    Set adoParam			= adoCmd.CreateParameter("@P9_UseAlternate")
		adoParam.Type			= 11
		adoParam.Direction		= &H0001
		adoParam.Size			= 5
		adoParam.Value			= m_bUseAltTaxID
		adoCmd.Parameters.Append adoParam
	
Set adoRS = Server.CreateObject("ADODB.Recordset")
adoRS.Open adoCmd,,3,3

' This code sets norecords variable to 1 if no records were returned.
' If your query was a success, the below code may be deleted.
IF dataRec.state = 1 THEN
	If dataRec.EOF THEN
		norecords = 1
	END IF
END IF

%>

<%
' DEFINE VARIABLES
DIM norecords			' This equals 1 if query returns no data, used to bypass data display

Set dataConn = Server.CreateObject("ADODB.Connection")
dataConn.Open "","SQLGuest","SQLGuest"

Set dataCmd = Server.CreateObject("ADODB.Command")
dataCmd.ActiveConnection = dataConn

Set dataRec = Server.CreateObject("ADODB.Recordset")
dataRec.Open dataCmd,,3,3

' This code sets norecords variable to 1 if no records were returned.
' If your query was a success, the below code may be deleted.
IF dataRec.state = 1 THEN
	If dataRec.EOF THEN
		norecords = 1
	END IF
END IF

%>

	<% IF norecords <> 1 THEN %>
	
<TABLE border="1" align="center">
		<TR bgcolor="White">
	<% dataRec.movefirst %>
	<% For each oField IN dataRec.Fields %>
		<TH align="center"><%= oField.name %></TH>
	<% NEXT %>
		</TR>
	<% dataRec.movefirst %>
	<% DO WHILE NOT dataRec.eof %>
		<% IF tt = 0 THEN
			color = "#C0C0C0"
			tt = tt + 1
		ELSE
			color = "#FFFFFF"
			tt = tt - 1
		END IF %>
	
	<TR bgcolor="<%= color %>">
		<% For each oField IN dataRec.Fields %>
			<TD><%= TRIM(oField.value) %>&nbsp;</TD>
		<% NEXT %>
	</TR>
	<% dataRec.movenext %>
	<% LOOP %>
</TABLE>

	
	<% ELSE %>
	<div align="center" style="font-weight:bold">Your query returned no data</div>
	<% END IF %>

<% 
' This code closes connection and releases ADO objects
IF dataRec.state = 1 THEN
	dataRec.Close
END IF
dataConn.Close
Set dataRec = Nothing
Set dataCmd = Nothing
Set dataConn = Nothing %>


<%
' DEFINE VARIABLES

Set dataConn = Server.CreateObject("ADODB.Connection")
dataConn.Open "","SQLGuest","SQLGuest"

Set dataCmd = Server.CreateObject("ADODB.Command")
dataCmd.ActiveConnection = dataConn
dataCmd.CommandType = &H0004
dataCmd.CommandText = "pb_GetTracking"
	
		Set dataParam			= dataCmd.CreateParameter("@P1_SSN")
		dataParam.Type			= 3
		dataParam.Direction		= &H0001
		dataParam.Size			= 4
		dataParam.Value			= m_iSSN
		dataCmd.Parameters.Append dataParam
	
	
		Set dataParam			= dataCmd.CreateParameter("@P2_DepNumber")
		dataParam.Type			= 3
		dataParam.Direction		= &H0001
		dataParam.Size			= 4
		dataParam.Value			= m_iDepNum
		dataCmd.Parameters.Append dataParam
	
Set dataRec = Server.CreateObject("ADODB.Recordset")
dataRec.Open dataCmd,,3,3

%>

