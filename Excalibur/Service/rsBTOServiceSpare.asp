<!-- #include file="../_ScriptLibrary/jsrsServer.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<% jsrsDispatch( "GetSKUSpareKits" ) %>
<script language="vbscript" runat="server">

'
' Get SKU Spare Kit Records 
'
Function GetSKUSpareKits(strSKU)
	Dim strHTTPRef
	Dim strData: strData=""
	Dim strRC: strRC=""
	Dim intRecs: intRecs=0
	Dim rs, dw, cn, cmd
	Dim i
	Dim strSKID

	strHTTPRef=TRIM(UCASE(Request.ServerVariables("HTTP_REFERER")))

	IF(InStr(strHTTPRef, "/PMVIEW.ASP")>0 or InStr(strHTTPRef, "/PMVIEWSERVICE.ASP")>0) THEN

		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
 
		Set cmd = dw.CreateCommAndSP(cn, "usp_SelectSKUSpareKits_BTOSS") 
		dw.CreateParameter cmd, "@SKU", adVarchar, adParamInput, 18, UCase(strSKU)
		dw.CreateParameter cmd, "@INCLUDE_AVS", adBoolean, adParamInput, 1, 1
		dw.CreateParameter cmd, "@SKU_DESC", adVarchar, adParamInputOutput, 50, ""
		dw.CreateParameter cmd, "@PROD_FVB_DESC", adVarchar, adParamInputOutput, 150, ""
		dw.CreateParameter cmd, "@RETURN_CODE", adInteger, adParamInputOutput, 8, 0
		dw.CreateParameter cmd, "@RETURN_DESC", adVarchar, adParamInputOutput, 255, ""

		Set rs = dw.ExecuteCommAndReturnRS(cmd)

	    	Dim strDataSource
		Dim intRC

		strDataSource=""

		intRC=CINT(cmd.Parameters("@RETURN_CODE"))

		IF(intRC=0) THEN
	

		    	If Not rs.Eof Then
        		' Create Table Header
			strData="<TABLE BORDER='1' class='MatrixTable'><THEAD><TR>"

			For i=0 to rs.Fields.Count-1								'lbound(rs.Fields) to ubound(rs.Fields)
				'strData=strData & "<TH class='pinnedRow'><b>" & rs.Fields(i).Name & "</b></TH>"

				If i>0 Then
					strData=strData & "<TH><b>" & rs.Fields(i).Name & "</b></TH>"
				End If
			Next

			strData=strData & "</TR></THEAD><TBODY>"

			Do Until rs.EOF
				intRecs=intRecs+1
				strData=strData & "<TR onmouseover='rslRow_onmouseover()' onmouseout='rslRow_onmouseout()'>"

				For i=0 to rs.Fields.Count-1							'lbound(rs.Fields) to ubound(rs.Fields)
					If Not IsNull(rs.Fields(i).value) Then

						Select Case i
						Case 0: ' SKID
							strSKID=rs.Fields(i).value

						Case 1: ' Spare Kit #
							strData=strData & "<TD SKID='" & strSKID & "' title='Click to Perform a QuickSearch on this Spare Kit' onclick=" & "" & "QuickSearch('" & rs.Fields(i) & "');" & "" & " style='cursor:pointer'>" & rs.Fields(i).Value & "</TD>"

						Case 4: ' AV Qualifiers
							strData=strData & "<TD title='Click to Perform a QuickSearch on this AV Qualifier list' onclick=" & "" & "QuickSearch('" & rs.Fields(i) & "');" & "" & " style='cursor:pointer'>" & rs.Fields(i).Value & "</TD>"

						Case Else: ' GPG Description, Category Name
							strData=strData & "<TD>" & rs.Fields(i).Value & "</TD>"

						End Select

					Else
						strData=strData & "<TD>&nbsp;</TD>"
					End If

				Next

				strData=strData & "</TR>"

				rs.MoveNext
			Loop

			rs.Close

			strData=strData & "</TBODY></TABLE>"

			Else
				strData=""
				strDataSource="<table border='0' class='MatrixTable'><tr><td>" & cmd.Parameters("@RETURN_DESC") & "</td><td>" & CStr(intRecs) & " record(s)</td></tr></table>"
    			End If	
	
			strDataSource="<table border='0' class='MatrixTable'><tr><td>" & cmd.Parameters("@RETURN_DESC") & "</td><td>" & CStr(intRecs) & " record(s)</td></tr></table>"

		ELSE
			strData=""
			strDataSource="<table border='0' class='MatrixTable'><tr><td>" & cmd.Parameters("@RETURN_DESC") & "</td><td>" & CStr(intRecs) & " record(s)</td></tr></table>"

		END IF

		Set rs=Nothing	

		Set cmd=Nothing

		cn.close
		Set cn=Nothing
		
		GetSKUSpareKits=strDataSource & strData  
	ELSE
		GetSKUSpareKits="<TABLE><TR><TD><b>INVALID ACCESS ATTEMPT.</b></TD></TR></TABLE>"
	END IF


End Function
</script>