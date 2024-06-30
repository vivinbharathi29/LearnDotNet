<%
'*************************************************************************************
'* FileName		: oRSProgramMatrix.asp
'* Description	: SQL recordset connection(s) from the root deliverable tables 
'* Creator		: Harris, Valerie
'* Created		: 09/15/2016 - PBI 23434 / Task 24367
'*************************************************************************************
Dim oErrors		'OUR ERROR OBJECT
Dim sErrorMessage

'-------------------------------------------------------------------------------------
'* Purpose		: Return dates for the selected Brand ID.
'* Inputs		: Brand ID
'-------------------------------------------------------------------------------------
Dim oRSProgramMatrixDates	   'Recordset object.
Sub GetProgramMatrixDates(bOpen, BrandID)	  
	On Error Resume Next
	Dim qsProgramMatrixDates  'Query string.
	If bOpen=True Then
		Set oRSProgramMatrixDates = Server.CreateObject("ADODB.Recordset")
		Set oRSProgramMatrixDates.ActiveConnection = oConnect
		Set oErrors = oRSProgramMatrixDates.ActiveConnection.Errors
		qsProgramMatrixDates = "EXECUTE usp_ListProgramMatrixPublishDates "
		qsProgramMatrixDates = qsProgramMatrixDates & ""& BrandID &""
		oRSProgramMatrixDates.Open qsProgramMatrixDates, oConnect, adOpenForwardOnly, adLockReadOnly, adCmdText
		If oErrors.Count > 0 Then		'CUSTOM ERROR HANDLING.   
            Response.Write("Error accessing database. Contact Pulsar Administrator.")
            Response.End()
        End If
	ElseIf bOpen=False Then		        'CLOSE RECORDSET
		If Not (oRSProgramMatrixDates Is Nothing) Then
			oRSProgramMatrixDates.Close
			Set oRSProgramMatrixDates = Nothing
		End If
	End If
End Sub
%>
