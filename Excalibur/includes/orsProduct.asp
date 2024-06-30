<%
'*************************************************************************************
'* Description	: SQL recordset connection(s) from the root deliverable tables 
'* Creator		: Harris, Valerie
'* Created		: 02/4/2016 - PBI 15660/ Task 16234
'*************************************************************************************
Dim oErrors		'OUR ERROR OBJECT
Dim sErrorMessage


'-------------------------------------------------------------------------------------
'* Purpose		: Return list of Products 
'* Inputs		: 
'-------------------------------------------------------------------------------------
Dim oRSProducts                	   'Recordset object.
Sub ListProducts(bOpen)	  
	On Error Resume Next
	Dim qsProducts                    'Query string.
	If bOpen=True Then
		Set oRSProducts = Server.CreateObject("ADODB.Recordset")
		Set oRSProducts.ActiveConnection = oConnect
		Set oErrors = oRSProducts.ActiveConnection.Errors
		qsProducts = "SELECT  v.TypeID, v.ID, f.Name, v.[Version], v.dotsname as Product, v.ProductLineId, " 
		qsProducts = qsProducts & "v.AllowDCR, v.ProductStatusID, v.DevCenter, dc.name as DevCenterName, case bs.BusinessID when 1 then case bs.Operation when 1 then isnull(ui.email, '') else '' end else '' end as PCEmail,"
        qsProducts = qsProducts & "ProductReleases = (SELECT ISNULL(COUNT(pv.Name),0) FROM ProductVersion_Release pr " 
		qsProducts = qsProducts & "INNER JOIN ProductVersionRelease pv on pr.ReleaseID = pv.ID "
		qsProducts = qsProducts & "WHERE pr.ProductVersionID = v.ID) "
        qsProducts = qsProducts & "FROM ProductFamily AS f with (NOLOCK) INNER JOIN "
        qsProducts = qsProducts & "ProductVersion AS v with (NOLOCK) ON f.ID = v.ProductFamilyID INNER JOIN "
        qsProducts = qsProducts & "DevCenter as dc with (Nolock) on v.DevCenter = dc.id INNER JOIN BusinessSegment bs on v.BusinessSegmentID = bs.BusinessSegmentID LEFT OUTER JOIN UserInfo ui on v.PCID = ui.UserID "
        qsProducts = qsProducts & "WHERE (v.TypeID IN (1, 3)) "
        qsProducts = qsProducts & "ORDER BY dc.name, v.dotsname, v.[Version]"
		oRSProducts.Open qsProducts, oConnect, adOpenForwardOnly, adLockReadOnly, adCmdText
		If oErrors.Count > 0 Then		    'CUSTOM ERROR HANDLING.   
            Response.Write("Error accessing database. Contact Pulsar Administrator.<br/>")
            Response.Write(Err.Description &"--"& Err.Source & "<br/>" & qsProducts)
            Response.End()
        End If
	ElseIf bOpen=False Then		            'CLOSE RECORDSET
		If Not (oRSProducts Is Nothing) Then
			oRSProducts.Close
			Set oRSProducts= Nothing
		End If
	End If
End Sub


'-------------------------------------------------------------------------------------
'* Purpose		: Return list of Release
'-------------------------------------------------------------------------------------
Dim oRSProductRelease	                    'Recordset object.
Sub ListProductRelease(bOpen)	  
	On Error Resume Next
	Dim qsProductRelease
	If bOpen=True Then
		Set oRSProductRelease = Server.CreateObject("ADODB.Recordset")
		Set oRSProductRelease.ActiveConnection = oConnect
		Set oErrors = oRSProductRelease.ActiveConnection.Errors
		qsProductRelease = "SELECT pr.ProductVersionID, pv.ID, pv.Name FROM ProductVersion_Release pr " _
                           & "INNER JOIN ProductVersionRelease pv " _
                           & "ON pr.ReleaseID = pv.ID " _
                           & "ORDER BY pr.ProductVersionID, pv.ID"
        oRSProductRelease.CursorType = adOpenStatic
        oRSProductRelease.CursorLocation = adUseClient
        oRSProductRelease.Open qsProductRelease, oConnect
        If oRSProductRelease.State = adStateOpen Then
            oRSProductRelease.Fields("ProductVersionID").Properties("Optimize") = True
        End If
	Else
		If Not (oRSProductRelease Is Nothing) Then
			oRSProductRelease.Close
			Set oRSProductRelease = Nothing
		End If
	End If
End Sub

'-------------------------------------------------------------------------------------
'* Purpose		: Return ProductIDs by Business Segment; Product Line for optional DCR Product Selection 
'* Inputs		: iSelectionType
'-------------------------------------------------------------------------------------
Dim oRSProductSelection	                    'Recordset object.
Sub ListProductSelection(bOpen, iSelectionType)	  
	On Error Resume Next
	Dim qsProductSelection	                'Query string.
	If bOpen=True And iSelectionType <> Empty Then
		Set oRSProductSelection = Server.CreateObject("ADODB.Recordset")
		Set oRSProductSelection.ActiveConnection = oConnect
		Set oErrors = oRSProductSelection.ActiveConnection.Errors
		qsProductSelection = "EXECUTE usp_ChangeRequest_GetProductSelectionList "
		qsProductSelection = qsProductSelection & ""& iSelectionType &""
        oRSProductSelection.Open qsProductSelection, oConnect, adOpenForwardOnly, adLockReadOnly, adCmdText
        If oErrors.Count > 0 Then		    'CUSTOM ERROR HANDLING.   
            Response.Write("Error accessing database. Contact Pulsar Administrator.<br/>")
            Response.Write(Err.Description &"--"& Err.Source & "<br/>" & qsProductSelection)
            Response.End()
        End If
	ElseIf bOpen=False Then		            'CLOSE RECORDSET
		If Not (oRSProductSelection Is Nothing) Then
			oRSProductSelection.Close
			Set oRSProductSelection = Nothing
		End If
	End If
End Sub

'-------------------------------------------------------------------------------------
'* Purpose		: Return Releases by Business Segment; for optional DCR Product Selection 
'* Inputs		: iSelectionType, iSelectionFilterID
'-------------------------------------------------------------------------------------
Dim oRSProductOption                        'Recordset object.
Sub ListProductOption(bOpen, iSelectionType, iSelectionFilterID)	  
	On Error Resume Next
	Dim qsProductOption	                    'Query string.
	If bOpen=True And iSelectionType <> Empty Then
		Set oRSProductOption = Server.CreateObject("ADODB.Recordset")
		Set oRSProductOption.ActiveConnection = oConnect
		Set oErrors = oRSProductOption.ActiveConnection.Errors
		qsProductOption = "EXECUTE usp_ChangeRequest_GetProductSelectionList "
		qsProductOption = qsProductOption & ""& iSelectionType &", " & iSelectionFilterID & " "
        oRSProductOption.Open qsProductOption, oConnect, adOpenForwardOnly, adLockReadOnly, adCmdText
        If oErrors.Count > 0 Then		    'CUSTOM ERROR HANDLING.   
            Response.Write("Error accessing database. Contact Pulsar Administrator.<br/>")
            Response.Write(Err.Description &"--"& Err.Source & "<br/>" & qsProductOption)
            Response.End()
        End If
	ElseIf bOpen=False Then		            'CLOSE RECORDSET
		If Not (oRSProductOption Is Nothing) Then
			oRSProductOption.Close
			Set oRSProductOption = Nothing
		End If
	End If
End Sub


'-------------------------------------------------------------------------------------
'* Purpose		: Return product's change request id, null is 0
'* Inputs		: iIssueID
'-------------------------------------------------------------------------------------
Dim oRSChangeRequest	                    'Recordset object.
Sub GetChangeRequest(bOpen, iIssueID)  
	On Error Resume Next
	Dim qsChangeRequest                     'Query string.
	If bOpen=True And iIssueID <> Empty Then
        Request.Write(iIssueID)
        Request.End()
		Set oRSChangeRequest = Server.CreateObject("ADODB.Recordset")
		Set oRSChangeRequest.ActiveConnection = oConnect
		Set oErrors = oRSChangeRequest.ActiveConnection.Errors
		qsChangeRequest = "SELECT ISNULL(ChangeRequestID,0) as ChangeRequestID FROM DeliverableIssues "
		qsChangeRequest = qsChangeRequest & "WHERE ID = " & iIssueID & ""
        oRSChangeRequest.Open qsChangeRequest, oConnect, adOpenForwardOnly, adLockReadOnly, adCmdText
        If oErrors.Count > 0 Then		    'CUSTOM ERROR HANDLING.   
            Response.Write("Error accessing database. Contact Pulsar Administrator.<br/>")
            Response.Write(Err.Description &"--"& Err.Source & "<br/>" & qsChangeRequest)
            Response.End()
        End If
	ElseIf bOpen=False Then		            'CLOSE RECORDSET
		If Not (oRSChangeRequest Is Nothing) Then
			oRSChangeRequest.Close
			Set oRSChangeRequest = Nothing
		End If
	End If
End Sub

'-------------------------------------------------------------------------------------
'* Purpose		: Return list of products submitted using the same change request form
'* Inputs		: iIssueID, iChangeRequestID
'-------------------------------------------------------------------------------------
Dim oRSChangeRequestGroup	                    'Recordset object.
Sub ListChangeRequestGroup(bOpen, iIssueID, iChangeRequestID)  
	On Error Resume Next
	Dim qsChangeRequestGroup                    'Query string.

	If bOpen=True And iIssueID <> Empty Then
		Set oRSChangeRequestGroup = Server.CreateObject("ADODB.Recordset")
		Set oRSChangeRequestGroup.ActiveConnection = oConnect
		Set oErrors = oRSChangeRequestGroup.ActiveConnection.Errors
		qsChangeRequestGroup = "SELECT DISTINCT ID, ProductVersionID, "
        qsChangeRequestGroup = qsChangeRequestGroup & "ProductName = ([dbo].[fnGetProductName](ProductVersionID)),  "
        qsChangeRequestGroup = qsChangeRequestGroup & "ProductVersionRelease, ChangeRequestID  "
        qsChangeRequestGroup = qsChangeRequestGroup & "FROM DeliverableIssues "
		qsChangeRequestGroup = qsChangeRequestGroup & "WHERE (ID = " & iIssueID & ") "
        If iChangeRequestID <> "0" Then
            qsChangeRequestGroup = qsChangeRequestGroup & "OR (ChangeRequestID = " & iChangeRequestID & ") "
        End If
        oRSChangeRequestGroup.Open qsChangeRequestGroup, oConnect, adOpenForwardOnly, adLockReadOnly, adCmdText
        If oErrors.Count > 0 Then		    'CUSTOM ERROR HANDLING.   
            Response.Write("Error accessing database. Contact Pulsar Administrator.<br/>")
            Response.Write(Err.Description &"--"& Err.Source & "<br/>" & qsChangeRequestGroup)
            Response.End()
        End If
	ElseIf bOpen=False Then		            'CLOSE RECORDSET
		If Not (oRSChangeRequestGroup Is Nothing) Then
			oRSChangeRequestGroup.Close
			Set oRSChangeRequestGroup = Nothing
		End If
	End If
End Sub
%>

