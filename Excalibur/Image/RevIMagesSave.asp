<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
    if (typeof (txtSuccess) != "undefined") {
        if (txtSuccess.value == "1") {
            if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                parent.window.parent.reloadFromPopUp(pulsarplusDivId);
                parent.window.parent.closeExternalPopup();
            }
            else {
                window.returnValue = 1;
                window.parent.close();
            }
        }
    }
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">
<font size=2 face=verdana>This page should close automatically when it is done.<BR><BR>Saving. Please wait
<%
	Response.Flush
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	cn.CommandTimeout = 5400
	set rs = server.CreateObject("ADODB.recordset")
	cn.CursorLocation = adUseClient

	cn.BeginTrans

	dim SKUArray
	dim SKUItem
	dim strID
	dim strValue
	dim ItemArray
	dim strImages
	dim strImageID
	dim strImages2Add
	dim strExceptions
	dim strNewImages
	
	strImages2Add = ""
	strExceptions = ""
	strImages = ""
	dim strSKUNUmber
	dim NewImageDefID

    'response.write request("txtSKU")
	SKUArray = split(request("txtSKU"),",")
	for each SKUItem in SKUArray
		if trim(SKUItem) <> "" then
'			Response.Write "Adding ImageDef: "
			ItemArray = split(SKUItem,chr(9))
			'Response.Write ItemArray(0) & "_"
			if ubound(ItemArray) = 1 then 'Else Empty
'				Response.Write ItemArray(1)
				strSKUNUmber = ItemArray(1)
			else
'				Response.Write "<font color=red>None</font>"
				strSKUNUmber = ""
			end if
'			Response.Write "<BR>"
			
			
			
			'Start: Add new IMage Definition			
			set cm = server.CreateObject("ADODB.Command")
			cm.CommandType =  &H0004
			cm.ActiveConnection = cn
		
			cm.CommandText = "spCopyImageDefinition"	

			Set p = cm.CreateParameter("@CopyID", 3,  &H0001)
			p.Value = clng(ItemArray(0))
			cm.Parameters.Append p
		
			Set p = cm.CreateParameter("@SKU", 200, &H0001, 20)
			p.value = left(strSKUNUmber,20)
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@NewID", 3,  &H0002)
			cm.Parameters.Append p

			cm.Execute rowschanged

			if rowschanged <> 1 then
				FoundErrors = true
			else
				NewImageDefID = cm("@NewID")
			end if	
		
		set cm = nothing
'			Response.Write "<BR>NewImage DefinitionID:" & NewImageDefID & "<BR>"
	'Done: Add new IMage Definition			
			
			
			
			
			if not FoundErrors then
				rs.Open "spListImages4Definition " & ItemArray(0),cn,adOpenKeyset
				do while not rs.EOF
				
					'Start: Add Images			
					set cm = server.CreateObject("ADODB.Command")
					cm.CommandType =  &H0004
					cm.ActiveConnection = cn
		
					cm.CommandText = "spCopyRegion4Image"	

					Set p = cm.CreateParameter("@CopyID", 3,  &H0001)
					p.Value = clng(rs("ID"))
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@ImageDefinitionID", 3,  &H0001)
					p.Value = clng(NewImageDefID)
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@NewID", 3,  &H0002)
					cm.Parameters.Append p

					cm.Execute rowschanged

					if rowschanged <> 1 then
						FoundErrors = true
						exit do
					else
						NewImageID = cm("@NewID")
					end if	
		
					set cm = nothing
				
				'-----End Add images
				
				'-----Log Image Updates
				
				set cm = server.CreateObject("ADODB.Command")
				cm.ActiveConnection = cn
				cm.CommandType =  &H0004
				cm.CommandText = "spAddIMageLog"

				Set p = cm.CreateParameter("@EmployeeID", 3,  &H0001)
				p.Value = clng(request("txtCurrentUserID"))
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@DCRID", 3,  &H0001)
				p.Value = null
				cm.Parameters.Append p
		
				Set p = cm.CreateParameter("@ImageID", 3,  &H0001)
				p.Value = NewImageID
				cm.Parameters.Append p
		
				Set p = cm.CreateParameter("@Details", 200,  &H0001,7500)
				p.Value = "Added"
				cm.Parameters.Append p
		
				cm.Execute rowschanged
	
				if rowschanged <> 1 then
					FoundErrors = true
				end if				
				
				
				'----End Log IMage Updates
					strImages = strImages & "," & rs("ID")
					strNewImages = strNewImages & "," & NewImageID 
					rs.MoveNext
				loop
				rs.Close
			end if
		end if
	next
'Response.Write "<BR><BR>"
			if strImages <> "" then
				strimages = mid(strImages,2)
			end if
			if strNewImages <> "" then
				strNewimages = mid(strNewImages,2)
			end if
'			Response.Write "<b><table border=1 width=""100%""><TR><TD>Original Image ID</TD><TD>New Image ID</TD></TR>"
'			Response.Write "<TR><TD>" & replace(strimages,",",", ") & "</TD><TD>" & replace(strNewimages,",",", ") & "</TD></TR></table>"
			
'			Response.Write "<BR><BR>"
			
			ImageArray = split(strImages,",")
			NewArray = split(strNewImages,",")
			

		strBGColor = split("Lavender,Thistle",",")
		'Update Deliverable Root and version images
		if not FoundErrors then

			for ReportCount = 0 to 1
			rs.Open "spListDeliverables4ProductInImage " & clng(request("txtProductID")) & "," & ReportCount,cn,adOpenStatic
'			Response.Write "<TABLE bgcolor=""" & strBGColor(ReportCount) & """width=""100%"" border=1><TR><TD>ID</TD><TD>OldIMageList</TD><TD>NewImageList</TD>"
			dim strLastID
			dim strLastImageList
			do while not rs.EOF
				if trim(rs("Images")) <> "" then
'					Response.Write "<TR><TD>" & rs("ID") & "</TD>"
'					Response.Write "<TD>" & rs("Images") & "</TD>"
					
					strImages2Add = ""
					strExceptions = ""
					Response.Write "."					
					Response.flush
					
					for i = lbound(ImageArray) to ubound(ImageArray)
						if trim(ImageArray(i)) <> "" and instr(", " & rs("Images"),", " & trim(ImageArray(i)) & ",") > 0 then 
							strImages2Add = strImages2Add & ", " & NewArray(i)
						end if
						if  trim(ImageArray(i)) <> "" and instr(rs("Images"),"(" & trim(ImageArray(i)) & "=") > 0 then
							strTemp = mid(rs("Images"),instr(rs("Images"),"(" & trim(ImageArray(i)) & "=")+1)
							strTemp = mid(strTemp,instr(strTemp,"=")+1)
							strTemp = left(strTemp, instr(strTemp,")"))
							strExceptions = strExceptions & "(" & trim(NewArray(i)) & "=" & strTemp  & ";"
						end if
					next
					if strImages2Add <> "" or strExceptions <> "" then
						strImages2Add = mid(strImages2Add,2) & ","
						TypeArray = split(rs("Images"),":")
						strImagesLeft = TypeArray(0)
						if ubound(TypeArray)>0 then
							strImagesRight = TypeArray(1)
						else
							strImagesRight = ""
						end if
						strNewImageString = ""
						if strImagesRight = "" then
							strNewImageString = strImagesLeft & strimages2add 
'							Response.Write"<TD>" & strNewImageString
						else
							strNewImageString = strImagesLeft & strimages2add & ":" & strImagesRight & strExceptions
'							Response.Write"<TD>" & strNewImageString
						end if
						if strNewImageString <> "" then
						
							'STart - Image List Update
							set cm = server.CreateObject("ADODB.Command")
							cm.CommandType =  &H0004
							cm.ActiveConnection = cn
		
							cm.CommandText = "spUpdateImageList"	

							Set p = cm.CreateParameter("@DelID", 3,  &H0001)
							p.Value = clng(rs("ID"))
							cm.Parameters.Append p

							Set p = cm.CreateParameter("@ReportID", 3,  &H0001)
							p.Value = clng(ReportCount)
							cm.Parameters.Append p


							Set p = cm.CreateParameter("@ImageList", 201, &H0001, 2147483647)
							p.value = strNewImageString
							cm.Parameters.Append p

							cm.Execute rowschanged
	
							if rowschanged <> 1 then
								FoundErrors = true
'								Response.Write "p1:" & cm.Parameters(0).value & "<BR>"
'								Response.Write "p2:" & cm.Parameters(1).value & "<BR>"
'								Response.Write "p3:" & cm.Parameters(2).value & "<BR>"
'								Response.Write "p4:" & cm.Parameters(3).value & "<BR>"
								exit do
							end if	
		
							'rs("Images") = strNewImageString
		
							set cm = nothing
						
							'End - ImageListUpdate
							
'							Response.Write " - Update Confirmed"
						end if
'						Response.Write "</TD></TR>"
'					else
'						Response.Write "<TD>Update Not Required</TD></TR>"
					end if
				end if
				rs.MoveNext
			loop
			rs.Close	
'			Response.Write "</TABLE><BR><BR>"
			
			next
		end if

	if FoundErrors then
		cn.RollbackTrans
		strProcStatus = "0"
	else
		cn.CommitTrans
		strProcStatus = "1"
	end if

	set rs = nothing
	cn.Close
	set cn = nothing

%>
Done</font>
<BR><BR><INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strProcStatus%>">
<BR><INPUT type="hidden" id=txtUser name=txtUser value="<%=request("txtCurrentUserID")%>">
</BODY>
</HTML>
