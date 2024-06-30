<%
'********************************************************************************
'* Display values for given variables
'* 
'* Parameters:
'*  parmVars = string of comma delimited variable names
'********************************************************************************
SUB ShowVars(parmVars)
      dim myparameters, strAll, counter
      myparameters=SPLIT(parmVars,",")
      for counter=0 to ubound(myparameters)
         strall=strall & "<b>" & myparameters(counter) & "</b>=" & eval(myparameters(counter)) & "<br>"
      next
      response.write strall
END SUB

'********************************************************************************
'* Display resulting recordset
'* 
'* Parameters:
'*  oRecordSet = recordset to display
'********************************************************************************
sub PrintDebug( oRecordSet )
	dim i
	dim intCount

	intCount = 0
	if oRecordSet.RecordCount <> 0 then
		oRecordSet.MoveFirst
	end if
	Do While Not oRecordSet.EOF
		intCount = intCount + 1
		response.write "<b>" & intCount & "</b>= "
		for i=0 to oRecordSet.Fields.Count-1
		response.write oRecordSet(i).Name & " (<b>" & oRecordSet(i) & "</b>)&nbsp;&nbsp;&nbsp;"
		next
		response.write "<br>"
		oRecordSet.MoveNext
	loop
	if oRecordSet.RecordCount <> 0 then
		oRecordSet.MoveFirst
	end if
	response.write "-----------------------<br>"
end sub

'********************************************************************************
'* Display passed values for debugging purposes
'* 
'* Parameters:
'*  none
'********************************************************************************
sub PrintRequest()
	dim item
	
	Response.Write "<b>Querystring:</b><br>=================<br>"
	for each item in Request.QueryString
		Response.Write "<b>" & item & "</b>=" & Request.QueryString(item) & "<br>"
	next
	Response.Write "<BR><b>Form:</b><br>=================<br>"
	for each item in Request.Form
		Response.Write "<b>" & item & "</b>=" & Request.Form(item) & "<br>"
	next
	Response.Write "=================<br>"
end sub

'********************************************************************************
'* Display string with a <br> at the end
'* 
'* Parameters:
'*  strString = string to display
'********************************************************************************
SUB PrintLine(strString)
	response.write strString & "<br>"
END SUB

'********************************************************************************
'* Display string with a <br> at the end and then end the program
'* 
'* Parameters:
'*  strString = string to display
'********************************************************************************
SUB PrintEnd(strString)
	response.write strString & "<br>"
	response.end
END SUB

%>


