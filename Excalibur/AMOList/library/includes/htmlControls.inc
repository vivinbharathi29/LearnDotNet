<%

sub SeparateUnitValue(ByVal strAttributeValue, ByRef strValue, ByRef strUnit)
	dim intPos
	strValue = ""
	strUnit = ""
	intPos = InStr(1, strAttributeValue, " ", 1)
	' Get the key
	if (intPos > 0) and (strAttributeValue <> "") then
		strValue = Left(strAttributeValue, intPos - 1)
		strUnit = Right(strAttributeValue, (Len(strAttributeValue) - intPos))
	end if
end sub

function StatusOptions()
	dim adoRs
    ' Create new record set to merge
    Set adoRs = Server.CreateObject("ADODB.Recordset")
    adoRs.Fields.Append "ID", 3, 4
    adoRs.Fields.Append "Description", 129, 64
    adoRs.CursorLocation = 2
    adoRs.Open

	adoRs.AddNew
	adoRs.Fields("ID").Value = 1
	adoRs.Fields("Description").Value = "Planning"
	adoRs.AddNew
	adoRs.Fields("ID").Value = 2
	adoRs.Fields("Description").Value = "In Test"
	adoRs.AddNew
	adoRs.Fields("ID").Value = 3
	adoRs.Fields("Description").Value = "Production"
	adoRs.AddNew
	adoRs.Fields("ID").Value = 5
	adoRs.Fields("Description").Value = "End-of-Life"
	adoRs.AddNew
	adoRs.Fields("ID").Value = 4
	adoRs.Fields("Description").Value = "Invalid - Do not test"
	set StatusOptions = adoRs
end function

function ShippingOptions()
	dim adoRs
    ' Create new record set to merge
    Set adoRs = Server.CreateObject("ADODB.Recordset")
    adoRs.Fields.Append "ID", 3, 4
    adoRs.Fields.Append "Description", 129, 64
    adoRs.CursorLocation = 2
    adoRs.Open

	adoRs.AddNew
	adoRs.Fields("ID").Value = 1
	adoRs.Fields("Description").Value = "Compaq"
	adoRs.AddNew
	adoRs.Fields("ID").Value = 2
	adoRs.Fields("Description").Value = "Third-party"
	adoRs.AddNew
	adoRs.Fields("ID").Value = 3
	adoRs.Fields("Description").Value = "In Development"
'	adoRs.AddNew
'	adoRs.Fields("ID").Value = 4
'	adoRs.Fields("Description").Value = "TBD"
	set ShippingOptions = adoRs
end function

function CapacityOptions()
	dim adoRs
    ' Create new record set to merge
    Set adoRs = Server.CreateObject("ADODB.Recordset")
    adoRs.Fields.Append "ID", 129, 2
    adoRs.Fields.Append "Description", 129, 64
    adoRs.CursorLocation = 2
    adoRs.Open

	adoRs.AddNew
	adoRs.Fields("ID").Value = "MB"
	adoRs.Fields("Description").Value = "MB"
	adoRs.AddNew
	adoRs.Fields("ID").Value = "GB"
	adoRs.Fields("Description").Value = "GB"
	adoRs.AddNew
	adoRs.Fields("ID").Value = "TB"
	adoRs.Fields("Description").Value = "TB"
	set CapacityOptions = adoRs
end function

function BusSpeedOptions()
	dim adoRs
    ' Create new record set to merge
    Set adoRs = Server.CreateObject("ADODB.Recordset")
    adoRs.Fields.Append "ID", 129, 3
    adoRs.Fields.Append "Description", 129, 64
    adoRs.CursorLocation = 2
    adoRs.Open

	adoRs.AddNew
	adoRs.Fields("ID").Value = "MHz"
	adoRs.Fields("Description").Value = "MHz"
	adoRs.AddNew
	adoRs.Fields("ID").Value = "GHz"
	adoRs.Fields("Description").Value = "GHz"
	set BusSpeedOptions = adoRs
end function

function CacheSizeOptions()
	dim adoRs
    ' Create new record set to merge
    Set adoRs = Server.CreateObject("ADODB.Recordset")
    adoRs.Fields.Append "ID", 129, 2
    adoRs.Fields.Append "Description", 129, 64
    adoRs.CursorLocation = 2
    adoRs.Open

	adoRs.AddNew
	adoRs.Fields("ID").Value = "KB"
	adoRs.Fields("Description").Value = "KB"
	adoRs.AddNew
	adoRs.Fields("ID").Value = "MB"
	adoRs.Fields("Description").Value = "MB"
	adoRs.AddNew
	adoRs.Fields("ID").Value = "GB"
	adoRs.Fields("Description").Value = "GB"
	set CacheSizeOptions = adoRs
end function

function TextBox(ByVal sID, ByVal sName, ByVal sValue, ByVal sStyle, ByVal sMaxLength, ByVal sStatus)
	TextBox = TextBoxType(sID, sName, sValue, sStyle, sMaxLength, sStatus, "")
end function

function TextBoxType(ByVal sID, ByVal sName, ByVal sValue, ByVal sStyle, ByVal sMaxLength, ByVal sStatus, ByVal sType)
	dim htmlStr

	htmlStr = "<input " & sType & " id='" & sID & "' "
	htmlStr = htmlStr & "name='" & sName & "' "
	if isnull(sValue) then sValue = ""
	htmlStr = htmlStr & "value=""" & server.htmlencode(sValue) & """ "
	if (sName = "txtAGPSpeed") and (sValue = "") then
		htmlStr = htmlStr & "disabled='true' "
	end if
	if (sStatus <> "Uncommitted") and (sStatus <> "") and (sStatus <> "Draft") then
		htmlStr = htmlStr & "disabled='true' "
	end if
	htmlStr = htmlStr & "style='" & sStyle & "' "
	htmlStr = htmlStr & "maxLength='" & sMaxLength & "'></input>"
	TextBoxType = htmlStr
end function

function TextAreaRows(ByVal sID, ByVal sName, ByVal sValue, ByVal sStyle, ByVal sMaxLength, ByVal sStatus, sRows)
	dim htmlStr
	
	htmlStr = "<TEXTAREA rows='" & sRows & "' id='" & sID & "' "
	htmlStr = htmlStr & "name='" & sName & "' "
	htmlStr = htmlStr & "style='" & sStyle & "' "
	if (sStatus <> "Uncommitted") and (sStatus <> "") and (sStatus <> "Draft") then
		htmlStr = htmlStr & "readonly "
	end if
	htmlStr = htmlStr & "maxLength='" & sMaxLength & "'>"
	htmlStr = htmlStr & sValue & "</TEXTAREA>"
	TextAreaRows = htmlStr
end function

function TextArea(ByVal sID, ByVal sName, ByVal sValue, ByVal sStyle, ByVal sMaxLength, ByVal sStatus)

	TextArea = TextAreaRows(sID, sName, sValue, sStyle, sMaxLength, sStatus, "7")
end function

function DropDown(ByVal sID, ByVal sName, ByVal iDefaultValue, ByVal sStyle, ByVal adoOption, ByVal colName1, ByVal colName2, ByVal bAddBlank, ByVal sStatus)
    DropDown = DropDownHandler(sID, sName, iDefaultValue, sStyle, adoOption, colName1, colName2, bAddBlank, sStatus, "")
end function

function DropDownHandler(ByVal sID, ByVal sName, ByVal iDefaultValue, ByVal sStyle, ByVal adoOption, ByVal colName1, ByVal colName2, ByVal bAddBlank, ByVal sStatus, ByVal sHandler)
	dim htmlStr
	dim sValue
	
	htmlStr = "<SELECT id='" & sID & "' "
	htmlStr = htmlStr & "name='" & sName & "' onchange='bOnChange=true' "
	if (sStatus <> "Uncommitted") and (sStatus <> "") and (sStatus <> "Draft") then
		htmlStr = htmlStr & "disabled='true' "
	end if
	htmlStr = htmlStr & "style='" & sStyle & "' " & sHandler & " >"
	if not adoOption is nothing then
	    if (bAddBlank) then
	    	if (iDefaultValue = 0) then
	    		adoOption.Find colName2 & "='" & "TBD'"
	    		if (adoOption.EOF) then
	    			htmlStr = htmlStr & "<OPTION SELECTED value=0></OPTION>" 
	    		else
	    			iDefaultValue = adoOption.Fields(colName1).Value
	    		end if
	    	else
	    		htmlStr = htmlStr & "<OPTION value=0></OPTION>" 
	    	end if
	    end if
	    if adoOption.RecordCount > 0 then
	        adoOption.MoveFirst
	        while (not adoOption.EOF)
						sValue = adoOption.Fields(colName1).Value
						if isnull(sValue) then sValue = ""
	        	if (sValue = iDefaultValue) then
	        		htmlStr = htmlStr & "<OPTION SELECTED value=""" & server.htmlencode(sValue) & """>" & Trim(adoOption.Fields(colName2).Value) & "</OPTION>" 
	        	else
	        		htmlStr = htmlStr & "<OPTION value=""" & server.htmlencode(sValue) & """>" & Trim(adoOption.Fields(colName2).Value) & "</OPTION>" 
	        	end if
	        	adoOption.MoveNext
	        wend
        end if
    end if
	htmlStr = htmlStr & "</SELECT>"
	DropDownHandler = htmlStr
end function

function strDropDown(ByVal sID, ByVal sName, ByVal strDefaultValue, ByVal sStyle, ByVal adoOption, ByVal colName1, ByVal colName2, ByVal bAddBlank, ByVal sStatus)
	dim htmlStr
	dim sValue
	
	htmlStr = "<SELECT id='" & sID & "' "
	htmlStr = htmlStr & "name='" & sName & "' "
	if (sName = "ddlAGPSpeed") and (strDefaultValue = "") then
		htmlStr = htmlStr & "disabled='true' "
	end if
	if (sStatus <> "Uncommitted") and (sStatus <> "") and (sStatus <> "Draft") then
		htmlStr = htmlStr & "disabled='true' "
	end if
	htmlStr = htmlStr & "style='" & sStyle & "'>"
	if (bAddBlank) then
		if (strDefaultValue = "") then
			htmlStr = htmlStr & "<OPTION SELECTED value=''></OPTION>" 
		else
			htmlStr = htmlStr & "<OPTION value=''></OPTION>" 
		end if
	end if
	adoOption.MoveFirst
	while (not adoOption.EOF)
		sValue = adoOption.Fields(colName1).Value
		if isnull(sValue) then sValue = ""
		if (sValue = strDefaultValue) then
			htmlStr = htmlStr & "<OPTION SELECTED value=""" & server.htmlencode(sValue) & """>" & Trim(adoOption.Fields(colName2).Value) & "</OPTION>" 
		else
			htmlStr = htmlStr & "<OPTION value=""" & server.htmlencode(sValue) & """>" & Trim(adoOption.Fields(colName2).Value) & "</OPTION>" 
		end if
		adoOption.MoveNext
	wend
	htmlStr = htmlStr & "</SELECT>"
	strDropDown = htmlStr
end function

function RadioButton(ByVal sID, ByVal sName, ByVal sValue, ByVal sDisplay, ByVal sChecked, ByVal sStatus)
	dim htmlStr
	
	htmlStr = "<input id='" & sID & "' "
	htmlStr = htmlStr & "name='" & sName & "' "
	htmlStr = htmlStr & "type=radio "
	if (sStatus <> "Uncommitted") and (sStatus <> "") and (sStatus <> "Draft") then
		htmlStr = htmlStr & "disabled='true' "
	end if
	if isnull(sValue) then sValue = ""
	htmlStr = htmlStr & "value=""" & server.htmlencode(sValue) & """ " & sChecked & ">"
	htmlStr = htmlStr & sDisplay & "</input>"
	RadioButton = htmlStr
end function

function ListBox(ByVal sID, ByVal sName, ByVal sStyle, ByVal adoOption, ByVal colName1, ByVal colName2)
	dim htmlStr
	dim sValue
	
	htmlStr = "<select id='" & sID & "' " 
	htmlStr = htmlStr & "name='" & sName & "' multiple " 
	htmlStr = htmlStr & "style='" & sStyle & "'>"
	adoOption.MoveFirst
	while (not adoOption.EOF)
		sValue = adoOption.Fields(colName1).Value
		if isnull(sValue) then sValue = ""
		htmlStr = htmlStr & "<OPTION value=""" & server.htmlencode(sValue) & """>" & Trim(adoOption.Fields(colName2).Value) & "</OPTION>" 
		adoOption.MoveNext
	wend
	htmlStr = htmlStr & "</SELECT>"
	ListBox = htmlStr
end function

%>