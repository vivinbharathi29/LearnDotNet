<%
Function XmlSafe( value )
	Dim output
	output = value
	'output = replace(value, "'", "&apos;")
	output = replace(output, "&", "&amp;")
	output = replace(output, "<", "&lt;")
	output = replace(output, ">", "&gt;")
	output = replace(output, """", "&quot;")
	output = replace(output, vbcrlf, "&#10;")
	
	XmlSafe = output
End Function
%>