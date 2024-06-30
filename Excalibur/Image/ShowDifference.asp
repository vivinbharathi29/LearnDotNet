<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<STYLE>
TD{
   FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana
    }
TH{
   FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana
    }
</STYLE>


<BODY>

<font size=3 face=verdana><b>Compare Strings</b></font><BR><BR>

<%
	dim StringArray
	dim i
	dim j
	dim CellColor
	dim strSystem
	
	
	
	StringArray = split(request("chkResult"),",")
	
	if ubound(StringArray) > 0 then
	
	for i = lbound(StringArray) to ubound(StringArray)
		if left(trim(stringarray(i)),1) = "E" then
			strSystem = "<b>Excalibur</b>"
		elseif left(trim(stringarray(i)),1) = "C" then
			strSystem = "<b>Conveyor</b>"
		else
			strSystem = ""
		end if
		Response.Write "<table border=1><tr><td colspan=" & len(stringarray(i)) & ">" & strSystem & ": " & mid(trim(stringArray(i)),2) & "<td></tr><tr>"
		for j = 2 to len(trim(stringarray(i)))
			if i<ubound (stringarray)then
				if mid(trim(stringarray(i)),j,1) = mid(trim(stringarray(i+1)),j,1) then
					CellColor = "white"
				else
					CellColor = "mistyrose"
				end if
			elseif ubound (stringarray)>= 1 then
				if mid(trim(stringarray(i)),j,1) = mid(trim(stringarray(i-1)),j,1) then
					CellColor = "white"
				else
					CellColor = "mistyrose"
				end if
			else
				CellColor = "white"
			end if
			response.write "<TD bgcolor=" & CellColor & " align=middle>" & mid(trim(stringarray(i)),j,1) & "<BR>" & asc(mid(trim(stringarray(i)),j,1)) & "</td>"
		next
		Response.Write "</tr></table>"
	next
	else
		Response.Write "<br>You must select 2 or more lines to compare.<br>"
	end if
%>

</BODY>
</HTML>
