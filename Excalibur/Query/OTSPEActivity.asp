<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	lblWait.innerText= Date();
}

//-->
</SCRIPT>
<!-- #include file = "../includes/noaccess.inc" -->

<STYLE>
Body{
	FONT-Size: x-small;
	FONT-FAMILY: Verdana;
}
TD{
	font-family:Verdana;
	font-size: x-small;
}
</STYLE>
<BODY LANGUAGE=javascript onload="return window_onload()">
<center>
	<font size=3 face=verdana><b>PE Observations Investigated</b></font><BR><BR>
	<font face=verdana size=2><span id=lblWait>Generating Report.  Please wait...</span></font><BR><BR>
</center>
<%
	dim PEIDArray
	dim PENameArray
	dim PEClosedArray
	dim PETimeArray
	dim i
	dim j
	dim WeekCount
	dim WeekTotals
	dim WeekDate
	
	WeekDate = formatdatetime(dateadd("d",-weekday(Now,VBMonday),Now),vbshortdate)	
	
	PEIDArray = split("8124,8120,8574,8122,8123",",")
	'PEIDArray = split("benson,DavisAndre,afisher,valencig,RZainfeld",",")

	PENameArray = split("Benson,Davis,Fisher,Valencia,Zainfeld",",")
'	PEIDArray = split("benson,smilechen,seanfschiang,DavisAndre,afisher,michaelshhsu,bnorcross,valencig,cm_wang,RZainfeld",",")
'	PENameArray = split("Benson,Chen,Chiang,Davis,Fisher,Hsu,Norcross,Valencia,Wang,Zainfeld",",")
	PEClosedArray = split("0,0,0,0,0",",")
	PETimeArray = split("0,0,0,0,0",",")
	PEOTSArray = split("0,0,0,0,0,0",",")
	PEDaysArray = split("0,0,0,0,0,0",",")

	'Response.Flush
				
	strConnect = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	set cn = server.CreateObject("ADODB.Connection")
	
	cn.ConnectionString = strConnect
	cn.CommandTimeout = 300
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn
	

	'Response.Write "<font face=Verdana size=3><b></b><BR><BR>"
	Response.Write "<table bgcolor=Ivory><tr bgcolor=""006697""><td rowspan=2><font color=""white""><strong>&nbsp;Week<BR>&nbsp;Ending&nbsp;</strong></font>"
	for i = lbound(PENameArray) to ubound(PENameArray)
		Response.write "<td colspan=2 width=200 align=center ><font color=""white""><strong>&nbsp;" & PENameArray(i) & "</strong></font>"
	next
   Response.Write "<td colspan=2 width=200 align=""center""><font color=""white""><strong>&nbsp;All</strong></font>"
   Response.Write "</TR>"
	Response.Write "<tr bgcolor=""Gainsboro"">"
	for i = lbound(PENameArray) to ubound(PENameArray)
		Response.write "<td width=100 align=center ><font color=""black""><b>&nbsp;OTS</font></td>"
		Response.write "<td width=100 align=center ><font color=""black""><b>&nbsp;Avg.&nbsp;Days</font></td>"
	next

    Response.Write "<td width=100 align=""center""><font color=""black""><b>&nbsp;OTS&nbsp;</font></td>"
    Response.Write "<td width=100 align=""center""><font color=""black""><b>&nbsp;Avg.&nbsp;Days</font></td>"
    Response.write "</TR>"
	
	rs.Open "tmpOTSPEActivityReport",cn,adOpenStatic
	if not (rs.EOF and rs.BOF) then
		for j = 1 to 13
			WeekCount = 0
			WeekTotals = 0
			Response.Write "<tr><td bgcolor=""006697""><font color=""white""><strong>&nbsp;" & formatdatetime(DateAdd("ww",-j+1,WeekDate),vbShortDate) & "&nbsp;</strong></font>"
			
			do while rs("DuringWeek") < j		
				rs.MoveNext
			loop
			PEClosedArray = split("0,0,0,0,0",",")
			PETimeArray = split("0,0,0,0,0",",")

			do while  not rs.EOF
				if rs("DuringWeek") <> j then
					exit do
				end if
				for i = lbound(PEIDArray) to ubound(PEIDArray)
					if lcase(trim(PEIDArray(i))) = lcase(trim(rs("ownerID") & "")) then
						WeekCount = WeekCount + rs("Resolved")
						WeekTotals =  WeekTotals + (rs("Resolved") * rs("AverageDaysWorked"))

						PEClosedArray(i) = rs("Resolved") & ""
						PETimeArray(i) =rs("AverageDaysWorked") & ""
						
						PEOTSArray(i) = PEOTSArray(i) + rs("Resolved")
						PEDaysArray(i) = PEDaysArray(i) + (rs("AverageDaysWorked") * rs("Resolved"))
						exit for
					end if
				next
				rs.MoveNext
			loop
			for i = lbound(PEIDArray) to ubound(PEIDArray)
				Response.Write "<td bgcolor=gainsboro align=middle>" & PEClosedArray(i) & "</td><td  bgcolor=gainsboro align=middle>" & PETimeArray(i) & "</td>"
			next
			if WeekCount = 0 then
				Response.Write "<td bgcolor=lightsteelblue align=middle>0</td><td bgcolor=lightsteelblue align=middle>0</td>"
			else
				Response.Write "<td bgcolor=lightsteelblue align=middle>" & WeekCount & "</td><td bgcolor=lightsteelblue align=middle>" & round(WeekTotals/WeekCount) & "</td>"
				PEOTSArray(ubound(PEOTSArray)) = PEOTSArray(ubound(PEOTSArray)) + WeekCount
				PEDaysArray(ubound(PEDaysArray)) = PEDaysArray(ubound(PEDaysArray)) + (round(WeekTotals/WeekCount) * WeekCount)
			end if
			Response.Write "</TR>"
	
		next

		Response.Write "<tr bgcolor=lightsteelblue><td bgcolor=""006697""><font color=""white""><strong>&nbsp;Total&nbsp;</strong></font>"
		for i = lbound(PEIDArray) to ubound(PEIDArray)
			Response.Write "<td align=middle>" & PEOTSArray(i) & "</td>"
			if PEOTSArray(i) = 0 then
				Response.Write "<td align=middle>0</td>"
			else
				Response.Write "<td align=middle>" & round(PEDaysArray(i)/PEOTSArray(i)) & "</td>"
			end if
		next
		Response.Write "<td align=middle>" & PEOTSArray(ubound(PEOTSArray)) & "</td>"
		if PEDaysArray(ubound(PEDaysArray)) = 0 then
			Response.Write "<td bgcolor=lightsteelblue align=middle>0</td>"
		else
			Response.Write "<td bgcolor=lightsteelblue align=middle>" & round(PEDaysArray(ubound(PEDaysArray))/PEOTSArray(ubound(PEOTSArray))) & "</td>"
		 end if
		Response.Write "</TR>"

	else
		Response.Write "No Observations Found"
	end if
	


	Response.Write "</table>"


	set rs = nothing
	cn.Close
	set cn = nothing
	
%>



</BODY>
</HTML>
