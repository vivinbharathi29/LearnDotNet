<%@  language="VBScript" %>
<!-- #include file = "../../includes/noaccess.inc" -->
<!-- #include file="emailwrapper.asp" -->
<%
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
%>
<html>
<head>
	<title>Immediate Action Required! - CMIT Component PMs/Developers</title>
	<script id="clientEventHandlersJS" type="text/javascript">
<!--
	function window_onload() {
		if (txtAuto.value == "1") {
			window.opener = 'X';
			window.open('', '_parent', '');
			window.close();
		}
	}
	//-->
	</script>
</head>
<body onload="window_onload();">
	<%
	dim index, output
	dim productGroups, senders
	
	'product group names must be surrounded by single quotes
	productGroups = split("'BNB 2013','BNB 2014';'BNB 2015';'BNB 2016'", ";")
	senders = split("pony.ma@hp.com;pony.ma@hp.com;pony.ma@hp.com", ";")
	subjectTitles = split("Sustaining Platforms P1 Issues;2015 Platforms P1 Issues;2016 Platforms P1 Issues", ";")
	for index = 0 to ubound(productGroups)
		output = compileNagEmailFor(productGroups(index))
		sendNagEmail subjectTitles(index), senders(index), output
	next

	function ReplaceAndHTMLEncodeFrom(someText)
		dim sameCellBreak
		dim workingText
		sameCellBreak = "<BR style=""mso-data-placement: same-cell""/>"
		workingText = someText & ""
		workingText = replace(workingText, chr(160), " ")
		workingText = Server.HTMLEncode(workingText)
		workingText = replace(workingText, vbcr, sameCellBreak)
		workingText = replace(workingText, vblf, sameCellBreak)
		workingText = replace(workingText, vbcrlf, sameCellBreak)
		ReplaceAndHTMLEncodeFrom = workingText
	end function
	
	function compileNagEmailFor(productGroups)
		dim strOutput
		dim strSQL
		dim cnExcalibur,cnSIO,rsSIO,rs
		dim strAffectedProducts
		dim recordCount
		dim strDeliverableIds,strProductIds,strGroupProductIds
		
		set rs = server.CreateObject("ADODB.recordset")
		set cnExcalibur = server.CreateObject("ADODB.Connection")
		cnExcalibur.ConnectionString = Session("PDPIMS_ConnectionString") 
		cnExcalibur.Open
		cnExcalibur.CommandTimeout = 120 '90 '50 '180
		set cnSIO = server.CreateObject("ADODB.Connection")
		cnSIO.ConnectionString = "Provider=SQLOLEDB.1;Data Source=housireport01.auth.hpicorp.net;Initial Catalog=sio;User ID=Excalibur_RO;Password=sQ8be9AyqPQKEcqsa3mE;"
		cnSIO.Open
		cnSIO.CommandTimeout = 120 '90 '50 '180
		set rsSIO = server.CreateObject("ADODB.recordset")

		strDeliverableIds = getCommercialDeliverableIds(cnExcalibur, rs)
		strProductIds = getActiveCommercialProductIds(cnExcalibur, rs)
		strGroupProductIds = getProductIdsFor(productGroups, cnExcalibur, rs)
		
		strOutput = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small} A:link {COLOR: Blue;} A:visited{COLOR: Blue;} A:hover {COLOR: red;}</STYLE>" & _
			"<a name=""toc""><font size=2 face=verdana><b>Contents</b></font></a>" & _
			"<ul>" & _
			"<li><a href=""#FuncTestP1""><font size=2 face=verdana><b>Functional Test P1 Observations on Deliverables that are in the CMIT NB images</b></font></a>" & _
			"<li><a href=""#FixPastDue""><font size=2 face=verdana><b>CMIT NB P1 Observations in &quot;Fix in Progress&quot; or &quot;Fix in Progress – Waiting on Vendor&quot; state which are PAST DUE or have NO Target Date.</b></font></a>" & _
			"<li><a href=""#Stale""><font size=2 face=verdana><b>CMIT NB P1 Observations Last Modified over 7 days ago.</b></font></a>" & _
			"<li><a href=""#Rsvp""><font size=2 face=verdana><b>CMIT NB P1 Observations in Need Info/Retest State &gt; 24Hrs.</b></font></a>" & _
			"</ul>"

	'	Functional Test P1 Observations on Deliverables that are in the Commercial images
		strOutput = strOutput & "<font size=2 face=verdana><a name=""FuncTestP1""><b>Functional Test P1 Observations on Deliverables that are in the CMIT NB images</b></a> <a href=""#toc"">(top)</a></font><br><br>"
		blnHeaderWritten = false
		recordCount = 0

		'get list of functional test observations on the previously compiled list of component IDs
		if strDeliverableIds <> "" then
			strSQL = _
				"select distinct" & _
					" o.ObservationId" & _
					",o.Priority" & _
					",o.PrimaryProduct as Product" & _
					",o.Component + ' [' +  o.ComponentVersion + ']' as Deliverable" & _
					",o.State" & _
					",o.DaysOpen" & _
					",o.OwnerName" & _
					",o.Owner" & _
					",o.ShortDescription as Summary" & _
				" from dbo.SI_observation_Report o with (NOLOCK)" & _
				" where (o.GatingMilestone='Functional Test' or o.ProductFamily like 'Func Tst%')" & _
				" and o.Status <> 'Closed'" & _
				" and o.Priority = 1" & _
				" and o.ExcaliburNumber in (" & strDeliverableIds & ")" & _
				" and o.PrimaryProduct not like '%linux%'" & _
				" and o.DivisionID = 6" & _
				" order by o.PrimaryProduct,o.ObservationId" 
			rsSIO.open strSQL,cnSIO
			'strOutput = strOutput & "<br/>" & strSQL & "<br />"

			do while not rsSIO.eof
				'compile affected products list for each observation filtered on previously collected product ids
				recordCount = recordCount + 1
				strAffectedProducts = getAffectedProductsFor(rsSIO("ObservationID"), strProductIds, cnSIO, rs)

				if not blnheaderwritten then
					strOutput = strOutput & "<table bgcolor=ivory  border=1 bordercolor=""gainsboro"" cellpadding=2 cellspacing=0>" & _
						"<tr bgcolor=""beige"">" & _
						"<td><b>ObservationID</b></td>" & _
						"<td><b>Priority</b></td>" & _
						"<td><b>Product</b></td>" & _
						"<td><b>Component</b></td>" & _
						"<td><b>Affected&nbsp;Products</b></td>" & _
						"<td><b>State</b></td>" & _
						"<td><b>Days&nbsp;Open</b></td>" & _
						"<td><b>Owner</b></td>" & _
						"<td><b>Summary</b></td>" & _
						"</tr>"
					blnheaderWritten = true
				end if
				strOutput = strOutput & "<tr>" & _
					"<td valign=top><a href=""http://" & Application("Excalibur_ServerName") & "/Excalibur/search/ots/report.asp?txtReportSections=1&txtObservationID=" & _
						ReplaceAndHTMLEncodeFrom(rsSIO("ObservationID")) & """>" & ReplaceAndHTMLEncodeFrom(rsSIO("ObservationID")) & "</a></td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("Priority")) & "</td>" & _
					"<td valign=top>" & replace(ReplaceAndHTMLEncodeFrom(rsSIO("Product")) ," - ","<br>") & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("Deliverable")) & "</td>" & _
					"<td valign=top>" & strAffectedProducts & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("State")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("DaysOpen")) & "</td>" & _
					"<td valign=top nowrap><a href=""mailto:" & ReplaceAndHTMLEncodeFrom(rsSIO("Owner")) & """>" & ReplaceAndHTMLEncodeFrom(rsSIO("Ownername")) & "</a></td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("Summary")) & "</td>" & _
					"</tr>"
				rsSIO.movenext
			loop
			rsSIO.close    
		else
			strOutput = strOutput & _
				"<p>No Commercial Components Found</p>"
		end if

		if blnHeaderWritten then
			strOutput = strOutput & "</table><p><font size=1 face=verdana>Observations Displayed: " & recordCount & "</font></p><hr/><br/>"
		else
			strOutput = strOutput & "<p><font size=1 face=verdana>Observations Displayed: " & recordCount & "</font></p><hr/><br/>"
		end if 

	'	Commercial P1 OTS in "Fix in Progress" or "Fix in Progress – Waiting on Vendor" state which are PAST DUE or have NO Target Date.
		strOutput = strOutput & "<font size=2 face=verdana><a name=""FixPastDue""><b>CMIT NB P1 Observations in &quot;Fix in Progress&quot; or &quot;Fix in Progress – Waiting on Vendor&quot; state which are PAST DUE or have NO Target Date.</b></a> <a href=""#toc"">(top)</a></font><br><br>"
		blnHeaderWritten = false
		recordCount = 0
		if strGroupProductIds <> "" then
			strSQL = _
				"select distinct" & _
					" o.ObservationID" & _
					",o.Priority" & _
					",o.PrimaryProduct as Product" & _
					",dbo.ufn_getCoreTeamNameFromComponentName(o.Component) as CoreTeam" & _
					",o.Component + ' [' + o.ComponentVersion + ']' as Deliverable" & _
					",o.State" & _
					",convert(date,o.TargetDate) as TargetDate" & _
					",o.DaysOpen" & _
					",o.DaysInState" & _
					",o.OwnerName" & _
					",o.Owner" & _
					",o.ShortDescription as Summary" & _
				" from dbo.SI_observation_Report o with (NOLOCK)" & _
				" inner join dbo.Observation oo with (NOLOCK)" & _
				" on oo.Observation_ID = o.ObservationID" & _
				" inner join dbo.Product p with (NOLOCK)" & _
				" on p.Platform_Version_ID = oo.Platform_Version_ID" & _
				" and p.Source_Platform_Version_ID in (" & strGroupProductIds & ")" & _
				" where o.ComponentType not in ('Factory','HW')" & _
				" and o.status <> 'Closed'" & _
				" and o.Priority = 1" & _
				" and o.state in ('Fix in Progress','Fix in Progress - Waiting on Vendor')" & _
				" and o.DivisionID = 6" & _
				" and (" & _
					"[o].[TargetDate] is null" & _
					" or [o].[TargetDate] < dateadd(d, -1, GetDate())" & _
				")" & _
				" order by dbo.ufn_getCoreTeamNameFromComponentName(o.Component),Deliverable,o.observationid"
			rsSIO.open strSQL,cnSIO
			'strOutput = strOutput & "<br/>" & strSQL & "<br />"

			do while not rsSIO.eof
				'compile affected products list for each observation filtered on previously collected product ids
				recordCount = recordCount + 1
				strAffectedProducts = getAffectedProductsFor(rsSIO("ObservationID"), strProductIds, cnSIO, rs)

				if not blnheaderwritten then
					strOutput = strOutput & "<table bgcolor=ivory  border=1 bordercolor=""gainsboro"" cellpadding=2 cellspacing=0>" & _
						"<tr bgcolor=""beige"">" & _
						"<td><b>ObservationID</b></td>" & _
						"<td><b>Priority</b></td>" & _
						"<td><b>Core Team</b></td>" & _
						"<td><b>Component</b></td>" & _
						"<td><b>Product</b></td>" & _
						"<td><b>Affected&nbsp;Products</b></td>" & _
						"<td><b>State</b></td>" & _
						"<td><b>Target&nbsp;Date</b></td>" & _
						"<td><b>Days&nbsp;In&nbsp;State</b></td>" & _
						"<td><b>Days&nbsp;Open</b></td>" & _
						"<td><b>Owner</b></td>" & _
						"<td><b>Summary</b></td>" & _
						"</tr>"
					blnheaderWritten = true
				end if
				strOutput = strOutput & _
					"<tr>" & _
					"<td valign=top>" & _
						"<a href=""http://" & Application("Excalibur_ServerName") & "/Excalibur/search/ots/report.asp?txtReportSections=1&txtObservationID=" & rsSIO("ObservationID") & """>" & rsSIO("ObservationID") & "</a>" & _
					"</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("Priority")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("CoreTeam")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("Deliverable")) & "</td>" & _
					"<td valign=top>" & replace(ReplaceAndHTMLEncodeFrom(rsSIO("Product")) ," - ","<br>") & "</td>" & _
					"<td valign=top>" & strAffectedProducts & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("State")) & "</td>"
				if rsSIO("TargetDate") & "" <> "" then
					strOutput = strOutput &  "<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("TargetDate")) & "</td>"
				else
					strOutput = strOutput &  "<td valign=top>&nbsp;</td>"
				end if
				strOutput = strOutput & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("DaysInState")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("DaysOpen")) & "</td>" & _
					"<td valign=top nowrap><a href=""mailto:" & ReplaceAndHTMLEncodeFrom(rsSIO("Owner")) & """>" & ReplaceAndHTMLEncodeFrom(rsSIO("Ownername")) & "</a></td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("Summary")) & "</td>" & _
					"</tr>"
				rsSIO.movenext
			loop
			rsSIO.close    
		else
			strOutput = strOutput & _
				"<p>No Products found in Product Group " & productGroups & "</p>"
		end if

		if blnHeaderWritten then
			strOutput = strOutput & "</table><p><font size=1 face=verdana>Observations Displayed: " & recordCount & "</font></p><hr/><br/>"
		else
			strOutput = strOutput & "<p><font size=1 face=verdana>Observations Displayed: " & recordCount & "</font></p><hr/><br/>"
		end if 
		
	'	Commercial P1 OTS last modified more than 1 week ago
		strOutput = strOutput & "<font size=2 face=verdana><a name=""Stale""><b>CMIT NB P1 Observations Last Modified over 7 days ago.</b></a> <a href=""#toc"">(top)</a></font><br><br>"
		blnHeaderWritten = false
		recordCount = 0
		if strGroupProductIds <> "" then
			strSQL = _
				"select distinct" & _
					" o.ObservationID" & _
					",dbo.ufn_getCoreTeamNameFromComponentName(o.Component) as CoreTeam" & _
					",o.Priority" & _
					",o.PrimaryProduct as Product" & _
					",o.Component + ' [' + o.ComponentVersion + ']' as Deliverable" & _
					",o.ComponentPMName" & _
					",o.GatingMilestone" & _
					",o.State" & _
					",o.DaysOpen" & _
					",o.DateModified" & _
					",o.DateOpened" & _
					",convert(date,o.TargetDate) as TargetDate" & _
					",o.DaysCurrentOwner" & _
					",o.DaysInState" & _
					",o.OwnerName" & _
					",o.Owner" & _
					",o.ShortDescription as Summary" & _
				" from dbo.SI_observation_Report o with (NOLOCK)" & _
				" inner join dbo.Observation oo with (NOLOCK)" & _
				" on oo.Observation_ID = o.ObservationID" & _
				" inner join dbo.Product p with (NOLOCK)" & _
				" on p.Platform_Version_ID = oo.Platform_Version_ID" & _
				" and p.Source_Platform_Version_ID in (" & strGroupProductIds & ")" & _
				" where o.ComponentType not in ('Factory','HW')" & _
				" and o.status <> 'Closed'" & _
				" and o.Priority = 1" & _
				" and o.DivisionID = 6" & _
				" and (" & _
					" o.state not in ('Fix in Progress','Fix in Progress - Waiting on Vendor','Need Info','Fix Pending Product Verification/Retest')" & _
					" or [TargetDate] <= GetDate()" & _
				" )" & _
				" and DateDiff(dd, o.DateModified, GetDate()) > 7" & _
				" order by dbo.ufn_getCoreTeamNameFromComponentName(o.Component),Product, DaysInState desc,o.observationid"
			rsSIO.open strSQL,cnSIO
			'strOutput = strOutput & "<br/>" & strSQL & "<br />"

			do while not rsSIO.eof
				'compile affected products list for each observation filtered on previously collected product ids
				recordCount = recordCount + 1

				if not blnheaderwritten then
					strOutput = strOutput & "<table bgcolor=ivory  border=1 bordercolor=""gainsboro"" cellpadding=2 cellspacing=0>" & _
						"<tr bgcolor=""beige"">" & _
						"<td><b>ObservationID</b></td>" & _
						"<td><b>Priority</b></td>" & _
						"<td><b>Core Team</b></td>" & _
						"<td><b>Product</b></td>" & _
						"<td><b>Component</b></td>" & _
						"<td><b>Component&nbsp;PM</b></td>" & _
						"<td><b>Gating&nbsp;Milestone</b></td>" & _
						"<td><b>State</b></td>" & _
						"<td><b>Days&nbsp;Open</b></td>" & _
						"<td><b>Date&nbsp;Modified</b></td>" & _
						"<td><b>Date&nbsp;Opened</b></td>" & _
						"<td><b>Target&nbsp;Date</b></td>" & _
						"<td><b>Days&nbsp;With&nbsp;Owner</b></td>" & _
						"<td><b>Days&nbsp;In&nbsp;State</b></td>" & _
						"<td><b>Owner</b></td>" & _
						"<td><b>Summary</b></td>" & _
						"</tr>"
					blnheaderWritten = true
				end if
				strOutput = strOutput & _
					"<tr>" & _
					"<td valign=top>" & _
						"<a href=""http://" & Application("Excalibur_ServerName") & "/Excalibur/search/ots/report.asp?txtReportSections=1&txtObservationID=" & rsSIO("ObservationID") & """>" & rsSIO("ObservationID") & "</a>" & _
					"</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("Priority")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("CoreTeam")) & "</td>" & _
					"<td valign=top>" & replace(ReplaceAndHTMLEncodeFrom(rsSIO("Product")) ," - ","<br>") & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("Deliverable")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("ComponentPMName")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("GatingMilestone")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("State")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("DaysOpen")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("DateModified")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("DateOpened")) & "</td>"
				if rsSIO("TargetDate") & "" <> "" then
					strOutput = strOutput &  "<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("TargetDate")) & "</td>"
				else
					strOutput = strOutput &  "<td valign=top>&nbsp;</td>"
				end if
				strOutput = strOutput & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("DaysCurrentOwner")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("DaysInState")) & "</td>" & _
					"<td valign=top nowrap><a href=""mailto:" & ReplaceAndHTMLEncodeFrom(rsSIO("Owner")) & """>" & ReplaceAndHTMLEncodeFrom(rsSIO("Ownername")) & "</a></td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("Summary")) & "</td>" & _
					"</tr>"
				rsSIO.movenext
			loop
			rsSIO.close    
		else
			strOutput = strOutput & _
				"<p>No Products found in Product Group " & productGroups & "</p>"
		end if

		if blnHeaderWritten then
			strOutput = strOutput & "</table><p><font size=1 face=verdana>Observations Displayed: " & recordCount & "</font></p>"
		else
			strOutput = strOutput & "<p><font size=1 face=verdana>Observations Displayed: " & recordCount & "</font></p>"
		end if 

	'	Commercial P1 OTS in Need Info/Retest State > 24Hrs
		strOutput = strOutput & "<font size=2 face=verdana><a name=""Rsvp""><b>CMIT NB P1 Observations in Need Info/Retest State &gt; 24Hrs.</b></a> <a href=""#toc"">(top)</a></font><br><br>"
		blnHeaderWritten = false
		recordCount = 0
		if strGroupProductIds <> "" then
			strSQL = _
				"select distinct" & _
					" o.ObservationID" & _
					",dbo.ufn_getCoreTeamNameFromComponentName(o.Component) as CoreTeam" & _
					",o.Priority" & _
					",o.PrimaryProduct as Product" & _
					",o.Component + ' [' + o.ComponentVersion + ']' as Deliverable" & _
					",o.ComponentPMName" & _
					",o.GatingMilestone" & _
					",o.State" & _
					",o.DaysOpen" & _
					",o.DateModified" & _
					",o.DateOpened" & _
					",convert(date,o.TargetDate) as TargetDate" & _
					",o.DaysCurrentOwner" & _
					",o.DaysInState" & _
					",o.OwnerName" & _
					",o.Owner" & _
					",o.ShortDescription as Summary" & _
				" from dbo.SI_observation_Report o with (NOLOCK)" & _
				" inner join dbo.Observation oo with (NOLOCK)" & _
				" on oo.Observation_ID = o.ObservationID" & _
				" inner join dbo.Product p with (NOLOCK)" & _
				" on p.Platform_Version_ID = oo.Platform_Version_ID" & _
				" and p.Source_Platform_Version_ID in (" & strGroupProductIds & ")" & _
				" where o.ComponentType not in ('Factory','HW')" & _
				" and o.status <> 'Closed'" & _
				" and o.Priority = 1" & _
				" and o.DivisionID = 6" & _
				" and o.State in ('Need Info', 'Fix Pending Product Verification/Retest')" & _
				" and o.DaysInState >= 1" & _
				" order by dbo.ufn_getCoreTeamNameFromComponentName(o.Component),Product, DaysInState desc,o.observationid"
			rsSIO.open strSQL,cnSIO
			'strOutput = strOutput & "<br/>" & strSQL & "<br />"

			do while not rsSIO.eof
				'compile affected products list for each observation filtered on previously collected product ids
				recordCount = recordCount + 1

				if not blnheaderwritten then
					strOutput = strOutput & "<table bgcolor=ivory  border=1 bordercolor=""gainsboro"" cellpadding=2 cellspacing=0>" & _
						"<tr bgcolor=""beige"">" & _
						"<td><b>ObservationID</b></td>" & _
						"<td><b>Priority</b></td>" & _
						"<td><b>Core Team</b></td>" & _
						"<td><b>Product</b></td>" & _
						"<td><b>Component</b></td>" & _
						"<td><b>Component&nbsp;PM</b></td>" & _
						"<td><b>Gating&nbsp;Milestone</b></td>" & _
						"<td><b>State</b></td>" & _
						"<td><b>Days&nbsp;Open</b></td>" & _
						"<td><b>Date&nbsp;Modified</b></td>" & _
						"<td><b>Date&nbsp;Opened</b></td>" & _
						"<td><b>Target&nbsp;Date</b></td>" & _
						"<td><b>Days&nbsp;With&nbsp;Owner</b></td>" & _
						"<td><b>Days&nbsp;In&nbsp;State</b></td>" & _
						"<td><b>Owner</b></td>" & _
						"<td><b>Summary</b></td>" & _
						"</tr>"
					blnheaderWritten = true
				end if
				strOutput = strOutput & _
					"<tr>" & _
					"<td valign=top>" & _
						"<a href=""http://" & Application("Excalibur_ServerName") & "/Excalibur/search/ots/report.asp?txtReportSections=1&txtObservationID=" & rsSIO("ObservationID") & """>" & rsSIO("ObservationID") & "</a>" & _
					"</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("Priority")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("CoreTeam")) & "</td>" & _
					"<td valign=top>" & replace(ReplaceAndHTMLEncodeFrom(rsSIO("Product")) ," - ","<br>") & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("Deliverable")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("ComponentPMName")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("GatingMilestone")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("State")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("DaysOpen")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("DateModified")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("DateOpened")) & "</td>"
				if rsSIO("TargetDate") & "" <> "" then
					strOutput = strOutput &  "<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("TargetDate")) & "</td>"
				else
					strOutput = strOutput &  "<td valign=top>&nbsp;</td>"
				end if
				strOutput = strOutput & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("DaysCurrentOwner")) & "</td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("DaysInState")) & "</td>" & _
					"<td valign=top nowrap><a href=""mailto:" & ReplaceAndHTMLEncodeFrom(rsSIO("Owner")) & """>" & ReplaceAndHTMLEncodeFrom(rsSIO("Ownername")) & "</a></td>" & _
					"<td valign=top>" & ReplaceAndHTMLEncodeFrom(rsSIO("Summary")) & "</td>" & _
					"</tr>"
				rsSIO.movenext
			loop
			rsSIO.close    
		else
			strOutput = strOutput & _
				"<p>No Products found in Product Group " & productGroups & "</p>"
		end if

		if blnHeaderWritten then
			strOutput = strOutput & "</table><p><font size=1 face=verdana>Observations Displayed: " & recordCount & "</font></p>"
		else
			strOutput = strOutput & "<p><font size=1 face=verdana>Observations Displayed: " & recordCount & "</font></p>"
		end if 
		
		set rsSIO = nothing
		set rs = nothing
		cnExcalibur.Close
		set cnExcalibur = nothing
		cnSIO.Close
		set cnSIO = nothing
		compileNagEmailFor = strOutput
	end function
	
	function getCommercialDeliverableIds(connection, rs)
		dim deliverableIds,first,sql
		deliverableIds = ""
		'get list of IDs for deliverables in commercial images
		sql = _
			"select distinct v.id" & _
			" from dbo.DeliverableVersion v with (NOLOCK)" & _
			" inner join dbo.Product_Deliverable pd with (NOLOCK)" & _
			" on pd.DeliverableVersionId = v.id" & _
			" and pd.InImage = 1" & _
			" where v.CommercialReleaseStatus > 0" & _
			" and v.Location like '%Workflow Complete%'"
		rs.open sql,connection
		first = true
		do while not rs.eof
			if first then
				first = false
			else
				deliverableIds = deliverableIds & ","
			end if
			deliverableIds = deliverableIds & rs("id")
			rs.movenext
		loop
		rs.close
		getCommercialDeliverableIds = deliverableIds
	end function
	
	function getActiveCommercialProductIds(connection, rs)
		dim first,productIds,sql
		productIds = ""
		'get list of active commercial product ids
		sql = _
			"Select distinct id" & _
			" from dbo.ProductVersion p with (NOLOCK)" & _
			" where p.active = 1" & _
			" and p.devcenter <> 2"
		rs.open sql,connection
		first = true
		do while not rs.eof
			if first then
				first = false
			else
				productIds = productIds & ","
			end if
			productIds = productIds & rs("ID")
			rs.movenext
		loop
		rs.close
		getActiveCommercialProductIds = productIds
	end function

	function getProductIdsFor(productGroupSet, connection, rs)
		dim first,productIds,sql
		productIds = ""
		'get list of product ids from product groups
		sql = _
			"select distinct [pv].[ID]" & _
			" from [dbo].[ProductVersion] [pv] with (NOLOCK)" & _
			" inner join [dbo].[Product_Program] [pp] with (NOLOCK)" & _
			" on [pp].[ProductVersionID] = [pv].[ID]" & _
			" inner join [dbo].[Program] [p] with (NOLOCK)" & _
			" on [p].[ID] = [pp].[ProgramID]" & _
			" and [p].[Name] in (" & productGroupSet & ")"
		rs.open sql,connection
		first = true
		do while not rs.eof
			if first then
				first = false
			else
				productIds = productIds & ","
			end if
			productIds = productIds & rs("ID")
			rs.movenext
		loop
		rs.close
		getProductIdsFor = productIds
	end function
	
	function getAffectedProductsFor(observationId, productIdsFilter, connection, rs)
		dim first,affectedProducts,sql
		affectedProducts = ""
		sql = _
			"Select a.*" & _
			" from dbo.AffectedProducts a with (NOLOCK)" & _
			" inner join Product p with (NOLOCK)" & _
			" on p.Platform_Version_ID = a.Platform_Version_ID" & _
			" and p.Source_Platform_Version_ID in (" & productIdsFilter & ")" & _
			" where a.Observation_ID = " & observationId & _
			" and a.Affected_State_Name in ('test required','waiver requested','affected','Untested*')" & _
			" order by Platform_Cycle_Version"
		rs.open sql,connection
		first = true
		do while not rs.eof
			if first then
				first = false
			else
				affectedProducts = affectedProducts & ", "
			end if
			affectedProducts = affectedProducts & rs("Platform_Cycle_Version")
			rs.movenext
		loop
		rs.close
		if trim(affectedProducts) = "" then
			affectedProducts = "&nbsp;"
		else
			affectedProducts = ReplaceAndHTMLEncodeFrom(affectedProducts)
		end if
		getAffectedProductsFor = affectedProducts
	end function

	sub sendNagEmail(subjectTitle, sender, output)
		dim subject
		subject = "[Update Required!] " & subjectTitle & " - CMIT Component PMs/Developers - HP Restricted"
		if trim(sender & "") = "" then
			sender = "pony.ma@hp.com"
		end if
		if request("auto") = "1" then
			dim oMessage
			set oMessage = new EmailWrapper 
			if output <> "" then
				oMessage.From = sender
				oMessage.To= "bpcscmitnbsepms@hp.com;bnb.ctls.worldwide@hp.com;mondshine.developers@hp.com;mondshine.swpms@hp.com;mondshine.ctls@hp.com;asg_delivery@hp.com;cmitcommtdcswgroupall@hp.com;CMIT.NBMIT@hp.com;CMIT.NBSFT@hp.com;COMMsPMs@hp.com;JonLiuSWPMs@hp.com"
				'oMessage.To= "matt.hamilton@hp.com"
				oMessage.BCC = "matt.hamilton@hp.com"
				oMessage.Subject = subject
				oMessage.HTMLBody = output 
				oMessage.Send 
			else
				oMessage.From = sender
				oMessage.To= "pony.ma@hp.com;bpcscmitnbsepms@hp.com"
				oMessage.BCC = "matt.hamilton@hp.com"
				oMessage.Subject = subject
				oMessage.HTMLBody = "<br/>Output is blank, something is terribly wrong<br/>" 
				oMessage.Send 
			end if
			set oMessage = nothing 	    
		else
			response.write _
				"From: " & sender & "<br/>" & _
				"To: bpcscmitnbsepms@hp.com;bnb.ctls.worldwide@hp.com;mondshine.developers@hp.com;mondshine.swpms@hp.com;mondshine.ctls@hp.com;asg_delivery@hp.com;cmitcommtdcswgroupall@hp.com;CMIT.NBMIT@hp.com;CMIT.NBSFT@hp.com;COMMsPMs@hp.com;JonLiuSWPMs@hp.com<br/>" & _
				"Bcc: matt.hamilton@hp.com<br/>" & _
				"Subject: " & subject & "<br/><hr/><br/>" & _
				output
		end if
	end sub
	%>
	<input style="display: none" id="txtAuto" type="text" value="<%=request("auto")%>" />
</body>
</html>




