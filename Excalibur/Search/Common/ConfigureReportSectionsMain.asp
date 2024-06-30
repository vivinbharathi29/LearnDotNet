<%@  language="VBScript" %>
<%
	Response.Buffer = True
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"
%>
<html>
<head>
	<title>Configure Report</title>
	<script id="clientEventHandlersJS" type="text/javascript">
<!--
	function isDate(sDate) {
		var MyDate = new Date(sDate);
		if (MyDate.toString() == "NaN" || MyDate.toString() == "Invalid Date")
			return false;
		else
			return true;
	}

	function cmdCancel_click() {
		window.parent.close();
	}

	function isInteger(n) {
		if (isNaN(n))
			return false;
		if (n - Math.floor(n)) return false; return true;
	}

	function cmdOK_click() {
		//Validations

		if (cboEndDate.selectedIndex != 0) {
			if (txtEndDate.value == "") {
				alert("Date must be supplied when you choose a custom end date.");
				txtEndDate.focus();
				return;
			}
			else if (!isDate(txtEndDate.value)) {
				alert("End Date must be a valid date.");
				txtEndDate.focus();
				return;
			}
		}

		if (cboGridLine.selectedIndex != 0) {
			if (txtGridlines.value == "") {
				alert("Number of items to display per gridline must be entered if Custom gridline interval is selected.");
				txtGridlines.focus();
				return;
			}
			else if (!isInteger(txtGridlines.value)) {
				alert("Gridline interval must be a positive integer.");
				txtGridlines.focus();
				return;
			}
			else if (parseInt(txtGridlines.value) < 0) {
				alert("Gridline interval must be a positive integer.");
				txtGridlines.focus();
				return;
			}
		}

		//Build output string
		var strOutput = "";

		// field 0: end date
		if (cboEndDate.selectedIndex == 0)
			strOutput = "^";
		else
			strOutput = txtEndDate.value + "^";

		// field 1: weeks
		if (cboWeeks.options[cboWeeks.selectedIndex].text == "16")
			strOutput = strOutput + "^";
		else
			strOutput = strOutput + cboWeeks.options[cboWeeks.selectedIndex].text + "^";

		// field 2: legend
		if (cboLegend.options[cboLegend.selectedIndex].value == "1")
			strOutput = strOutput + "^";
		else
			strOutput = strOutput + cboLegend.options[cboLegend.selectedIndex].value + "^";

		// field 3: activity
		if (txtType.value == "3" || (txtType.value == "1" && cboActivity.options[cboActivity.selectedIndex].value == "1") || txtType.value == "2" && cboActivity.options[cboActivity.selectedIndex].value == "0")
			strOutput = strOutput + "^";
		else
			strOutput = strOutput + cboActivity.options[cboActivity.selectedIndex].value + "^";

		// field 4: title
		if (txtTitle.value == "Weekly Observation BackLog")
			strOutput = strOutput + "^";
		else
			strOutput = strOutput + txtTitle.value.replace(/\|/g, "_").replace(/\^/g, "_") + "^";

		// field 5: width
		if (cboWidth.options[cboWidth.selectedIndex].value == "2")
			strOutput = strOutput + "^";
		else
			strOutput = strOutput + cboWidth.options[cboWidth.selectedIndex].value + "^";

		// field 6: height
		if (cboHeight.options[cboHeight.selectedIndex].value == "2")
			strOutput = strOutput + "^";
		else
			strOutput = strOutput + cboHeight.options[cboHeight.selectedIndex].value + "^";

		// field 7: gridline
		if (cboGridLine.options[cboGridLine.selectedIndex].value == "0")
			strOutput = strOutput + "^";
		else
			strOutput = strOutput + txtGridlines.value + "^";

		// field 8: series
		if (txtType.value != "1")
			strOutput = strOutput + cboSeries.options[cboSeries.selectedIndex].value + "^";
		else
			strOutput = strOutput + "^";

		// field 9: group by
		if (txtType.value == "1" || (txtType.value != "2" && cboGroupBy.options[cboGroupBy.selectedIndex].value == "Priority"))
			strOutput = strOutput + "^";
		else
			strOutput = strOutput + cboGroupBy.options[cboGroupBy.selectedIndex].value + "^";

		// field 10: chart type
		if (cboChartType.options[cboChartType.selectedIndex].value == "LineMarkers")
			strOutput = strOutput + "^";
		else
			strOutput = strOutput + cboChartType.options[cboChartType.selectedIndex].value + "^";

		// field 11: milestone date 1 label
		strOutput = strOutput + milestoneLabel1.value + "^";

		// field 12: milestone date 1
		strOutput = strOutput + milestoneDate1.value + "^";

		// field 13: milestone date 2 label
		strOutput = strOutput + milestoneLabel2.value + "^";

		// field 14: milestone date 1 label
		strOutput = strOutput + milestoneDate2.value + "^";

		// field 15: show total backlog
		if (txtType.value == "3" || (txtType.value == "1" && cboBacklog.options[cboBacklog.selectedIndex].value == "1") || txtType.value == "2" && cboBacklog.options[cboBacklog.selectedIndex].value == "0")
			strOutput = strOutput + "^";
		else
			strOutput = strOutput + cboBacklog.options[cboBacklog.selectedIndex].value + "^";

		if (strOutput == "^^^^^^^^^^^^^^^^")
			strOutput = "";

		txtParams.value = strOutput;
		if (navigator.appName != "Microsoft Internet Explorer" && navigator.appName != "Internet Explorer" && navigator.appName != "IE")
			if (typeof (window.parent.opener) != "undefined") {
				if (strOutput == "")
					window.parent.opener.document.all("Node" + txtID.value).style.color = "black";
				else
					window.parent.opener.document.all("Node" + txtID.value).style.color = "green";

				window.parent.opener.document.all("Params" + txtID.value).innerText = strOutput;
			}
		window.returnValue = txtParams.value;
		window.parent.close();
	}

	function cmdChooseDate_onclick(DateFieldID) {
		var strID;
		strID = window.showModalDialog("../../mobilese/today/caldraw1.asp", DateFieldID.value, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
		if (typeof (strID) != "undefined") {
			DateFieldID.value = strID;
			if (DateFieldID.ClientID == txtEndDate.ClientID)
				cboWeeks_onchange();
		}
	}

	function cboSeries_onchange() {
		if (cboSeries.options[cboSeries.selectedIndex].text > 15)
			SeriesWarning.style.display = "";
		else
			SeriesWarning.style.display = "none";
	}

	function cboEndDate_onchange() {
		if (cboEndDate.selectedIndex == 0)
			spnEndDate.style.display = "none";
		else {
			spnEndDate.style.display = "";
			txtEndDate.focus();
		}
		cboWeeks_onchange();
	}

	function cboGridLine_onchange() {
		if (cboGridLine.selectedIndex == 0)
			spnGridline.style.display = "none";
		else {
			spnGridline.style.display = "";
			txtGridlines.focus();
		}
	}

	function window_onload() {
		txtTitle.focus();
	}

	function txtDate_onchange(control) {
		if (control.value != "" && !isDate(control.value)) {
			alert("You must specify a valid date in the Date field.");
			control.focus();
		}
	}

	function cboWeeks_onchange() {
		var MinDate = new Date("1/1/2011");
		var StartDate;

		if (cboEndDate.selectedIndex != 0 && txtEndDate.value != "" && !isDate(txtEndDate.value)) {
			alert("You must specify a valid date in the End Date field.");
			txtEndDate.focus();
			return;
		}
		else if (cboEndDate.selectedIndex == 0 || txtEndDate.value == "" || !isDate(txtEndDate.value)) {
			StartDate = new Date();
			if (txtID.value >= 21) {
				StartDate.setDate(StartDate.getDate());
			} else {
				StartDate.setDate(StartDate.getDate() - 1);
			}
		}
		else
			StartDate = new Date(txtEndDate.value);

		StartDate.setDate(StartDate.getDate() - (7 * (parseInt(cboWeeks.options[cboWeeks.selectedIndex].text) - 1)));

		if (StartDate < MinDate) {
			StartDate = MinDate;
			spnStartWarning.style.display = "";
		}
		else
			spnStartWarning.style.display = "none";

		spnStart.innerText = StartDate.format("mm/dd/yyyy");
	}

	function txtEndDate_lostfocus() {
		cboWeeks_onchange();
	}

	/*
	* Date Format 1.2.3
	* (c) 2007-2009 Steven Levithan <stevenlevithan.com>
	* MIT license
	*
	* Includes enhancements by Scott Trenda <scott.trenda.net>
	* and Kris Kowal <cixar.com/~kris.kowal/>
	*
	* Accepts a date, a mask, or a date and a mask.
	* Returns a formatted version of the given date.
	* The date defaults to the current date/time.
	* The mask defaults to dateFormat.masks.default.
	*/

	var dateFormat = function () {
		var token = /d{1,4}|m{1,4}|yy(?:yy)?|([HhMsTt])\1?|[LloSZ]|"[^"]*"|'[^']*'/g,
	timezone = /\b(?:[PMCEA][SDP]T|(?:Pacific|Mountain|Central|Eastern|Atlantic) (?:Standard|Daylight|Prevailing) Time|(?:GMT|UTC)(?:[-+]\d{4})?)\b/g,
	timezoneClip = /[^-+\dA-Z]/g,
	pad = function (val, len) {
		val = String(val);
		len = len || 2;
		while (val.length < len) val = "0" + val;
		return val;
	};

		// Regexes and supporting functions are cached through closure
		return function (date, mask, utc) {
			var dF = dateFormat;

			// You can't provide utc if you skip other args (use the "UTC:" mask prefix)
			if (arguments.length == 1 && Object.prototype.toString.call(date) == "[object String]" && !/\d/.test(date)) {
				mask = date;
				date = undefined;
			}

			// Passing date through Date applies Date.parse, if necessary
			date = date ? new Date(date) : new Date;
			if (isNaN(date)) throw SyntaxError("invalid date");

			mask = String(dF.masks[mask] || mask || dF.masks["default"]);

			// Allow setting the utc argument via the mask
			if (mask.slice(0, 4) == "UTC:") {
				mask = mask.slice(4);
				utc = true;
			}

			var _ = utc ? "getUTC" : "get",
		d = date[_ + "Date"](),
		D = date[_ + "Day"](),
		m = date[_ + "Month"](),
		y = date[_ + "FullYear"](),
		H = date[_ + "Hours"](),
		M = date[_ + "Minutes"](),
		s = date[_ + "Seconds"](),
		L = date[_ + "Milliseconds"](),
		o = utc ? 0 : date.getTimezoneOffset(),
		flags = {
			d: d,
			dd: pad(d),
			ddd: dF.i18n.dayNames[D],
			dddd: dF.i18n.dayNames[D + 7],
			m: m + 1,
			mm: pad(m + 1),
			mmm: dF.i18n.monthNames[m],
			mmmm: dF.i18n.monthNames[m + 12],
			yy: String(y).slice(2),
			yyyy: y,
			h: H % 12 || 12,
			hh: pad(H % 12 || 12),
			H: H,
			HH: pad(H),
			M: M,
			MM: pad(M),
			s: s,
			ss: pad(s),
			l: pad(L, 3),
			L: pad(L > 99 ? Math.round(L / 10) : L),
			t: H < 12 ? "a" : "p",
			tt: H < 12 ? "am" : "pm",
			T: H < 12 ? "A" : "P",
			TT: H < 12 ? "AM" : "PM",
			Z: utc ? "UTC" : (String(date).match(timezone) || [""]).pop().replace(timezoneClip, ""),
			o: (o > 0 ? "-" : "+") + pad(Math.floor(Math.abs(o) / 60) * 100 + Math.abs(o) % 60, 4),
			S: ["th", "st", "nd", "rd"][d % 10 > 3 ? 0 : (d % 100 - d % 10 != 10) * d % 10]
		};

			return mask.replace(token, function ($0) {
				return $0 in flags ? flags[$0] : $0.slice(1, $0.length - 1);
			});
		};
	}();

	// Some common format strings
	dateFormat.masks = {
		"default": "ddd mmm dd yyyy HH:MM:ss",
		shortDate: "m/d/yy",
		mediumDate: "mmm d, yyyy",
		longDate: "mmmm d, yyyy",
		fullDate: "dddd, mmmm d, yyyy",
		shortTime: "h:MM TT",
		mediumTime: "h:MM:ss TT",
		longTime: "h:MM:ss TT Z",
		isoDate: "yyyy-mm-dd",
		isoTime: "HH:MM:ss",
		isoDateTime: "yyyy-mm-dd'T'HH:MM:ss",
		isoUtcDateTime: "UTC:yyyy-mm-dd'T'HH:MM:ss'Z'"
	};

	// Internationalization strings
	dateFormat.i18n = {
		dayNames: [
	"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat",
	"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
		],
		monthNames: [
	"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
	"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"
		]
	};

	// For convenience...
	Date.prototype.format = function (mask, utc) {
		return dateFormat(this, mask, utc);
	};

	//-->
	</script>
	<style type="text/css">
		h1
		{
			font-family: Verdana;
			font-size: small;
		}
		td
		{
			font-family: Verdana;
			font-size: xx-small;
		}
		body
		{
			background-color: #e6e6fa; /* lavender */
			font-family: Verdana;
			font-size: xx-small;
		}
	</style>
</head>
<body onload="window_onload();">
	<%
	dim strWeeks
	dim strGroupBy
	dim strGroupByValue
	dim GroupByArray
	dim strItem
	dim SeriesArray
	dim strSeries
	dim ShowGridLine
	dim ShowEndDate
	dim ParameterArray
	dim EndDateIndex
	dim strLegend
	dim strTitle
	dim strWidth
	dim strHeight
	dim strGridlines
	dim WidthArray
	dim HeightArray
	dim strChartType
	dim strMilestoneDate1
	dim strMilestoneDate2
	dim strMilestoneDateLabel1
	dim strMilestoneDateLabel2
	dim showField

	strChartType = "LineMarkers"
	strMilestoneDateLabel1 = "Milestone 1 Date"
	strMilestoneDateLabel2 = "Milestone 2 Date"

	if trim(request("txtParams")) = "" then
		ParameterArray = split("^^^^^^^^^","^")
	else
		ParameterArray = split(request("txtParams"),"^")
	end if
	for i = 0 to ubound(ParameterArray)
		select case i
			case 0 'Process End Date
				if trim(ParameterArray(i)) = "" then
					EndDateIndex = 0
					if request("txtID") >= 21 then
						strEndDate = formatdatetime(now(), vbshortdate)
					else
						strEndDate = formatdatetime(now()-1, vbshortdate)
					end if
					ShowEndDate = "none" 'Manual Entry Field
'					strStartDate = formatdatetime((now()-1) - (7*(cint(strWeeks)-1)),vbshortdate)
				else
					EndDateIndex = 1
					strEndDate = trim(cdate(ParameterArray(i)))
					ShowEndDate = "" 'Manual Entry Field
'					strStartDate = formatdatetime(cdate(strEndDate) - (7*(cint(strWeeks)-1)),vbshortdate)
				end if
'
'				'Calculate the StartDate and show/hide as necessary
'				if datediff("d",strStartDate ,"1/1/2011") > 0 then
'					ShowStartWarning=""
'					strStartDate = "1/1/2011"
'				else
'					ShowStartWarning="none"
'				end if
			case 1 'Process Weeks
				if trim(ParameterArray(i)) = "" then
					strWeeks=16
				else
					strWeeks= clng(ParameterArray(i))
				end if
			case 2 'Process Legend
				if trim(ParameterArray(i)) = "" then
					strLegend= "1"
				else
					strLegend= trim(ParameterArray(i))
				end if
			case 3 'Process Activity Graph
				if trim(ParameterArray(i)) = "" then
					if request("TypeID") = "1" then
						strActivity = "1"
					else
						strActivity = "0"
					end if
				else
					strActivity = trim(ParameterArray(i))
				end if
			case 4 'Process Title
				if trim(ParameterArray(i)) <> "" then
					strTitle = trim(ParameterArray(i))
				end if
			case 5 'Process Width
				if trim(ParameterArray(i)) = "" then
					strWidth="2"
				else
					strWidth=trim(ParameterArray(i))
				end if
			case 6 'Process Height
				if trim(ParameterArray(i)) = "" then
					strHeight="2"
				else
					strHeight=trim(ParameterArray(i))
				end if
			case 7 'Process GridLine
				if trim(ParameterArray(i)) = "" then
					ShowGridLine = "none"
					strGridLine = ""
				else
					ShowGridLine = ""
					strGridLine = clng(trim(ParameterArray(i)))
				end if
			case 8 'Process Series
				if trim(ParameterArray(i)) = "" then
					strSeries = "15"
				else
					strSeries = clng(trim(ParameterArray(i)))
				end if
			case 9 'Process GroupBy
				if trim(ParameterArray(i)) = "" then
					if request("txtID") >= 25 then
						strGroupBy = "Component"
					else
						strGroupBy = "Priority"
					end if
				else
					strGroupBy = trim(ParameterArray(i))
				end if
			case 10 'Process Chart Type
				if trim(ParameterArray(i)) = "" then
					strChartType = "LineMarkers"
				else
					strChartType = trim(ParameterArray(i))
				end if
			case 11 'Process Milestone Date Label 1
				if trim(ParameterArray(i)) = "" then
					strMilestoneDateLabel1 = "Milestone 1 Date"
				else
					strMilestoneDateLabel1 = trim(ParameterArray(i))
				end if
			case 12 'Process Milestone Date 1
				if trim(ParameterArray(i)) = "" then
					strMilestoneDate1 = ""
				else
					strMilestoneDate1 = cdate(trim(ParameterArray(i)))
				end if
			case 13 'Process Milestone Date Label 2
				if trim(ParameterArray(i)) = "" then
					strMilestoneDateLabel2 = "Milestone 2 Date"
				else
					strMilestoneDateLabel2 = trim(ParameterArray(i))
				end if
			case 14 'Process Milestone Date 2
				if trim(ParameterArray(i)) = "" then
					strMilestoneDate2 = ""
				else
					strMilestoneDate2 = cdate(trim(ParameterArray(i)))
				end if
			case 15 'Include backlog total on Graph
				if trim(ParameterArray(i)) = "" then
					if request("TypeID") = "1" then
						strBacklog = "1"
					else
						strBacklog = "0"
					end if
				else
					strBacklog = trim(ParameterArray(i))
				end if
		end select
	next
	strStartDate = formatdatetime(cdate(strEndDate) - (7*(cint(strWeeks)-1)),vbshortdate)
	'Calculate the StartDate and show/hide as necessary
	if datediff("d",strStartDate ,"1/1/2011") > 0 then
		ShowStartWarning=""
		strStartDate = "1/1/2011"
	else
		ShowStartWarning="none"
	end if

	if trim(strTitle) = "" then
		strTitle = getSectionTitle(request("txtID"))
	end if

	response.write "<h1>Configure Graph</h1>"
'	if request("TypeID") = "1" then
'		response.write "<h1>Configure Backlog Graph</h1>"
'	elseif request("TypeID") = "2" then
'		response.write "<h1>Configure Backlog Group Graph</h1>"
'	else
'		response.write "Unable to determine the type of field."
'		respone.write "<hr><table><tr><td><input id=""cmdClose"" type=""button"" value=""Cancel"" onclick=""cmdCancel_click();""></td></tr></table>"
'	end if

	if request("TypeID") = "1" or request("TypeID") = "2" or request("TypeID") = "3" then
	%>
	<table style="border-width: 1px; width: 100%; background-color: #E8E8E8; border-color: #a9a9a9" <%'darkgray%> border="1" cellpadding="2" cellspacing="0">
		<tr>
			<td style="font-weight: bold">Title:&nbsp;&nbsp;&nbsp;
			</td>
			<td style="width: 100%">
				<input style="width: 100%" id="txtTitle" maxlength="50" type="text" value="<%=strTitle%>" />
			</td>
		</tr>
		<%
		if request("TypeID") <> "1" then
			if request("TypeID") = "2" then
				GroupByArray = split("Approver,Approver Group,Component,Component PM,Component PM Group,Component Test Lead,Component Test Lead Group,Component Type,Core Team,Developer,Developer Group,Feature,Frequency,Gating Milestone,Originator,Originator Group,Owner,Owner Group,Primary Product,Priority,Product Family,Product PM,Product PM Group,Product Test Lead,Product Test Lead Group,State,Status,Sub System,Tester,Tester Group",",")
			else
				GroupByArray = split("Approver,Approver Group,Approver Manager,Component,Component PM,Component PM Group,Component PM Manager,Component Test Lead,Component Test Lead Group,Component Test Lead Manager,Component Type,Core Team,Developer,Developer Group,Developer Manager,Feature,Frequency,Gating Milestone,Originator,Originator Group,Originator Manager,Owner,Owner Group,Owner Manager,Primary Product,Priority,Product Family,Product PM,Product PM Group,Product PM Manager,Product Test Lead,Product Test Lead Group,Product Test Lead Manager,State,Status,Sub System,Tester,Tester Group,Tester Manager",",")
			end if

			showField = "inline"
			if request("txtID") = "21" then
				showField = "none"
			end if
		%>
		<tr style="display: <%=showField%>">
			<td style="white-space: nowrap; font-weight: bold">Group By:&nbsp;&nbsp;&nbsp;
			</td>
			<td style="white-space: nowrap">
				<select id="cboGroupBy" style="width: 180px">
					<%
			for each strItem in GroupByArray
				if trim(strItem) = "Tester" or trim(strItem) = "Approver" or trim(strItem) = "Developer" or trim(strItem) = "Component PM" or trim(strItem) = "Component Lest Lead" or trim(strItem) = "Product Test Lead" or trim(strItem) = "Product PM" or trim(strItem) = "Owner" or trim(strItem) = "Originator" then
					strItemValue = replace(strItem," ","") & "Name"
				else
					strItemValue = replace(strItem," ","")
				end if

				if trim(strGroupBy) = trim(strItemValue) then
					response.write "<option selected=""selected"" value=""" & strItemValue & """>" & strItem & "</option>"
				else
					response.write "<option value=""" & strItemValue & """>" & strItem & "</option>"
				end if
			next
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td style="white-space: nowrap">
				<b>Series&nbsp;to&nbsp;Display:&nbsp;&nbsp;</b>
			</td>
			<td style="white-space: nowrap">Top
				<select id="cboSeries" style="width: 50px" onchange="cboSeries_onchange();">
					<%
			if clng(strSeries) > 15 then
				ShowSeriesWarning = ""
			else
				ShowSeriesWarning = "none"
			end if
			SeriesArray = split("1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20",",")
			for each strItem in SeriesArray
				if trim(strSeries) = trim(strItem) then
					if trim(strItem) = "15" then
						response.write "<option selected=""selected"" value="""">" & strItem & "</option>"
					else
						response.write "<option selected=""selected"" value=""" & strItem & """>" & strItem & "</option>"
					end if
				else
					if trim(strItem) = "15" then
						response.write "<option value="""">" & strItem & "</option>"
					else
						response.write "<option value=""" & strItem & """>" & strItem & "</option>"
					end if
				end if
			next
					%>
				</select>&nbsp;series.&nbsp;<span style="display: <%=ShowSeriesWarning%>; color: Green; font-size: xx-small; font-family: Verdana" id="SeriesWarning">&nbsp;Note: Some series names may not fit in Legend.</span>
			</td>
		</tr>
		<%
		end if
		showField = "inline"
		if request("TypeID") = "3" then
			showField = "none"
		end if
		dim strDefaultEndDay
		if request("txtID") = "21" then
			strDefaultEndDay = "Today"
		else
			strDefaultEndDay = "Yesterday"
		end if
		%>
		<tr style="display: <%=showField%>">
			<td style="white-space: nowrap; font-weight: bold">End Date:&nbsp;&nbsp;&nbsp;
			</td>
			<td style="white-space: nowrap">
				<select id="cboEndDate" onchange="cboEndDate_onchange();" style="width: 140px">
					<%if EndDateIndex=0 then%>
					<option selected="selected"><%=strDefaultEndDay%></option>
					<option>Custom End Date</option>
					<%else%>
					<option><%=strDefaultEndDay%></option>
					<option selected="selected">Custom End Date</option>
					<%end if%>
				</select>&nbsp;&nbsp;&nbsp; <span id="spnEndDate" style="display: <%=ShowEndDate%>">&nbsp;<input onblur="javascript: txtEndDate_lostfocus();" style="width: 80" id="txtEndDate" type="text" value="<%=strEndDate%>" />&nbsp;<a href="javascript: cmdChooseDate_onclick(txtEndDate)"><img style="margin-bottom: -4px" id="picTarget" src="../../mobilese/today/images/calendar.gif" alt="Choose Date" width="26" height="21" /></a></span>
			</td>
		</tr>
		<%
		showField = "none"
		if request("txtID") = "21" then
			showField = "inline"
		end if
		%>
		<tr style="display: <%=showField%>">
			<td style="white-space: nowrap; font-weight: bold">
				<input style="width: 100%" id="milestoneLabel1" maxlength="50" type="text" value="<%=strMilestoneDateLabel1%>" />
			</td>
			<td style="white-space: nowrap">
				<span id="Span1" style="display: <%=showField%>">&nbsp;
					<input style="width: auto" id="milestoneDate1" type="text" value="<%=strMilestoneDate1%>" onchange="txtDate_onchange(this);" />&nbsp;
					<a href="javascript: cmdChooseDate_onclick(milestoneDate1)"><img style="margin-bottom: -4px" id="Img1" src="../../mobilese/today/images/calendar.gif" alt="Choose Date" width="26" height="21" /></a>
				</span>
			</td>
		</tr>
		<tr style="display: <%=showField%>">
			<td style="white-space: nowrap; font-weight: bold">
				<input style="width: 100%" id="milestoneLabel2" maxlength="50" type="text" value="<%=strMilestoneDateLabel2%>" />
			</td>
			<td style="white-space: nowrap">
				<span id="Span2" style="display: <%=showField%>">&nbsp;
					<input style="width: auto" id="milestoneDate2" type="text" value="<%=strMilestoneDate2%>" onchange="txtDate_onchange(this);" />&nbsp;
					<a href="javascript: cmdChooseDate_onclick(milestoneDate2)"><img style="margin-bottom: -4px" id="Img2" src="../../mobilese/today/images/calendar.gif" alt="Choose Date" width="26" height="21" /></a>
				</span>
			</td>
		</tr>
		<%
		showField = "inline"
		if request("TypeID") = "3" then
			showField = "none"
		end if
		%>
		<tr style="display: <%=showField%>">
			<td style="white-space: nowrap">
				<b>Weeks:&nbsp;&nbsp;&nbsp;</b>
			</td>
			<td>
				<select id="cboWeeks" style="width: 80px" onchange="cboWeeks_onchange();">
					<%
			for i = 4 to 52
				if i=16 then
					if trim(strWeeks) = trim(i) then
						response.write "<option selected=""selected"" value="""">" & i & "</option>"
					else
						response.write "<option value="""">" & i & "</option>"
					end if
				else
					if trim(strWeeks) = trim(i) then
						response.write "<option selected=""selected"" value=""" & i & """>" & i & "</option>"
					else
						response.write "<option value=""" & i & """>" & i & "</option>"
					end if
				end if
			next
					%>
				</select>&nbsp;Start Date: <span id="spnStart">
					<%=strStartDate%></span><span style="display: <%=ShowStartWarning%>" id="spnStartWarning"> - No earlier data available.</span>
			</td>
		</tr>
		<tr>
			<td>
				<b>Gridline&nbsp;Interval:&nbsp;&nbsp;</b>
			</td>
			<td style="width: 100%">
				<select id="cboGridLine" style="width: 80px" onchange="cboGridLine_onchange();">
					<%if trim(strGridLine)="" then%>
					<option selected="selected" value="0">Auto</option>
					<option value="1">Custom</option>
					<%else%>
					<option value="0">Auto</option>
					<option selected="selected" value="1">Custom</option>
					<%end if%>
				</select>&nbsp;<span style="display: <%=ShowGridLine%>" id="spnGridline">Display a gridline every&nbsp;<input style="width: 40" id="txtGridlines" type="text" value="<%=strGridLine%>" />&nbsp;observations.</span>
			</td>
		</tr>
		<tr>
			<td>
				<b>Graph Width:&nbsp;&nbsp;</b>
			</td>
			<td style="width: 100%">
				<select id="cboWidth" style="width: 80px">
					<%
			WidthArray=split(",,Small,,Medium,,Large",",")
			for i = 0 to ubound(widtharray)
				if trim(WidthArray(i)) <> "" then
					if trim(strWidth) = trim(i) then
						response.write "<option selected=""selected"" value=""" & i & """>" & WidthArray(i) & "</option>"
					else
						response.write "<option value=""" & i & """>" & WidthArray(i) & "</option>"
					end if
				end if
			next
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<b>Graph Height:&nbsp;&nbsp;</b>
			</td>
			<td style="width: 100%">
				<select id="cboHeight" style="width: 80px">
					<%
			HeightArray=split(",,Small,,Medium,,Large",",")
			for i = 0 to ubound(HeightArray)
				if trim(HeightArray(i)) <> "" then
					if trim(strHeight) = trim(i) then
						response.write "<option selected=""selected"" value=""" & i & """>" & HeightArray(i) & "</option>"
					else
						response.write "<option value=""" & i & """>" & HeightArray(i) & "</option>"
					end if
				end if
			next
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td>
				<b>Legend&nbsp;&nbsp;</b>
			</td>
			<td style="width: 100%">
				<%LegendArray = split("Hide,Right,Bottom",",")%>
				<select id="cboLegend" style="width: 80px">
					<%
			for i = 0 to ubound(LegendArray)
				if trim(strlegend) = trim(i) then
					response.write "<option selected=""selected"" value=""" & i & """>" & LegendArray(i) & "</option>"
				elseif not (request("TypeID") = "2" and i = 0) then
					response.write "<option value=""" & i & """>" & LegendArray(i) & "</option>"
				end if
			next
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td style="font-weight: bold">Chart Type&nbsp;&nbsp;
			</td>
			<td style="width: 100%">
				<select id="cboChartType" style="width: 140px">
					<%
'			ChartTypeArray = split("BarClustered,BarClustered3D,BarStacked,BarStacked3D,ColumnClustered,ColumnClustered3D,ColumnStacked,ColumnStacked3D,LineMarkers,LineStacked",",")
			'NOTE: OWC control has a bug that displays wrong colors for 3D chart types, so we're not going to offer 3D chart types
			if trim(strChartType) = "BarClustered3D" then
				strChartType = "BarClustered"
			end if
			if trim(strChartType) = "BarStacked3D" then
				strChartType = "BarStacked"
			end if
			if trim(strChartType) = "ColumnClustered3D" then
				strChartType = "ColumnClustered"
			end if
			if trim(strChartType) = "ColumnStacked3D" then
				strChartType = "ColumnStacked"
			end if
			ChartTypeArray = split("BarClustered,BarStacked,ColumnClustered,ColumnStacked,LineMarkers,LineStacked",",")
			for i = 0 to ubound(ChartTypeArray)
				if trim(strChartType) = trim(ChartTypeArray(i)) then
					response.write "<option selected=""selected"" value=""" & ChartTypeArray(i) & """>" & ChartTypeArray(i) & "</option>"
				else'if not (request("TypeID") = "2" and i = 0) then
					response.write "<option value=""" & ChartTypeArray(i) & """>" & ChartTypeArray(i) & "</option>"
				end if
			next
					%>
				</select>
			</td>
		</tr>
		<%
			dim strDisplayType
			strDisplayType="none"
			if request("TypeID") = "1" and request("txtID") <> 21 then
				strDisplayType="inline"
			end if
		%>
		<tr style="display: <%=strDisplayType%>">
			<td>
				<b>Activity&nbsp;Graph:&nbsp;&nbsp;</b>
			</td>
			<td style="width: 100%">
				<select id="cboActivity" style="width: 80px">
					<%if strActivity="1" then%>
					<option value="0">Hide</option>
					<option selected="selected" value="1">Show</option>
					<%else%>
					<option selected="selected" value="0">Hide</option>
					<option value="1">Show</option>
					<%end if%>
				</select>
			</td>
		</tr>
		<tr style="display: <%=strDisplayType%>">
			<td>
				<b>Show&nbsp;Backlog&nbsp;Total?&nbsp;&nbsp;</b>
			</td>
			<td style="width: 100%">
				<select id="cboBacklog" style="width: 80px">
					<%if strBacklog="1" then%>
					<option value="0">Hide</option>
					<option selected="selected" value="1">Show</option>
					<%else%>
					<option selected="selected" value="0">Hide</option>
					<option value="1">Show</option>
					<%end if%>
				</select>
			</td>
		</tr>
	</table>
	<hr />
	<table style="width: 100%">
		<tr>
			<td align="right">
				<input id="cmdOk" type="button" value="OK" onclick="cmdOK_click();" />
			</td>
		</tr>
	</table>
	<%
		end if
	%>
	<input style="display: none; width: 100%" id="txtParams" type="text" value="<%=request("txtParams")%>" />
	<input style="display: none; width: 100%" id="txtType" type="text" value="<%=request("TypeID")%>" />
	<input style="display: none; width: 100%" id="txtID" type="text" value="<%=request("txtID")%>" />
	<%
	function getSectionTitle(sectionID)
		select case clng(sectionID)
			case 5
				getSectionTitle = "Observations by Priority"
			case 9
				getSectionTitle = "Observations by Sub System"
			case 11
				getSectionTitle = "Observations by Core Team"
			case 12
				getSectionTitle = "Observations by Component PM"
			case 10
				getSectionTitle = "Observations by State"
			case 13
				getSectionTitle = "Observations by Status"
			case 8
				getSectionTitle = "Observations by Deliverable"
			case 6
				getSectionTitle = "Observations by Developer"
			case 4,14
				getSectionTitle = "Weekly Backlog Graph"
			case 15,16,17,18,19,20
				getSectionTitle = "Weekly Backlog Group Graph"
			case 21
				getSectionTitle = "Weekly Observation Counts Graph"
			case 22,23,24
				getSectionTitle = "Current Risk Observations by Group Graph"
			case 25,26,27
				getSectionTitle = "Average Days Open by Group Graph"
			case 0
				getSectionTitle = "Summary Report"
			case else
				getSectionTitle = "Weekly Observation Backlog"
		end select
	end function
	%>
</body>
</html>
