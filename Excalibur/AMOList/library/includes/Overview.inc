<script type="text/javascript" language=JavaScript runat=Server>
// When page is run, modifies navigation menu's Help menu to display Help for this page
function SetHelp(sLink) {
	var sHTML = "";
	var sLess = "<"
	if (sLink != '') {
		sHTML += sLess + "script LANGUAGE=\"JavaScript\">\n"
		sHTML += "document.all.item(\"pagehelp\").href = \"" + sLink + "\"\n"
		sHTML += "document.all.item(\"pagehelpdisplay1\").style.display = \"\"\n"
		sHTML += "document.all.item(\"pagehelpdisplay2\").style.display = \"\"\n"
		sHTML += "document.all.item(\"pagehelpdisplay3\").style.display = \"\"\n"
		sHTML += sLess + "/script>\n"
	}
	return sHTML
}

// Old way to display Overview link
function BuildOverview(sLink) {
	var sHTML = "";
	sHTML += "<table border='0' cellPadding='0' cellSpacing='0' width='100%'>\n";
	sHTML += "\t<tr>\n";
	sHTML += "\t\t<td align='left' width='100%'><font face size='2'><a href='";
	sHTML += sLink;
	sHTML += "' target='_blank'>Overview</a></font></td>\n";
	sHTML += "\t</tr>\n";
	sHTML += "\t<tr>\n";
	sHTML += "\t\t<td colspan='2'><hr align='left' size='2' width='100%'></td>\n";
	sHTML += "\t</tr>\n";
	sHTML += "</table>\n";
	return sHTML;
}

// This displays the help link at the far right of the screen
function BuildHelpLink(sTitle, sLink) {
	var sHTML = "";
	sHTML += "<table border='0' cellPadding='0' cellSpacing='0' width='100%'>\n";
	sHTML += "\t<tr>\n";
	sHTML += "\t\t<td align='left'><h1 valign='top' align='left'>" + sTitle + "</h1></td>\n";
	sHTML += "\t\t<td id=tdHelp name=tdHelp align='right'>";
	if (sLink == "") {
		sHTML += "</td>\n";
	} else {
		sHTML += "<font face size='2'><a href='"
		sHTML += sLink;
		sHTML += "' target='_blank'>Help</a></font></td>\n";
	}
	sHTML += "\t</tr>\n";
	sHTML += "\t<tr>\n";
	sHTML += "\t\t<td colspan='2'><hr align='left' size='2' width='100%'></td>\n";
	sHTML += "\t</tr>\n";
	sHTML += "</table>\n";
	return sHTML;
}

// This puts the help link in the page's navigation Help menu
function BuildHelp(sTitle, sLink) {
	var sHTML = "";

	if (sLink != "")
		sHTML += SetHelp(sLink)

	sHTML += "<table border='0' cellPadding='0' cellSpacing='0' width='100%'>\n";
	sHTML += "\t<tr>\n";
	if (sTitle == "") {
		sHTML += "\t\t<td align='left'>&nbsp;</td>\n";
	} else {
		sHTML += "\t\t<td align='left'><h1 valign='top' align='left'>" + sTitle + "</h1></td>\n";
	}
	// following help cell kept for compatibility with pages that want to hide it
	sHTML += "\t\t<td id=tdHelp name=tdHelp align='right'></td>\n";

	sHTML += "\t</tr>\n";
	sHTML += "\t<tr>\n";
	sHTML += "\t\t<td colspan='2'><hr align='left' size='2' width='100%'></td>\n";
	sHTML += "\t</tr>\n";
	sHTML += "</table>\n";
	return sHTML;
}

// Built from BuildHelp
function BuildHelpNoLine(sTitle, sLink) {
	var sHTML = "";

	if (sLink != "")
		sHTML += SetHelp(sLink)

	sHTML += "<table border='0' cellPadding='0' cellSpacing='0' width='100%'>\n";
	sHTML += "\t<tr>\n";
	if (sTitle == "") {
		sHTML += "\t\t<td align='left'>&nbsp;</td>\n";
	} else {
		sHTML += "\t\t<td align='left'><h1 valign='top' align='left'>" + sTitle + "</h1></td>\n";
	}
	// following help cell kept for compatibility with pages that want to hide it
	sHTML += "\t\t<td id=tdHelp name=tdHelp align='right'></td>\n";

	sHTML += "\t</tr>\n";
	sHTML += "\t<tr><td colspan=2></td>\n";
	sHTML += "\t</tr>\n";
	sHTML += "</table>\n";
	return sHTML;
}

// This puts the help link in the page's navigation Help menu and centers info
function BuildHelpCentered(sTitle, sLink, sWidth) {
	var sHTML = "";
	if (sLink != "")
		sHTML += SetHelp(sLink)

	sHTML += "<div align='center'>\n";
	sHTML += "<table border='0' cellPadding='0' cellSpacing='0' align='center' width='" + sWidth + "'>\n";
	sHTML += "\t<tr>\n";
	sHTML += "\t\t<td align='left'><h1 valign='top' align='left'>" + sTitle + "</h1></td>\n";
	sHTML += "\t</tr>\n";
	sHTML += "\t<tr>\n";
	sHTML += "\t\t<td><hr align='left' size='2' width='100%'></td>\n";
	sHTML += "\t</tr>\n";
	sHTML += "</table>\n";
	sHTML += "</div>\n";
	return sHTML;
}

// Old function, probably not used
function BuildHelpCenteredOld(sTitle, sLink, sWidth) {
	var sHTML = "";
	sHTML += "<table border='0' cellPadding='0' cellSpacing='0' align='center' width='" + sWidth + "'>\n";
	sHTML += "\t<tr>\n";
	sHTML += "\t\t<td align='left'><h1 valign='top' align='left'>" + sTitle + "</h1></td>\n";
	sHTML += "\t\t<td align='right'><font face size='2'><a href='";
	sHTML += sLink;
	sHTML += "' target='_blank'>Help</a></font></td>\n";
	sHTML += "\t</tr>\n";
	sHTML += "\t<tr>\n";
	sHTML += "\t\t<td colspan='2'><hr align='left' size='2' width='100%'></td>\n";
	sHTML += "\t</tr>\n";
	sHTML += "</table>\n";
	return sHTML;
}

</script>
