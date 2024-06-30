<%@ Language=VBScript %>
<% OPTION EXPLICIT %>

<HTML>
<HEAD>
 <TITLE>After Market Option List - Platforms</TITLE>
 <LINK rel="StyleSheet" HREF="../library/stylesheets/IRSHELP.css">
</HEAD>
<BODY>
<p align="right"><a href="HELP_AMO_Overview.asp">After Market Options Overview</a>
  <H1>After Market Option List - Platforms</H1>
  <P>The Platforms tab of the After Market Option List shows the Platforms for the Options in a table sorted by
	Option Category.</p>

	<p>The list is filtered depending on the category filter at the top of the page. The various filters
	available are:
	<ul>
		<li>Option Category: Hardware and Software categories</li>
		<li>AMO Status</li>
		<li>Business Segment</li>
		<li>Show options with RAS Discontinue Date on or after</li>
	</ul>
	</p>

	<p>Following is a description of some of the columns:</p>

	<ul>
	<li><p>To edit all the option properties, click the <b>View</b> link in the Properties column.
	This will need to be done if you want to add a Platform	that is not currently listed in the table.</p>

	<li><p>To specify that a Platform is assigned to a particular option, simply click the corresponding
	table cell and select the <b>Add to Platform</b> menu item. To uncheck the table cell, click the cell and
	select the <b>Remove from Platform</b> menu item.</p>

	<li><p>Any option that has a status of <b>RAS Review</b> is locked from further editing. The RAS Admin will have to either
	set the option to Complete or Reject in order to edit it again.</p>

	<li><p>If a value in a table cell changes, the background of the table cell will be turned yellow. This is to allow
	the RAS Admin to easily identify the changes made. Once the RAS Admin sets an option to Complete, the yellow highlighting
	will be removed.</p></li>
	</ul>



</BODY>
</HTML>
