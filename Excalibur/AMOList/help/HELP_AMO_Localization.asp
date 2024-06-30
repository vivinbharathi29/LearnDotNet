<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<HTML>
<HEAD>
 <TITLE>After Market Option List - Localization</TITLE>
 <LINK rel="StyleSheet" HREF="../library/stylesheets/IRSHELP.css">
</HEAD>
<BODY>
<p align="right"><a href="HELP_AMO_Overview.asp">After Market Options Overview</a>
  <H1>After Market Option List - Localization</H1>
  <P>The Localization tab of the After Market Option List shows the Localization information for the Options
	 in a table sorted by Option Category. Most
	of the fields can be modified by directly clicking the table cell. Depending on the cell, a popup menu may appear
	or a text box where you can directly enter the data. When finished editing, either click elsewhere on the page or press
	Enter.</p>

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
	<li><p>To edit all the option properties at once, click the <b>View</b> link in the Properties column.</p>

	<li><p>Any option that has a status of RAS Review is locked from further editing. The RAS Admin will have to either
	set the option to Complete or Reject in order to edit it again.</p>

	<li><p>When sending an option to the RAS Admin, you may want to provide a comment to them for any particular
	exceptions in the Localizations. If so, click the <b>Add</b> link in the <b>Regional Comment</b> column to enter the comment.</p>

	<li><p>Enter a date for the GEOs listed.</p></li>

	<li><p>To specify that a Region is assigned to a particular option, simply click the corresponding
	table cell and select the <b>Add to Region</b> menu item. To uncheck the table cell, click the cell and
	select the <b>Remove from Region</b> menu item.</p>

	<li><p>If the option is applicable for ALL regions, simply check the <b>All Countries</b> table cell.</p></li>

	<li><p>If a value in a table cell changes, the background of the table cell will be turned yellow. This is to allow
	the RAS Admin to easily identify the changes made. Once the RAS Admin sets an option to Complete, the yellow highlighting
	will be removed.</p></li>
	</ul>
</BODY>
</HTML>
