<%@ Language=VBScript %>
<% OPTION EXPLICIT %>

<HTML>
<HEAD>
 <TITLE>After Market Option List - Options</TITLE>
 <LINK rel="StyleSheet" HREF="../library/stylesheets/IRSHELP.css">
</HEAD>
<BODY>
<p align="right"><a href="HELP_AMO_Overview.asp">After Market Options Overview</a>
  <H1>After Market Option List - Options</H1>
  <P>The Options tab of the After Market Option List shows the Options in a table sorted by Option Category. Most
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

	<li><p>When an option is ready to be sent to RAS Admin, click the status and select the <b>RAS Review</b> menu item.
	Any option that has a status of <b>RAS Review</b> is locked from further editing. The RAS Admin will have to either
	set the option to Complete or Reject in order to edit it again.</p>

	<li><p>When sending an option to the RAS Admin, you may want to provide a comment to them for any particular
	exceptions. If so, click the <b>Add</b> link in the <b>Comment to RAS</b> column to enter the comment.</p>

	<li><p>To specify that an option should not be available to add to a Module and Option List, simply click the corresponding
	table cell and select the <b>Hide from Module and Option List</b> menu item. To uncheck the table cell, click the cell and
	select the <b>Show in Module and Option List</b> menu item.</p>

	<li><p>If the <b>RAS Available Date (Release to BOM Rev A.)</b> is changed, a calculation will be done and
	make the <b>CPL Blind Date</b> one month prior. If the new <b>CPL Blind Date</b> is not correct, it can be changed
	separately.</p></li>

	<li><p>If the <b>AMO Cost</b> value is changed, the <b>AMO Price</b> is set equal to <b>AMO Cost</b> times 2 unless
	greater than 20 characters. The <b>AMO Price</b> field can be changed separately if needed.</p></li>

	<li><p>If a value in a table cell changes, the background of the table cell will be turned yellow. This is to allow
	the RAS Admin to easily identify the changes made. Once the RAS Admin sets an option to Complete, the yellow highlighting
	will be removed.</p></li>

	</ul>


</BODY>
</HTML>
