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
	<h2>RAS View</h2>

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
	<li><p>When an option has been sent to <b>RAS Review</b>, it needs to be entered in RAS and GPSy. After
	the information is entered into either system, check the appropriate checkbox to indicate that it has been done. When
	both <b>RAS</b> and <b>GPSy</b> checkboxes are checked, a popup message will appear asking if you are sure you
	want to put the option into Complete status. If you Cancel, the last checkbox you checked will be unchecked. If you
	answer OK, the option will be unlocked and put into Complete status.</p></li>

	<li><p>If there is a problem with the option, it can be rejected and sent back to the AMO Admin. Click the <b>RAS Review</b>
	status and select the <b>Reject</b> menu item.</p></li>

	<li><p>When an option is ready to be sent to back to the AMO Admin, you may want to provide a comment to them. This
	is particularly true if you Reject the option. To enter a comment, click the <b>Add</b> link in the <b>Comment from RAS</b>
	column to enter the comment.</p>

	<li><p>If a value in a table cell changes, the background of the table cell will be turned yellow. This is to allow
	the RAS Admin to easily identify the changes made. Once the RAS Admin sets an option to Complete, the yellow highlighting
	will be removed.</p></li>
	</ul>



</BODY>
</HTML>
