<%@ Language=VBScript %>
<% OPTION EXPLICIT %>

<HTML>
 <HEAD>
  <TITLE>View/Modify After Market Option</TITLE>
 <LINK REL="StyleSheet" HREF="../library/stylesheets/IRSHELP.css"></HEAD>
  <BODY BGCOLOR="#FFFFFF">
<p align="right"> <a href="HELP_AMO_Overview.asp">After Market Options Overview</a>
  <H1>View/Modify After Market Option</H1>

	<p>Most of the fields for an After Market Option can be edited on one page. The properties page can be reached
	either from the main menu <b>Modules</b> item and selecting an option or selecting the <b>View/Modify Option Properties</b> menu
	item of an option being displayed in the After Market Option List.</p>
	
	<p>Any option that has a status of <b>RAS Review</b> is locked from further editing. The RAS Admin will have to either
	set the option to Complete or Reject in order to edit it again.</p>	
	
	<p>Following is a description of some of the fields:</p>

	<ul>
	<li><p>If the option should not be available to add to a Module and Option List, check the checkbox that says 
	<b>Hide from Module and Option List</b>.</p></li>

	<li><p>If the <b>RAS Available Date (Release to BOM Rev A.)</b> is changed, a calculation will be done and
	make the <b>CPL Blind Date</b> one month prior. If the new <b>CPL Blind Date</b> is not correct, it can be changed
	separately.</p></li>
	
	<li><p>If the <b>AMO Cost</b> value is changed, the <b>AMO Price</b> is set equal to <b>AMO Cost</b> times 2 unless 
	greater than 20 characters. The <b>AMO Price</b> field can be changed separately if needed.</p></li>

	<li><p>The properties page is the only place where additional Platforms can be added to the After Market Option List.</p></li>
	</ul>	


 </BODY>
</HTML>
