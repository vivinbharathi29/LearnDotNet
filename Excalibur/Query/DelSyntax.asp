<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<FONT face=Verdana size=2>
<B>Deliverable Query Syntax</B><BR><BR>
Here is a brief overview of the syntax required in the Other Criteria field.<BR><BR>

To enter a basic filter, use this format:<BR><BR>
<P style="TEXT-INDENT: 0.5in">&lt;FieldName&gt; &lt;Operator&gt; &lt;Value&gt;</P>
<P style="MARGIN-LEFT: 1in">Examples:</P>
<TABLE style="BORDER-RIGHT: medium none; BORDER-TOP: medium none; MARGIN-LEFT: 66.2pt; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none; BORDER-COLLAPSE: collapse" cellSpacing=0 cellPadding=0 border=1>
<TR>
	<TD style="BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: windowtext 1pt solid; PADDING-LEFT: 5.4pt; BACKGROUND: #99ccff; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 1pt solid; WIDTH: 185.4pt; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 1pt solid" vAlign=top width=247 bgColor=#99ccff><font size=2 face=verdana><B>If You Want Observations Where</B></font></TD>
	<TD style="BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: windowtext 1pt solid; PADDING-LEFT: 5.4pt; BACKGROUND: #99ccff; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 1pt solid; WIDTH: 185.4pt; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 1pt solid" vAlign=top width=247 bgColor=#99ccff><font size=2 face=verdana><B>Enter This Filter</B></font></TD>
</TR>
<TR>
	<TD style="BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 1pt solid; WIDTH: 185.4pt; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 1pt solid" vAlign=top width=247><FONT face=verdana size=2>Jim Johnson is listed as developer</FONT></TD>
	<TD style="BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 1pt solid; WIDTH: 185.4pt; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 1pt solid" vAlign=top width=247><FONT face=verdana size=2>Developer = 'jjohnson'</FONT></TD>
</TR>
<TR>
	<TD style="BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 1pt solid; WIDTH: 185.4pt; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 1pt solid" vAlign=top width=247><FONT face=verdana size=2>Only Priority 1</FONT></TD>
	<TD style="BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 1pt solid; WIDTH: 185.4pt; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 1pt solid" vAlign=top width=247><FONT face=verdana size=2>Priority = 1</FONT></TD>
</TR>
<TR>
	<TD style="BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 1pt solid; WIDTH: 185.4pt; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 1pt solid" vAlign=top width=247><FONT face=verdana size=2>Status is Closed (2) or Sustain (3)</FONT></TD>
	<TD style="BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 1pt solid; WIDTH: 185.4pt; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 1pt solid" vAlign=top width=247><FONT face=verdana size=2>Status &gt; 1</FONT></TD>
</TR>
<TR>
	<TD style="BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 1pt solid; WIDTH: 185.4pt; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 1pt solid" vAlign=top width=247><FONT face=verdana size=2>State is not “Retest”</FONT></TD>
	<TD style="BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 1pt solid; WIDTH: 185.4pt; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 1pt solid" vAlign=top width=247><FONT face=verdana size=2>State &lt;&gt; 'Retest'</FONT></TD>
</TR>
<TR>
	<TD style="BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 1pt solid; WIDTH: 185.4pt; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 1pt solid" vAlign=top width=247><FONT face=verdana size=2>All Components with “video” in name</FONT></TD>
	<TD style="BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 1pt solid; WIDTH: 185.4pt; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 1pt solid" vAlign=top width=247><FONT face=verdana size=2>Components like '%video%'</FONT></TD>
</TR>
<TR>
	<TD style="BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 1pt solid; WIDTH: 185.4pt; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 1pt solid" vAlign=top width=247><FONT face=verdana size=2>ReleaseFixImplement field is not empty</FONT></TD>
	<TD style="BORDER-RIGHT: windowtext 1pt solid; PADDING-RIGHT: 5.4pt; BORDER-TOP: medium none; PADDING-LEFT: 5.4pt; PADDING-BOTTOM: 0in; BORDER-LEFT: windowtext 1pt solid; WIDTH: 185.4pt; PADDING-TOP: 0in; BORDER-BOTTOM: windowtext 1pt solid" vAlign=top width=247><FONT face=verdana size=2>ReleaseFixImplemented is not null</FONT></TD>
</TR>
</TABLE>
<b><BR><BR>Helpful Hints:</b><BR><BR>
<OL>
	<LI>Use the Field chooser to find the names of the 
  fields.&nbsp;They are not always named the same as in the interface and 
  sometimes there is a prefix that it needs to have. 
  <LI>Use Parentheses around the basic filter elements when using OR<BR><BR>Example: (PM = 'sogle' or PM = 'rpyra')<BR><BR></LI>
  <LI>Use AND to add more filters<BR><BR>Example: (PM = 'sogle' or PM = 'rpyra') and (Priority = 1 or Priority = 0)<BR><BR></LI>
  <LI>Use IN to match a field to any item in a list<BR><BR>Example (Equilivant to item 3 above): (PM in ('sogle','rpyra') ) and (Priority in (1, 0))<BR><BR></LI>
  <LI>Capitalization in field names or values hardly every matters</LI>
  <LI>OTS Statuses are: 0=new, 1=Open, 2=Closed, 3=Sustain</LI>
  <LI>User single quotes around strings.&nbsp;Example: 'sogle'</LI>
  <LI>Use LIKE and wildcard characters to search text:&nbsp;Component like 'Video%' finds all component names that start with “video”</LI>
</OL>
</BODY>
</HTML>
