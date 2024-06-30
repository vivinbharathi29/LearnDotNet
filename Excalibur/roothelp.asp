<%@ Language=VBScript %>
<%Response.Expires = 0%>

<html>
<head>
<title>Context Help</title>


</HEAD>


<BODY bgColor=ivory>




<%if Request("Context") = "" then
%>
	<font size=4 face=verdana>Root Deliverable Help</font>

	<BLOCKQUOTE><FONT face=Verdana size=2>
	    <FONT size=2 color=red>Note: In 
  addition to the Help button, you can also select any field&nbsp;and press F1 
  for Context Help. </FONT></FONT></BLOCKQUOTE>
<BLOCKQUOTE>
  <P><FONT face=Verdana size=2><FONT size=2><STRONG>Help 
  Topics</STRONG></FONT></FONT></P>
  <BLOCKQUOTE dir=ltr style="MARGIN-RIGHT: 0px">
    <P>
	<A href="#QETester"><font face=verdana size=2>QE Tester</font></A><BR>
	<A href="#BasePart"><font face=verdana size=2>Base 6 Part Number</font></A><BR>
	<A href="#Filename"><font face=verdana size=2>Root Filename</font></A>
  </P></BLOCKQUOTE></BLOCKQuote>
<%
  end if
  
  if Request("Context") = "1" or Request("Context") = "" then
%>

	<BLOCKQUOTE>
  <P><FONT face=Verdana size=2><strong><BR><a name=#QETester>QE&nbsp;Tester</a></STRONG> </FONT></P>
		<P dir=ltr><FONT face=Verdana size=2> The QE&nbsp;Tester is the person that is 
	      responsible for ensuring that a deliverable gets tested. </FONT></P>
		<UL>
	        <LI dir=ltr><FONT face=Verdana size=2>If no such person has been 
			assigned to your deliverable, leave this field blank.</FONT> 
			<LI dir=ltr><FONT face=Verdana size=2>  If the QA&nbsp;Tester for 
			your deliverable is not in the QA&nbsp;Tester List, click the add button to enter 
			him or her into the system.&nbsp; Please be very careful when entering 
			the employee&nbsp;information to ensure that notifications and other 
			automated processes will function properly.&nbsp; If you do not know the 
			information requested, you can find it in the Outlook address 
			book.</FONT></LI></UL></BLOCKQUOTE></STRONG></FONT>
<%
	end if
  if Request("Context") = "3" or Request("Context") = "" then
%>
	<BLOCKQUOTE>
	<P><FONT face=Verdana size=2>
	    <FONT size=2><STRONG><BR><a name=#Filename>Root Filename</a></STRONG>&nbsp; </FONT>
		<P dir=ltr><FONT face=Verdana size=2>  The first&nbsp;part of the deliverable filename</FONT></P>
  <UL>
	        <LI dir=ltr> May be up to 12 characters 
    long&nbsp;
    
			<LI dir=ltr><FONT face=Verdana size=2>      
			              
			            
			        
			           
			     The Version, Revision and Pass will be      
			              
			            
			        
			           
			     added as each version is 
			entered</FONT></LI></UL>
  <P>Example:</P>
  <P>If the&nbsp;filename for version 1.00,A,1&nbsp;of this 
  deliverable&nbsp;would be MARGI_A1.100, you would enter <FONT color=red>MARGI 
  </FONT>in this field.</P></BLOCKQUOTE></STRONG></FONT>
  
  
<%
	end if
  if Request("Context") = "4" or Request("Context") = "" then
%>
	<BLOCKQUOTE>
	<P><FONT face=Verdana size=2>
	    <FONT size=2><STRONG><BR>  Deliverable Name</STRONG>&nbsp; </FONT>
		<P dir=ltr><FONT face=Verdana size=2>        
	        The official Name of the Deliverable. </FONT></P>
  <UL>
	        <LI dir=ltr>   Do not enter the version, revision, and pass 
    of the deliverable in this field.&nbsp; That information&nbsp;will be entered 
    when&nbsp;each new version is added.</LI></UL>
	</BLOCKQUOTE></FONT>
<%
	end if
  if Request("Context") = "5" or Request("Context") = "" then
%>	
	<BLOCKQUOTE>
	<P><FONT face=Verdana size=2>
	    <FONT size=2><STRONG><BR>  Software Type</STRONG>&nbsp; </FONT>
		<P dir=ltr><FONT face=Verdana size=2>        
	          One&nbsp;categorization of this deliverable. </FONT></P>
		<UL>
	        <LI dir=ltr>   Used for reports and displays.
			<LI dir=ltr><FONT face=Verdana size=2>This list is provided to support IRS.</FONT>
			<LI dir=ltr>The values in this list can not be changed without updating IRS.</LI></UL>
		</BLOCKQUOTE></STRONG></FONT>


<%
	end if
  if Request("Context") = "6" or Request("Context") = "" then
%>
	<BLOCKQUOTE>
	<P><FONT face=Verdana size=2>
	    <FONT size=2><STRONG><BR>   Category</STRONG>&nbsp; </FONT>
		<P dir=ltr><FONT face=Verdana size=2>        
	         One categorization of this deliverable. </FONT></P>
		<UL>
	        <LI dir=ltr>   Used for reports and 
    displays.&nbsp;
    
			<LI dir=ltr><FONT face=Verdana size=2>      
			              
			            
			        
			           
			     This      
			              
			            
			        
			           
			     list is provided to support 
			IRS.</FONT>
    <LI dir=ltr>The values in this list can not be changed without updating 
    IRS.</LI></UL></BLOCKQUOTE></STRONG></FONT>


<%
	end if
  if Request("Context") = "7" or Request("Context") = "" then
%>
	<BLOCKQUOTE>
	<P><FONT face=Verdana size=2>
	    <FONT size=2><STRONG><BR>PM (Development Manager)</STRONG>&nbsp; </FONT>
		<P dir=ltr><FONT face=Verdana size=2>       The person who is 
	      responsible for the development of this deliverable.&nbsp; &nbsp; </FONT></P>
  <UL>
	        <LI dir=ltr>Usually the Developer's Manager.
    <LI dir=ltr>If the&nbsp;PM for your deliverable is not in the&nbsp;PM List, 
    click the add button to enter him or her into the system.&nbsp; Please be 
    very careful when entering the employee&nbsp;information to ensure that 
    notifications and other automated processes will function properly.&nbsp; If 
    you do not know the information requested, you can find it in the Outlook 
    address book.&nbsp;</LI></UL></BLOCKQUOTE></STRONG></FONT>

<%
	end if 
  if Request("Context") = "8" or Request("Context") = "" then
%>
	<BLOCKQUOTE>
	<P><FONT face=Verdana size=2>
	    <FONT size=2><STRONG><BR>  Add Employee</STRONG>&nbsp; </FONT>
		<P dir=ltr><FONT face=Verdana size=2>       Registers a new employee to the database. <BR><BR><FONT 
  color=red>NOTE: Please enter employee information very carefully.&nbsp; All of 
  the information requested can be found in the Outlook address 
  book.</FONT>        
	              </FONT></P>
		<UL>
	        <LI dir=ltr>The new employee will be added to all employee dropdowns&nbsp;in the 
    system&nbsp;
    
			<LI dir=ltr><FONT face=Verdana size=2>      
			              
			            
			        
			           
			     The new employee will automatically be selected in 
    the&nbsp;dropdown      
			              
			            
			        
			           
			     to the left of this 
			button</FONT>
    <LI dir=ltr>A new account will be created in the system for the new 
    employee.&nbsp; This person&nbsp;will be access the system as soon as you 
    save the employee information assuming you enter all of his or her 
    information correctly.</LI></UL></BLOCKQUOTE></STRONG></FONT>

<%
	end if
%>
</BODY></html>
