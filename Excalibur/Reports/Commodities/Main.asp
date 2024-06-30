<%@ Language="VBScript" %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../../includes/noaccess.inc" -->


<html>
<head>
    <title>Commodity Yearly Reports</title>
    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
        function cmdDate_onclick(FieldID) {
        	var strID;
        	var oldValue = frmMain.elements(FieldID).value;
		
	        strID = window.showModalDialog("../../mobilese/today/caldraw1.asp",oldValue,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	        if (typeof(strID) == "undefined")
		        return
        	frmMain.elements(FieldID).value = strID;
        }
        
        function RunReport(ID){
            if (ID==1)
                {
                frmMain.action="WeeklyVersionsCreated.asp?Title=Versions Tagged to Products&Report=1";
                frmMain.submit(); 
                }
            if (ID==10)
                {
                frmMain.action="WeeklyVersionsCreated.asp?Title=Versions That Went QComplete&Report=2";
                frmMain.submit(); 
                }
            if (ID==11)
                {
                frmMain.action="WeeklyVersionsCreated.asp?Title=Deliverable Versions Created&Report=3";
                frmMain.submit(); 
                }
            else if (ID==2)
                {
                frmMain.action="PartNumbersTaggedQComplete.asp?Report=1&Title=Part Numbers Tagged QComplete";
                frmMain.submit(); 
                }
            else if (ID==3)
                {
                frmMain.action="PartNumbersTaggedQComplete.asp?Report=2&Title=Total OTS Before QComplete";
                frmMain.submit(); 
                }
            else if (ID==4)
                {
                frmMain.action="PartNumbersTaggedQComplete.asp?Report=3&Title=Total OTS After QComplete";
                frmMain.submit(); 
                }
            else if (ID==5)
                {
                frmMain.action="PartNumbersTaggedQComplete.asp?Report=4&Title=No OTS Before QComplete";
                frmMain.submit(); 
                }
            else if (ID==6)
                {
                frmMain.action="PartNumbersTaggedQComplete.asp?Report=5&Title=No OTS After QComplete";
                frmMain.submit(); 
                }
            else if (ID==7)
                {
                frmMain.action="PartNumbersTaggedQComplete.asp?Report=6&Title=No OTS Before or After QComplete";
                frmMain.submit(); 
                }
            else if (ID==8)
                {
                frmMain.action="PartNumbersTaggedQComplete.asp?Report=7&Title=OTS Before QComplete";
                frmMain.submit(); 
                }
            else if (ID==9)
                {
                frmMain.action="PartNumbersTaggedQComplete.asp?Report=8&Title=OTS Only After QComplete";
                frmMain.submit(); 
                }
            else if (ID==12)
                {
                frmMain.action="WorkflowCompleteTaggedQComplete.asp?Title=When Versions First Went QComplete";
                frmMain.submit(); 
                }
        }
        
    //-->
</SCRIPT>
</head>

<STYLE>
    td{
        FONT-FAMILY: Verdana;   
        FONT-SIZE: x-small;
    }
    A:link
    {
        COLOR: blue
    }
    A:visited
    {
        COLOR: blue
    }
    A:hover
    {
        COLOR: red
    }    
</STYLE>

<body bgcolor=ivory>
<%
    set cn = server.CreateObject("ADODB.Connection")
    cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
    cn.Open

    set rs = server.CreateObject("ADODB.recordset")
%>
<font size=3 face=verdana><b>Commodity Yearly Reporting Tools</b><BR><BR></font>
<font size=2 face=verdana><u><b>Report Options</b></u></font><BR><BR>
<form id=frmMain target=_blank method=post>
<table width=100% cellpadding=1 cellspacing=0 border=1>
    <TR>
        <TD nowrap valign=top><b>&nbsp;Date Range:&nbsp;</b></td>
        <TD>
            <table cellpadding=0 cellspacing=0>
                <tr>
                    <TD>
			            &nbsp;Start:&nbsp;
			        </td>
			        <td nowrap>
			            <input style="width:80px" type="text" id="txtStartDate" name="txtStartDate" value="<%=formatdatetime(now()-365,vbshortdate)%>">
			            <a href="javascript:cmdDate_onclick('txtStartDate');"><img SRC="../../MobileSE/Today/images/calendar.gif" alt="Choose" border="0" align="absmiddle" WIDTH="26" HEIGHT="21"></a>&nbsp;
			        </td>
	    	    </tr>
    			<tr>
    			    <td>   
            			&nbsp;End:&nbsp;
            		</td>
            	    <td>	
            			<input style="width:80px" type="text" id="txtEndDate" name="txtEndDate" value="<%=formatdatetime(now(),vbshortdate)%>">
			            <a href="javascript:cmdDate_onclick('txtEndDate');"><img SRC="../../MobileSE/Today/images/calendar.gif" alt="Choose" border="0" align="absmiddle" WIDTH="26" HEIGHT="21"></a></font>&nbsp;
                    </td>
                </tr>
            </table>
        </TD>
        <TD valign=top rowspan=3><b>&nbsp;Categories:&nbsp;</b></td>
        <TD width=100% rowspan=3 valign=top>
        <%
            dim ColumnCount
            ColumnCount=0
            rs.open "spListCommodityCategories",cn
            response.write "<table><TR>"
            do while not rs.eof
                if Columncount mod 3 = 0 and columncount <> 0 then
                    response.write "</tr><tr>"
                end if
                ColumnCount=ColumnCount + 1
                response.write "<td><input id=""lstCategories"" name=""lstCategories"" type=""checkbox"" checked value=""" & rs("ID") & """>" & rs("Name") & "&nbsp;&nbsp;&nbsp;</td>"
                rs.movenext
            loop
            rs.close
            response.write "</tr></Table>"
                
         %>
        </TD>
    </TR>

    <TR>
        <TD><b>&nbsp;Options:</b></td>
        <TD><input id="chkZeros" name="chkZeros" checked type="checkbox">&nbsp;Remove empty rows</TD>
        
    </TR>

    <TR>
        <TD><b>&nbsp;Format:</b></td>
        <TD>
            <SELECT style="width:170px" id=cboFormat name=cboFormat>			    <OPTION value=0 selected>HTML</OPTION>				<OPTION value=1>Excel</OPTION> 		    </SELECT>        
        </TD>
        
    </TR>

</table>
</form>
<font size=2 face=verdana><u><b>Available Reports</b></u></font><BR><BR>
<table width=100% cellpadding=1 cellspacing=0 border=1>
    <TR bgcolor=gainsboro>
        <TD nowrap><b>Report Name</b>&nbsp;&nbsp;&nbsp;</TD>
        <TD width=100%><b>Description</b></TD>
    </TR>
    <TR>
        <TD nowrap valign=top>&nbsp;<a href="javascript: RunReport(1);">Versions Tagged to Products</a>&nbsp;&nbsp;
        </TD>
        <TD>
            Displays how many times the deliverables in each category were added to All, Commercial, and Consumer products each week.  So, one deliverable version added to the <font color=red>supported</font> list on 10 products counts as 10 in this report.
        </TD>
    </TR>
    <TR>
        <TD nowrap valign=top>&nbsp;<a href="javascript: RunReport(10);">Versions That Went QComplete</a>&nbsp;&nbsp;
        </TD>
        <TD>
            Displays how many times the deliverables in each category went QComplete on All, Commercial, and Consumer products each week.  So, one deliverable version went <font color=red>QComplete</font> on 10 products counts as 10 in this report.
        </TD>
    </TR>
    <TR>
        <TD nowrap valign=top>&nbsp;<a href="javascript: RunReport(11);">Deliverable Versions Created</a>&nbsp;&nbsp;
        </TD>
        <TD>
            Displays how many deliverable versions in each category were <font color=red>created</font> each week.
        </TD>
    </TR>
    <TR>
        <TD nowrap valign=top>&nbsp;<a href="javascript: RunReport(12);">When Versions First Went QComplete</a>&nbsp;&nbsp;
        </TD>
        <TD>Displays all deliverable versions that have gone QComplete on at least one product.
        </TD>
    </TR>
    <TR>
        <TD nowrap valign=top>&nbsp;<!--<a href="javascript: RunReport(2);">Part Numbers Tagged QComplete</a>-->Part Numbers Tagged QComplete&nbsp;&nbsp;[Disabled]
        </TD>
        <TD>
            Commdity Observation Report: Tab 1 - Displays all part numbers that went Qcomplete (last Qcomplete) during the selected date range. Includes how many observations were writted before Qcomplete, after QComplete, and the total number of QCompletes (HW Matrix Cells). 
        </TD>
    </TR>
    <TR>
        <TD nowrap valign=top>&nbsp;<!--<a href="javascript: RunReport(3);">Total OTS Before QComplete</a>-->Total OTS Before QComplete&nbsp;&nbsp;[Disabled]
        </TD>
        <TD>
            Commdity Observation Report: Tab 2 - Same as tab 1 except it only includes part numbers that had observations and it only includes a column to display how many occurred before QComplete.
        </TD>
    </TR>
    <TR>
        <TD nowrap valign=top>&nbsp;<!--<a href="javascript: RunReport(4);">Total OTS After QComplete</a>-->Total OTS After QComplete&nbsp;&nbsp;[Disabled]
        </TD>
        <TD>
            Commdity Observation Report: Tab 3 - Same as tab 1 except it only includes part numbers that had observations and it only includes a column to display how many occurred after QComplete. 
        </TD>
    </TR>
    <TR>
        <TD nowrap valign=top>&nbsp;<!--<a href="javascript: RunReport(5);">No OTS Before QComplete</a>-->No OTS Before QComplete&nbsp;&nbsp;[Disabled]
        </TD>
        <TD>
            Commdity Observation Report: Tab 4  - Same as tab 1 except it only includes part numbers that had no observations before QComplete
        </TD>
    </TR>
    <TR>
        <TD nowrap valign=top>&nbsp;<!--<a href="javascript: RunReport(6);">No OTS After QComplete</a>-->No OTS After QComplete&nbsp;&nbsp;[Disabled]
        </TD>
        <TD>
            Commdity Observation Report: Tab 5  - Same as tab 1 except it only includes part numbers that had observations after QComplete
        </TD>
    </TR>
    <TR>
        <TD nowrap valign=top>&nbsp;<!--<a href="javascript: RunReport(7);">No OTS Before or After QComplete</a>-->No OTS Before or After QComplete&nbsp;&nbsp;[Disabled]
        </TD>
        <TD>
            Commdity Observation Report: Tab 6  - Same as tab 1 except it only includes part numbers that had no observations
        </TD>
    </TR>
    <TR>
        <TD nowrap valign=top>&nbsp;<!--<a href="javascript: RunReport(8);">OTS Before QComplete</a>-->OTS Before QComplete&nbsp;&nbsp;[Disabled]
        </TD>
        <TD>
            Commdity Observation Report: Tab 7  - Same as tab 1 except it only includes part numbers that had observations before QComplete
        </TD>
    </TR>
    <TR>
        <TD nowrap valign=top>&nbsp;<!--<a href="javascript: RunReport(9);">OTS Only After QComplete</a>-->OTS Only After QComplete&nbsp;&nbsp;[Disabled]
        </TD>
        <TD>
            Commdity Observation Report: Tab 8  - Same as tab 1 except it only includes part numbers that had observations after QComplete
        </TD>
    </TR>
    <TR>
        <TD nowrap valign=top>&nbsp;List Observations Before Complete&nbsp;&nbsp;
        </TD>
        <TD>
            Commdity Observation Report: Tab 9  - Currently this query can not be updated to run in a web page.  Contact Dave Whorton to get the data for this page.
        </TD>
    </TR>
    <TR>
        <TD nowrap valign=top>&nbsp;List Observations After QComplete&nbsp;&nbsp;
        </TD>
        <TD>
            Commdity Observation Report: Tab 10  - Currently this query can not be updated to run in a web page.  Contact Dave Whorton to get the data for this page.
        </TD>
    </TR>
</table>

<%
    set rs = nothing
    cn.close
    set cn = nothing
%>

</body>
</html>

