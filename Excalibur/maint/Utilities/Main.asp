<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<STYLE>
body{
	Font-Size:xx-small;
	font-family:Verdana;
}
td{
	Font-Size:xx-small;
	font-family:Verdana;
}
th{
	Font-Size:xx-small;
	font-family:Verdana;
    font-weight:bold;
    text-align: left;
}
</STYLE>
<BODY>
    <strong>Sudden Impact Utilities
    </strong>
    <table width="100%" bgcolor=ivory cellpadding=2 cellspacing=0 border=1 bordercolor=gainsboro>
        <tr bgcolor=beige>
            <th>Name</th>
            <th>Action</th>
            <th>Scheduled</th>
            <th>Function</th>
        </tr>
        <tr>
            <td valign=top>Update Functional Test Products</td>
            <td valign=top nowrap><a target=_blank href="SuddenImpact/SIFunctionalTestDataPush.asp">Run</a></td>
            <td valign=top nowrap>Run As Needed</td>
            <td valign=top>Sends the Name, PM, and Tester for each Functional Test Product to Sudden Impact.  Make any required updates to the OTSVirtualProducts table in the PRS DB first and then run this job to push the changes to Sudden Impact.</td>
        </tr>
        <tr>
            <td valign=top>Sync Component Owners</td>
            <td valign=top nowrap><a target=_blank href="SuddenImpact/DeliverableSync.asp">Preview</a> | <a target=_blank href="SuddenImpact/DeliverableSync.asp?UpdateOK=1">Run</a></td>
            <td valign=top nowrap>Week Days - 10:05pm</td>
            <td valign=top>Syncs DevManagers/OTS PMs, Developers, and Test Leads</td>
        </tr>
        <tr>
            <td valign=top>Fix missing ODM Functional Test Products</td>
            <td valign=top nowrap><a target=_blank href="SuddenImpact/UpdateSIFunctionalTestODMProducts.asp">Preview</a> | <a target=_blank href="SuddenImpact/UpdateSIFunctionalTestODMProducts.asp?UpdateOK=1">Run</a></td>
            <td valign=top nowrap>Hourly - 8 After</td>
            <td valign=top>Looks for any components that are attached to only one Functional Test product and links those to the Required ODM Products.</td>
        </tr>
    </table>

    <br><br>
    <strong>Excalibur Deliverable Updates
    </strong>
    <table width="100%" bgcolor=ivory cellpadding=2 cellspacing=0 border=1 bordercolor=gainsboro>
        <tr bgcolor=beige>
            <th>Name</th>
            <th>Action</th>
            <th>Scheduled</th>
            <th>Function</th>
        </tr>
        <tr>
            <td valign=top>Remove Rejected Products from Roots</td>
            <td valign=top nowrap><a target=_blank href="Deliverables/RemoveRootsfromRejectedProducts.asp">Preview</a> | <a target=_blank href="Deliverables/RemoveRootsfromRejectedProducts.asp?UpdateOK=1">Run</a></td>
            <td valign=top>Run As Needed</td>
            <td valign=top>Removes root deliverables from products that have been rejected by the development team.</td>
        </tr>
        <tr>
            <td valign=top>Update Deliverable Paths</td>
            <td valign=top nowrap><a target=_blank href="../UpdateStoredPaths.asp">Run</a></td>
            <td valign=top>Run As Needed</td>
            <td valign=top>Allows you to batch update deliverable paths.  This tool is usually used to correct the deliverable paths when server names change.</td>
        </tr>
    </table>


    <br><br>
    <strong>Procedures
    </strong>
    <table width="100%" bgcolor=ivory cellpadding=2 cellspacing=0 border=1 bordercolor=gainsboro>
        <tr bgcolor=beige>
            <th>Event</th>
            <th>Action</th>
        </tr>
        <tr>
            <td valign=top>Switch ODM user to login account</td>
            <td valign=top nowrap>
                1. Lookup the ID number of the HP person requesting the account (this will be the sponser).<br>
                2. Update the ODM User's employee record to:<br>
                &nbsp;&nbsp;&nbsp;Active=1<br />
                &nbsp;&nbsp;&nbsp;NTName='[requested]'<br />
                &nbsp;&nbsp;&nbsp;ODMLoginStatus=1<br />
                &nbsp;&nbsp;&nbsp;ManagerID=__ (enter the Sponser's ID from step 1.)
            </td>
        </tr>
        <tr>
            <td valign=top>SI Password Changes</td>
            <td valign=top nowrap>
                1. Stop the SI Sync service and the Excalibur Scheduler job on Excal05 a couple hours before the scheduled time.  This prevents the password from getting locked out.<br>
                2. Replace all references to the old password with the new password in the web pages.<br>
                3. Change the linked server passwords for both the live and integration SI servers.<br>
                4. Restart the SI Sync service and the Excalibur Scheduler and verify that they clear out to backlog of updates.<br>
            </td>
        </tr>
        <tr>
            <td valign=top>Dave Password Changes</td>
            <td valign=top nowrap>
                1. Stop the MSMQMonitor and SIExcalSync services on Excal03 before changing the password.  This prevents the password from getting locked out.<br>
                2. Change the password.<br>
                3. Update the password for the MSMQMonitor and SIExcalSync server on Excal03 and start the services.<br>
                4. Update the password for all scheduled jobs on Excal03 starting with "PortPDP".<br>
            </td>
        </tr>
        <tr>
        		<td valign=top>Service Password Chages (SvcDClerkAM)</td>
        		<td valign=top nowrap>
        			<ol>
        				<li>Scheduled Items on HOUHPQEXCAL01</li>
        				<li>Scheduled Items on HOUHPQEXCAL03</li>
        				<li>Services on HOUHPQEXCAL03</li>
        				<li>Config files on HOUHPQEXCAL03 web and Program Files</li>
        				<li>Scheduled Items on HOUHPQEXCAL05</li>
        				<li>Services on HOUHPQEXCAL05</li>
        			</ol>
        		</td>
        </tr>
        <tr>
            <td valign=top>Delete Duplicate User Record</td>
            <td valign=top nowrap>
                1. Determine the user id for both records.<br>
                2. Run spFindEmployeeRecords using the older (or inactive) user id.<br>
                3. If any records are found (other than in the employee table), replace the old id with the new one.<br>
                4. Delete the old employee record.<br>
            </td>
        </tr>
        <tr>
            <td valign=top>IMS Service Account Password (SVCPortInvAM)</td>
            <td valign=top nowrap>
                1. Send warning email to Coordinators telling them to logout of IMS that evening.<br>
                2. Change the password.<br>
                3. Update the Availability Monitor code to use the new password.  Currently it is a VB6 app.<br>
            </td>
        </tr>
    </table>
    
</BODY>
</HTML>
