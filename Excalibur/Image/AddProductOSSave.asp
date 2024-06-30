<%@ Language=VBScript %>

<% Option Explicit%>
<!-- #include file="../includes/emailwrapper.asp" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script type="text/javascript" src="../includes/client/jquery.min.js"></script>
<script type="text/javascript" src="../includes/client/json2.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--



function window_onload() {
	var OutArray = new Array();

	if (txtSuccess.value == "1") {

	    OutArray = txtOutput.value.split("|")

	    if (parent.window.parent.document.getElementById('modal_dialog')) {
	        //save array value and return to parent page: ---
	        parent.window.parent.modalDialog.passArgument(JSON.stringify(OutArray), 'productOS_query_array');
	        parent.window.parent.cmdAddOSResult();
	        parent.window.parent.modalDialog.cancel();
	    } else {
	        window.returnValue = OutArray;
	        window.parent.close();
	    }
	}
	else {
	    document.write("Unable to save selections.");
	}
		
}
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">


<%


	dim i
	dim cn
	dim cm
	dim p
	dim rowschanged
	dim blnFailed
	dim rs	
    dim strID
    dim OSArray
    dim oMessage
    dim strOutputOSList
    
    strID = clng(request("txtProductID"))
    blnFailed = false
    
	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open


	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open


	'Get User
	dim CurrentUser
	dim CurrentUserID
	dim CurrentDomain
	dim CurrentUserPartner
    dim CurrentUserEmail
    
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")

	cm.CommandType = 4
	cm.CommandText = "spGetUserInfo"
	

	Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
	p.Value = Currentuser
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	p.Value = CurrentDomain
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	set cm=nothing	

	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") 
		CurrentUserEmail = rs("Email") 
	end if
	rs.Close

'-------------------------------------------------------------------
    dim strOldOSPreinstallList
    dim strOldOSWebList

	strOldOSPreinstallList=""
	strOldOSWebList=""

	rs.open "spListProductOS " & strID,cn,adOpenForwardOnly
	do while not rs.eof
		if rs("Preinstall") then
		 strOldOSPreinstallList = strOldOSPreinstallList & ", " & rs("shortname")
		end if
		if rs("Web") then
		 strOldOSWebList = strOldOSWebList & ", " & rs("shortname")
		end if
		rs.movenext
	loop
	rs.close
		
   	cn.BeginTrans
	
	OSArray = split(request("chkOS"),",")
	for i = lbound(OSArray) to ubound(OSArray)
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spUpdateProductOSList"
			
		Set p = cm.CreateParameter("@ProductID", 3, &H0001)
		p.Value = strID
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@OSID", 3, &H0001)
		p.Value = OSArray(i)
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@Preinstall", 11, &H0001)
		p.Value = 1
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Web", 11, &H0001)
		p.Value = 0
		cm.Parameters.Append p
		
		cm.execute
		set cm = nothing
		
        response.write 	strID & "-" & OSArray(i) & "<BR>"
		if cn.Errors.count <> 0 then
            response.write "<BR>Failed"
			blnFailed = true
			exit for
		end if		
	next
		
    if blnFailed then
    	cn.RollbackTrans
		%><INPUT type="text" id=txtSuccess name=txtSuccess value="0"><%
    else
	    cn.CommitTrans
	    %><INPUT type="text" id=txtSuccess name=txtSuccess value="1"><%
    end if

    strOutputOSList = ""
	for i = lbound(OSArray) to ubound(OSArray)
        rs.open "spGetOSbyID " & OSArray(i),cn
        if not(rs.eof and rs.bof) then
            strOutputOSList = strOutputOSList & "|" & rs("ID") & "^" & rs("Name")
        end if
        rs.close
    next
    if strOutputOSList <> "" then
        strOutputOSList = mid(strOutputOSList,2)
    end if
    
        if request("chkOS") <> "" and not blnFailed then
            dim strSEPM
		    dim strNewOSPreinstallList
		    dim strNewOSWebList
		    
		    strSEPM = ""
		    rs.open "spListSystemTeam " & strID,cn
		    do while not rs.eof
		        if rs("Role") = "SE PM" then
		           strSEPM = rs("Email")
		           exit do 
		        end if
		        rs.movenext
		    loop
		    rs.close
		    strNewOSPreinstallList=""
		    strNewOSWebList=""
		    rs.open "spListProductOS " & strID,cn,adOpenForwardOnly
		    do while not rs.eof
			    if rs("Preinstall") then
			     strNewOSPreinstallList = strNewOSPreinstallList & ", " & rs("shortname")
			    end if
			    if rs("Web") then
			     strNewOSWebList = strNewOSWebList & ", " & rs("shortname")
			    end if
			    rs.movenext
		    loop
		    rs.close		
    		
		    if strNewOSPreinstallList <> "" then
			    strNewOSPreinstallList = mid(strNewOSPreinstallList,3)
		    end if
		    if strNewOSWebList <> "" then
			    strNewOSWebList = mid(strNewOSWebList,3)
		    end if
		    if strOldOSPreinstallList <> "" then
			    strOldOSPreinstallList = mid(strOldOSPreinstallList,3)
		    end if
		    if strOldOSWebList <> "" then
			    strOldOSWebList = mid(strOldOSWebList,3)
		    end if
    		
		    dim strOSChangedBody
		    strOSChangedBody = "<font face=Arial size=2 color=black>The " & request("txtProductName") & " OS list has been changed:<BR><BR>"
		    if strNewOSPreinstallList <> strOldOSPreinstallList then
			    strOSChangedBody=strOSChangedBody & "<b><u>Preinstall</u></b><BR><b>Old:</b> " & strOldOSPreinstallList & "<BR><b>New:</b> " & strNewOSPreinstallList & "<BR><BR>"
		    end if
		    if strNewOSWebList <> strOldOSWebList then
			    strOSChangedBody=strOSChangedBody &  "<b><u>Web</u></b><BR><b>Old:</b> " & strOldOSWebList & "<BR><b>New:</b> " & strNewOSWebList & "<br><br>"
		    end if
    		
		    response.write strNewOSPreinstallList & "<BR><BR>" & strOldOSPreinstallList
		    if strNewOSPreinstallList <> strOldOSPreinstallList or strNewOSWebList <> strOldOSWebList then
            
			    Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
			    'Set oMessage.Configuration = Application("CDO_Config")
			    oMessage.From = Currentuseremail
			    if trim(strID) = "100" then
				    oMessage.To= "max.yu@hp.com"
  		            oMessage.CC = "max.yu@hp.com"'strSEPM
			    else
			    	oMessage.To= "houreleasecoordinatorsswrel@hp.com;max.yu@hp.com"
			       if strSEPM <> "" then
			         oMessage.CC = strSEPM
			       end if
			    end if
			    oMessage.Subject = "Product OS List Updated in Excalibur" 
    	
			    oMessage.HTMLBody = strOSChangedBody & "</font>" 
    			
			    oMessage.Send 
			    Set oMessage = Nothing 			
		    end if
        end if		

'------------------------------------------------------------------
	
	cn.Close
	set rs = nothing
	set cn = nothing
	set p = nothing
	
%>
<INPUT type="hidden" id=txtOutput name=txtOutput value="<%=strOutputOSList%>">
</BODY>
</HTML>

