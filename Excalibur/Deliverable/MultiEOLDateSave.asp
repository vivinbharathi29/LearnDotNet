<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<script src="../Scripts/jquery-1.10.2.js"></script>
<script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload() {
        if (typeof (txtSuccess) != "undefined") {
            if (txtSuccess.value == "1") {
                if (IsFromPulsarPlus()) {
                    ClosePulsarPlusPopup();
                    window.parent.parent.parent.ComponentEndOfLifeDateExpiredReloadCallback(1);
                }
                else {
                    window.returnValue = "1";
                    window.parent.close();
                }
            }
            else
                document.write("<BR><font size=2 face=verdana>Unable to update deliverable Availablity Information.</font>");
        }
        else
            document.write("<BR><font size=2 face=verdana>Unable to update deliverable Availablity Information.</font>");
    }


//-->
</SCRIPT>

<BODY LANGUAGE=javascript onload="return window_onload()">
<%
	dim strSuccess
	strSuccess = "1"

	dim cn
	dim cm
	dim blnErrors
	dim IDArray
	dim ItemID
    dim rowschanged
	
	IDArray = split(request("lstID"),",")
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
		
	cn.BeginTrans
		
	for each ItemID in IDArray
		if isnumeric(ItemID) and trim(ItemID) <> "" then
            if trim(request("cboDateChange")) = "3" then
			    set cm = server.CreateObject("ADODB.Command")
			    cm.CommandType =  &H0004
			    cm.ActiveConnection = cn
			    cm.CommandText = "spSyncService2FactoryEOA"	
    	
			    Set p = cm.CreateParameter("@ID", 3,  &H0001)
			    p.Value = clng(ItemID)
			    cm.Parameters.Append p
        	
			    cm.Execute rowschanged

			    set cm=nothing
        	
        	else
			    set cm = server.CreateObject("ADODB.Command")
			    cm.CommandType =  &H0004
			    cm.ActiveConnection = cn
			    if trim(request("txtTypeID"))="2" then
				    cm.CommandText = "spUpdateDeliverableServiceEOL"	
			    else
				    cm.CommandText = "spUpdateDeliverableEOL"	
			    end if
    	
			    Set p = cm.CreateParameter("@ID", 3,  &H0001)
			    p.Value = clng(ItemID)
			    cm.Parameters.Append p

			    Set p = cm.CreateParameter("@EOLDate", 135,  &H0001)
			    if trim(request("cboDateChange")) <> "1" then
				    p.Value = null
			    elseif request("txtEOLDate") = "" then
				    p.Value = null
			    else
				    p.Value = cdate(request("txtEOLDate"))
			    end if
			    cm.Parameters.Append p

			    Set p = cm.CreateParameter("@Active", 11,  &H0001)
			    if trim(request("cboDateChange")) = "2" then
				    p.Value = false
			    else
				    p.Value = true
			    end if
			    cm.Parameters.Append p

			    cm.Execute rowschanged

			    set cm=nothing
            end if
		
			if rowschanged <> "-1" then
				strSuccess = "0"
				exit for
			end if
			
		end if
	next
	
	
	
	if strSuccess="1" then
		cn.CommitTrans
	else
		cn.RollbackTrans
	end if
	
	cn.Close
	set cn = nothing
	
	
	
	
%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">

</BODY>
</HTML>

