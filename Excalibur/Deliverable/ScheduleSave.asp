<%@ Language=VBScript %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
	{
		if (txtSuccess.value == "1")
		{
		    if (IsFromPulsarPlus()) {
		        window.parent.parent.parent.popupCallBack(1);
		        ClosePulsarPlusPopup();
		    }
		    else {
		        window.returnValue = 1;
		        var iframeName = parent.window.name;
		        if (iframeName != '') {
		            parent.window.parent.ClosePopUp();
		        } else {
		            window.parent.close();
		        }
		    }
		}
		else
			document.write ("<BR><font size=2 face=verdana>Unable to update deliverable schedule.</font>");
	}
	else
		document.write ("<BR><font size=2 face=verdana>Unable to update deliverable schedule.</font>");
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%
	dim strMilestones
	dim strTags
	dim strDates
	dim TagArray
	dim IDArray
	dim DateArray
	dim i
	dim blnErrors
	dim RowsChanged
	dim strSuccess
	
	strSuccess = ""
	
	strMilestones = request("txtMilestones")
	strTags = request("tagDate") & ", " & request("tagRelease")
	strDates = request("txtDate") & ", " & request("txtRelease")
	
	IDArray = split(strMilestones,",")
	DateArray = split(strDates,", ")
	TagArray = split(strTags,", ")
	

	if request("txtDate") = "" then
		strSuccess = "1"
	else

		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open
		
		cn.BeginTrans
		blnErrors = false
		
		for i = lbound(IDArray) to ubound(IDarray)
			if formatdatetime(TagArray(i),vbshortdate) <> formatdatetime(DateArray(i),vbshortdate) then
				Response.Write "Changed " & IDArray(i) & ": " & TagArray(i) & "-" & DateArray(i) & "<BR>"			
				cn.execute "spUpdateMilestonePlan " & IDArray(i) & ",'" & formatdatetime(DateArray(i),vbshortdate) &  "'" ,RowsChanged 
				if RowsChanged <> 1 then
					blnErrors = true
					exit for
				end if
			end if
		next
		
		if blnErrors then
			cn.RollbackTrans
		else
			cn.execute "spUpdateDeliverableLocation " & request("txtDisplayedID")
			if cn.Errors.count > 0 then
				cn.RollbackTrans
			else
				cn.CommitTrans
				strSuccess = "1"
			end if
		end if
		
		cn.Close
		set cn = nothing
		
	end if



%>


<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">


</BODY>
</HTML>
