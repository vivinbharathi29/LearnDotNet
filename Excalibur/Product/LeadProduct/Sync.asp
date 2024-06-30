<%@ Language=VBScript %>
	
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
    <script src="../../Scripts/PulsarPlus.js"></script>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value=="1")
		{
		    if (IsFromPulsarPlus()) {
		        window.parent.parent.parent.LeadProductSynchronizationCallback(txtSuccess.value);
		        ClosePulsarPlusPopup();
		    }
		    else {
		        window.returnValue = 1;
		        window.close();
		    }
			}
		}
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
td
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: verdana
}
</STYLE>
<BODY LANGUAGE=javascript onload="return window_onload()">
<BR>
<font size=2 face=verdana>

<%

	dim CurrentDomain
	dim CurrentUser
	
	CurrentUser = lcase(Session("LoggedInUser"))
	
	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if
	
	
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.recordset")

	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.CommandTimeout=120
	cn.Open

	Dim ActionArray
	dim strAction
	dim IDArray
	ActionArray = split(request("SyncList"),",")
	dim ProductID
    dim ReleaseID

	if Ubound(ActionArray) = -1 then
		Response.Write "Not enough info supplied to complete this request."
	else
		cn.BeginTrans
		strSuccess = "1"
		
		for each strAction in ActionArray
			IDArray=split(strAction,":")
            ProductID = IDArray(0)

			Response.Write "<b><u>ProductID: " & IDArray(0) & "</u>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<u>RootID: " & IDArray(1) & "</u></b><BR><BR>"
            
            if IDArray(2) = 0 then
			    rs.Open "Select DOTSName from productversion with (NOLOCK) where id = " & IDArray(0),cn,adOpenStatic
                if not (rs.EOF and rs.bof) then
                    Response.Write "Note: All of these actions are performed on " & rs("DotsName") & " for the deliverable version ID number listed.<BR><BR>"
			    end if
                rs.Close
                cn.Execute "spSyncRootToLeadProductLive2 " & clng(ProductID) & "," & clng(IDArray(1)) & ",'" & CurrentUser & "','" & CurrentDomain & "'"
            else                
                rs.Open "Select pv.ID, pv.DOTSName + ' (' + pvr.Name + ')' as ProductName, pv_r.ReleaseID from productversion pv with (NOLOCK) inner join ProductVersion_Release pv_r with (NOLOCK) on pv.id = pv_r.ProductVersionID inner join ProductVersionRelease pvr with (NOLOCK) on pvr.id = pv_r.ReleaseID where pv_r.id = " & IDArray(0) ,cn,adOpenStatic
                if not (rs.EOF and rs.bof) then
				    Response.Write "Note: All of these actions are performed on " & rs("ProductName") & " for the deliverable version ID number listed.<BR><BR>"
                    ProductID = rs("ID")
                    ReleaseID = rs("ReleaseID")
			    end if
                rs.Close
                cn.Execute "spSyncRootToLeadProductLive2_Pulsar " & clng(ProductID) & "," & clng(IDArray(1)) & ",'" & CurrentUser & "','" & CurrentDomain & "'," & clng(ReleaseID)
            end if
  
			if cn.Errors.count > 0 then
				strSuccess = "0"
				exit for
			end if
            
            cn.execute "usp_SSSB_SendSync_Message 'MSMQLegacy', 1, "& clng(IDArray(1))  & ",1,0,'Deliverable Root Refresh - Products Updated(ServiceBroker)'"

		next

		if strSuccess = "0" then
			cn.RollbackTrans
		else
			cn.CommitTrans
		end if
	
	end if
	
	set rs = nothing
	cn.Close
	set cn = nothing

    Sub PopulateSI (strMessage)
        Dim objQInfo
        Dim objQSend
        Dim objMessage
     
        'open the queue
        Set objQInfo = Server.CreateObject("MSMQ.MSMQQueueInfo")
        objQInfo.PathName = ".\private$\SIExcalSync" 
        Set objQSend = objQInfo.Open(2, 0)
  
        'build/send the message
        Set objMessage = Server.CreateObject("MSMQ.MSMQMessage")
        objMessage.Body = "<?xml version=""1.0""?><string>" & strMessage & "</string>"
        objMessage.Send objQSend
        objQSend.Close

        'clean up
        Set objQInfo = Nothing
        Set objQSend = Nothing
        Set objMessage = Nothing
    end sub
%>


</font>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>