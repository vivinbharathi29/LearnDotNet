<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../../includes/noaccess.inc" -->

<HTML>
<HEAD>
<TITLE>Add Lead Product Root Deliverable Exceptions</TITLE>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script src="../../Scripts/PulsarPlus.js"></script>
<script src="../../Scripts/Pulsar2.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
    if (typeof (txtSuccess) != "undefined") {
        if (txtSuccess.value == "1") {
            if (isFromPulsar2()) {
                closePulsar2Popup(true);
            }
            else if (IsFromPulsarPlus()) {
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
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">
<BR>
<table width=100%><TR><TD align=center>
<font face=verdana size =2>Saving Exceptions.  Please wait...</font>
</td></tr></table>
<%

	dim strVersionID
	dim strProductID
	dim strType
	dim strRejected
	dim cn
	dim rs
	dim cm
	dim ExceptionArray
    dim FusionRequirements
    dim ReleaseID
    ReleaseID = 0

	if request("ExceptionList") = "" then
		Response.Write "<BR><font face=verdana size=2>Not enough information supplied</font><BR>" 
		strSuccess = 0
	else


		set cn = server.CreateObject("ADODB.Connection")
		set rs = server.CreateObject("ADODB.recordset")
	
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
		cn.Open

		cn.BeginTrans
		
		ExceptionArray = split(request("ExceptionList"),",")
		
		for i = 0 to ubound(ExceptionArray)
			if instr(ExceptionArray(i),":") = 0 then
				Response.Write "<BR>InvalidID<BR>"
				Response.write "<BR>" & request("ExceptionList") & "</BR>"
				strSuccess = "0"
				exit for
			else
                dim arrData
                arrData = split(ExceptionArray(i), ":")

				strProductID = trim(arrData(0))		
				strRootID = trim(arrData(1))
	            FusionRequirements = trim(arrData(2))

                if FusionRequirements > 0 then
                    rs.Open "Select ProductVersionID, ReleaseID from ProductVersion_Release with (NOLOCK) where id = " & trim(arrData(0)) ,cn,adOpenStatic
                    if not (rs.EOF and rs.bof) then
                        strProductID = rs("ProductVersionID")
                        ReleaseID = rs("ReleaseID")
			        end if
                    rs.Close
                end if

				set cm = server.CreateObject("ADODB.Command")
		
				cm.ActiveConnection = cn
				cm.CommandText = "spAddLeadProductRootExclusion"
				cm.CommandType = &H0004
						
				Set p = cm.CreateParameter("@ProductVersionID",adInteger, &H0001)
				p.Value = clng(strProductID)
				cm.Parameters.Append p
			
				Set p = cm.CreateParameter("@DeliverableRootID",adInteger, &H0001)
				p.Value = clng(strRootID)
				cm.Parameters.Append p
					
                Set p = cm.CreateParameter("@ReleaseID",adInteger, &H0001)
				p.Value = clng(ReleaseID)
				cm.Parameters.Append p

				cm.Execute
				Set cm = Nothing
	
				if cn.Errors.count > 0 then
					strSuccess = "0"
					cn.RollbackTrans
					exit for
				else
					strSuccess = "1"
				end if
				Response.Write "<BR>DONE<BR>"
			end if
		next
		if strSuccess = "1" then
				cn.CommitTrans
		end if
					            
		set rs = nothing
		set cn = nothing


	end if
%>

<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>
