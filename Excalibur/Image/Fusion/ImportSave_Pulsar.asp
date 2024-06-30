<%@ Language=VBScript %>
<!-- #include file="../../includes/DataWrapper.asp" -->
<!-- #include file = "../../includes/Security.asp" -->
<%
    Dim Security
    Dim m_UserFullName
	
	
	Set Security = New ExcaliburSecurity   

    m_UserFullName = Security.CurrentUserFullName()  
    Set Security = Nothing  
 
    
      dim strSuccess
	dim ImportArray
	dim i
	dim cn
	dim rs
	dim cm
	dim NewID
	dim imagerow
    dim p
    
	strSuccess = "1"
    
	
	if trim(request("chkSelected")) <> "" then
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
        cn.ConnectionTimeout = 300
		cn.Open
	
'		set rs = server.CreateObject("ADODB.recordset")
	
		ImportArray = split(request("chkSelected"),",")
	
		cn.BeginTrans
		for i = lbound(ImportArray) to ubound(ImportArray)
	        imagerow = split(ImportArray(i),"-")
            
            set cm = server.CreateObject("ADODB.Command")
			cm.CommandType =  &H0004
			cm.CommandTimeout = 300
            cm.ActiveConnection = cn
		
			cm.CommandText = "usp_Image_ImportImageDefinition"	

			Set p = cm.CreateParameter("@p_intImageID", 3,  &H0001)
			p.Value = clng(imagerow(0))
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@p_intProductID", 3,  &H0001)
			p.Value = clng(request("txtID"))
			cm.Parameters.Append p

            Set p = cm.CreateParameter("@p_intFeatureID", 3,  &H0001)
			p.Value = clng(imagerow(1))
			cm.Parameters.Append p

            Set p = cm.CreateParameter("@p_intAlreadyInProduct", 3,  &H0001)
			p.Value = clng(imagerow(2))
			cm.Parameters.Append p

            Set p = cm.CreateParameter("@p_intImportLocalizations", 3,  &H0001)
			p.Value = clng(request("txtImportLocalizations"))
			cm.Parameters.Append p

            Set p = cm.CreateParameter("@p_intSourceProductID", 3,  &H0001)
			p.Value = clng(request("cboProduct"))
			cm.Parameters.Append p

            Set p = cm.CreateParameter("@p_chrUser", 200, &H0001, 80)
			p.Value = m_UserFullName
			cm.Parameters.Append p
            
            Set p = cm.CreateParameter("@p_ReleaseID", 200, &H0001, 80)
			p.Value = clng(request("cboRelease"+ ImportArray(i)))
			cm.Parameters.Append p
			

			Set p = cm.CreateParameter("@p_intNewID", 3,  &H0002)
			cm.Parameters.Append p

			cm.Execute rowschanged

			NewID = cm("@p_intNewID")
			set cm=nothing

			if cn.Errors.count > 0 then
				strSuccess = "0"
				exit for
			end if
		next
	
		if strSuccess = "0" then
			cn.RollbackTrans
		else
			cn.CommitTrans
		end if
	
'		set rs = nothing
		cn.Close
		set cn=nothing
	end if
	

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <title ></title>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload(pulsarplusDivId) {
        if (typeof (txtSuccess) != "undefined") {
            if (txtSuccess.value == "1") {
                if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                    parent.window.parent.closeExternalPopup();
                    parent.window.parent.reloadFromPopUp(pulsarplusDivId);                    
                }
                else {
                    parent.window.parent.ClosePropertiesDialog(txtSuccess.value);
                }
            }
            else
                document.write("<BR><font size=2 face=verdana>Unable to import the image list.</font>");
        }
        else
            document.write("<BR><font size=2 face=verdana>Unable to import the image list.</font>");
    }

    //-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript  onload="return window_onload('<%=Request("pulsarplusDivId")%>')">
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>
