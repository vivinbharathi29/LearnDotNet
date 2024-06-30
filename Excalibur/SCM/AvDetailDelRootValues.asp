<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="../includes/DataWrapper.asp" -->
<%

Dim AppRoot
AppRoot = Session("ApplicationRoot")
Dim dw, cn, cmd, rs
Dim sDelRootOpt : sDelRootOpt = ""
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim functionCalled
functionCalled = Request("Function")
If functionCalled = "" Then 
    Response.Write("No Function Called")
    Response.End
Else
    Select Case functionCalled
        Case "GetDelRootValues"
            Set cmd = dw.CreateCommAndSP(cn, "usp_SelectDeliverablesByAvFeatureCategory")
            dw.CreateParameter cmd, "@p_PVID", adInteger, adParamInput, 8, TRIM(Request("PVID"))
            dw.CreateParameter cmd, "@p_AvFeatureCategoryID", adInteger, adParamInput, 8, Request("AvFeatureCategoryID")
            Set rs = dw.ExecuteCommandReturnRS(cmd)
            sDelRootOpt = "<select id=""cboDeliverables"" name=""cboDeliverables"" style=""WIDTH: 70%""><OPTION VALUE=0>--- Please Make a Selection ---</OPTION>"
            Do Until rs.EOF
	            sDelRootOpt = sDelRootOpt & "<OPTION Value='" & rs("ID") & "'"
		        If TRIM(Request("DelRootID")) = TRIM(rs("ID")) Then
			        sDelRootOpt = sDelRootOpt & " SELECTED "
		        End If
		        sDelRootOpt = sDelRootOpt & ">" & rs("Name") & "</OPTION>" & VbCrLf
		        rs.MoveNext
	        Loop    
            sDelRootOpt = sDelRootOpt & "</select>"
            rs.Close            
            set cmd = nothing
            set cn = nothing
            set dw = nothing
            Response.Write(sDelRootOpt)
        Case Else
            Response.Write("No Function Called")
    End Select
End If

%>