<%@  language="VBScript" %>
<!--#include file="../includes/DataWrapper.asp"-->
<!--#include file="../includes/no-cache.asp"-->
<% 
    Dim regEx
    Set regEx = New RegExp
    regEx.Global = True

    regEx.Pattern = "[^0-9]"
    Dim PVID : PVID = regEx.Replace(Request.QueryString("PVID"), "")
    Dim DRID : DRID = regEx.Replace(Request.QueryString("DRID"), "")
    Dim SKID : SKID = regEx.Replace(Request.QueryString("SKID"), "")
    Dim CID : CID = regEx.Replace(Request.QueryString("CID"), "")
    regEx.Pattern = "[^0-9-]"
    Dim SFPN : SFPN = trim(Request.QueryString("SFPN"))
    

	Dim pageTitle
    Dim rs, dw, cn, cmd
    Set rs = Server.CreateObject("ADODB.RecordSet")
    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

    If SKID <> "" Then
        Set cmd = dw.CreateCommandSp(cn, "usp_SelectSpareKitDetails")
        dw.CreateParameter cmd, "@p_SpareKitId", adInteger, adParamInput, 0, CLng(SKID)
        dw.CreateParameter cmd, "@p_SpareKitNo", adChar, adParamInput, 10, ""
        dw.CreateParameter cmd, "@p_ServiceFamilyPn", adChar, adParamInput, 10, SFPN
        Set rs = dw.ExecuteCommandReturnRs(cmd)
        If Not rs.Eof Then
            pageTitle = rs("Description") & ""
            If(Len(Trim(pageTitle))>40) Then
                pageTitle=Mid(Trim(pageTitle),1,40)
            End If
        End If
        rs.close
    ElseIf DRID <> "" Then
        Set cmd = dw.CreateCommandSp(cn, "spGetDeliverableRootName")
        dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 0, CLng(DRID)
        Set rs = dw.ExecuteCommandReturnRs(cmd)
        If Not rs.Eof Then
            pageTitle = rs("Name") & ""
        End If
        rs.close
    End If

    If pageTitle <> "" Then 
        pageTitle = "Spare Kit Details - " & pageTitle
    Else
        pageTitle = "Spare Kit Details"
    End If

'*********************************************************************************************************************************************************************************************************************
' LIMIT OSSP USERS (PARTNERTYPEID=2) TO READ ONLY ACCESS
'
'*********************************************************************************************************************************************************************************************************************
Dim CurrentUser : CurrentUser = lcase(trim(Session("LoggedInUser")))
Dim CurrentDomain
Dim CurrentPartnerTypeID
Dim blnIsOSSPUser: blnIsOSSPUser=False

If instr(CurrentUser,"\") > 0 Then
	CurrentDomain = left(CurrentUser, instr(CurrentUser,"\") - 1)
	CurrentUser = mid(CurrentUser,instr(CurrentUser,"\") + 1)
End If

Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Set cmd = dw.CreateCommandSP(cn, "usp_GetUserType")
dw.CreateParameter cmd, "@UserName", adVarchar, adParamInput, 30, CurrentUser
dw.CreateParameter cmd, "@Domain", adVarchar, adParamInput, 30, CurrentDomain
Set rs = dw.ExecuteCommandReturnRS(cmd)

If not (rs.EOF And rs.BOF) Then

	CurrentPartnerTypeID=trim(rs("PartnerTypeID"))

	If(CurrentPartnerTypeID = "2" OR CurrentPartnerTypeID = 2) Then
		blnIsOSSPUser=true	
	End If
	
	rs.Close
End If

set rs=nothing

cn.Close
set cn=nothing
set cmd=nothing

'*********************************************************************************************************************************************************************************************************************
%>
<html>
<head>
    <title><%=pageTitle %></title>
</head>
<frameset rows="*,55" id="TopWindow">
<frame id="UpperWindow" name="UpperWindow" src="SpareKitDetailsMain.asp?PVID=<%=PVID%>&DRID=<%=DRID%>&SKID=<%=SKID%>&SFPN=<%=SFPN%>&CID=<%=CID%>">
<% If (Not blnIsOSSPUser) Then %>
<frame ID="LowerWindow" name="LowerWindow" src="SpareKitDetailsButtons.asp?PVID=<%=PVID%>&DRID=<%=DRID%>&SKID=<%=SKID%>&SFPN=<%=SFPN%>&CID=<%=CID%>&M=0" scrolling="no">
<% Else %>
<frame ID="LowerWindow" name="LowerWindow" src="SpareKitDetailsButtons.asp?PVID=<%=PVID%>&DRID=<%=DRID%>&SKID=<%=SKID%>&SFPN=<%=SFPN%>&CID=<%=CID%>&M=1" scrolling="no">
<% End If %>
</frameset>
</html>
