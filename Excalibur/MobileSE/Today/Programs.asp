<%@ Language=VBScript %>
<%
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
%>
<html>
<head>
<title>
<%
if Request("ID") <> "" then
	Response.Write "Product Properties"
else
	Response.Write "Add New Product"
end if
%>
</title>
</head>
<frameset ROWS="*,55" ID=TopWindow>
	<% if Request("Commodity") = 1 and Request("app") = "PulsarPlus"  then %>
		<frame frameborder="0" ID="UpperWindow" Name="UpperWindow" SRC="programsCommodityPM.asp?ID=<%=Request("ID")%>&app=<%=Request("app")%>"">
		<frame frameborder="0" ID="LowerWindow" Name="LowerWindow" SRC="programbuttons.asp?Commodity=<%=Request("Commodity")%>&ID=<%=Request("ID")%>&app=<%=Request("app")%>" scrolling=no>
    <% elseif Request("Commodity") = 1 then %>
		<frame frameborder="0" ID="UpperWindow" Name="UpperWindow" SRC="programsCommodityPM.asp?ID=<%=Request("ID")%>&app=<%=Request("app")%>">
		<frame frameborder="0" ID="LowerWindow" Name="LowerWindow" SRC="programbuttons.asp?Commodity=<%=Request("Commodity")%>&ID=<%=Request("ID")%>&app=<%=Request("app")%>" scrolling=no>
	<% elseif Request("HWPM") = 1 then %>
		<frame frameborder="0" ID="UpperWindow" Name="UpperWindow" SRC="programsHWPM.asp?ID=<%=Request("ID")%>&app=<%=Request("app")%>">
		<frame frameborder="0" ID="LowerWindow" Name="LowerWindow" SRC="programbuttons.asp?HWPM=<%=Request("HWPM")%>&ID=<%=Request("ID")%>&app=<%=Request("app")%>" scrolling=no>
	<% elseif Request("FactoryEngineer") = 1 then %>
		<frame frameborder="0" ID="UpperWindow" Name="UpperWindow" SRC="programFactoryEngineer.asp?ID=<%=Request("ID")%>&app=<%=Request("app")%>">
		<frame frameborder="0" ID="LowerWindow" Name="LowerWindow" SRC="programbuttons.asp?FactoryEngineer=<%=Request("FactoryEngineer")%>&app=<%=Request("app")%>&ID=<%=Request("ID")%>" scrolling=no>
	<% elseif Request("Accessory") = 1 then %>
		<frame frameborder="0" ID="UpperWindow" Name="UpperWindow" SRC="programAccessory.asp?ID=<%=Request("ID")%>&app=<%=Request("app")%>">
		<frame frameborder="0" ID="LowerWindow" Name="LowerWindow" SRC="programbuttons.asp?Accessory=<%=Request("Accessory")%>&app=<%=Request("app")%>&ID=<%=Request("ID")%>" scrolling=no>
    <% elseif Request("Pulsar") = 1 then %>
		<frame frameborder="0" ID="UpperWindow" Name="UpperWindow" SRC="programmain_Pulsar.asp?ID=<%=Request("ID")%>&app=<%=Request("app")%>&clone=<%=Request("Clone")%>&Tab=<%=request("Tab")%>">
		<frame frameborder="0" ID="LowerWindow" Name="LowerWindow" SRC="programbuttons_Pulsar.asp?ID=<%=Request("ID")%>&app=<%=Request("app")%>&clone=<%=Request("Clone")%>" scrolling=no>
	<% else
		'No parameter passed so check to see what kind of Product it is in case all the old Excalibur links were not all changed to pass the Pulsar parameter for a Pulsar Product.'
        dim IsPulsar
    	If Request("ID") = "" Then
            IsPulsar = False
        else
            dim cn, rs, cm
		    set cn = server.CreateObject("ADODB.Connection")
		    cn.ConnectionString = Session("PDPIMS_ConnectionString")
		    cn.Open
		    set rs = server.CreateObject("ADODB.recordset")
		    With rs
			    .ActiveConnection = cn
			    .CursorType = adOpenForwardOnly
			    .LockType=AdLockReadOnly
		    End With
		    set cm = server.CreateObject("ADODB.Command")
		    With cm
			    .ActiveConnection = cn
			    .CommandType = adCmdStoredProc
			    .CommandText = "usp_IsPulsarProduct"
			    .Parameters.Append .CreateParameter("@p_intPVID",adInteger, adParamInput)
			    .Parameters("@p_intPVID") = Request("ID")
			    set rs = .Execute
		    End With

		    if not (rs.EOF and rs.BOF) then
			    IsPulsar=rs("IsPulsar")
		    end if
		    rs.Close
		    set rs = nothing
		    cn.Close
		    set cn = nothing
		    set cm=nothing
        End If

		if IsPulsar then %>
			<frame frameborder="0" ID="UpperWindow" Name="UpperWindow" SRC="programmain_Pulsar.asp?ID=<%=Request("ID")%>&clone=<%=Request("Clone")%>&app=<%=Request("app")%>&Tab=<%=request("Tab")%>&Pulsar=1">
			<frame frameborder="0" ID="LowerWindow" Name="LowerWindow" SRC="programbuttons_Pulsar.asp?ID=<%=Request("ID")%>&clone=<%=Request("Clone")%>&app=<%=Request("app")%>&Pulsar=1" scrolling=no>
		<% else %>
			<frame frameborder="0" ID="UpperWindow" Name="UpperWindow" SRC="programmain.asp?ID=<%=Request("ID")%>&clone=<%=Request("Clone")%>&app=<%=Request("app")%>&Tab=<%=request("Tab")%>">
			<frame frameborder="0" ID="LowerWindow" Name="LowerWindow" SRC="programbuttons.asp?ID=<%=Request("ID")%>&app=<%=Request("app")%>&clone=<%=Request("Clone")%>" scrolling=no>
		<% end if %>
	<% end if %>
</frameset>
</html>