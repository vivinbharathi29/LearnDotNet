<%@ Language=VBScript %>

	<%
    if request("ID") <> "" then
            'Response.Redirect "/Excalibur/PMR/softpaqframe.asp?Title=Component Version Properties&Page=/Pulsar/Component/EditVersionProperties&ComponentVersionId=" & request("ID") & "&app=" & request("app")
            Response.Redirect "/Pulsar/Component/EditVersionProperties?ComponentVersionId=" & request("ID") & "&app=" & request("app")
    elseif request("RootID") <> "" then
            'Response.Redirect "/Excalibur/PMR/softpaqframe.asp?Title=Component Version Properties&Page=/Pulsar/Component/CreateVersionProperties&ComponentRootId=" & request("RootID") & "&app=" & request("app")
            Response.Redirect "/Pulsar/Component/CreateVersionProperties?ComponentRootId=" & request("RootID") & "&app=" & request("app")
    elseif request("CloneID") <> "" then
            'Response.Redirect "/Excalibur/PMR/softpaqframe.asp?Title=Component Version Properties&Page=/Pulsar/Component/CloneVersionProperties&ComponentVersionId=" & request("CloneID") & "&app=" & request("app")
            Response.Redirect "/Pulsar/Component/CloneVersionProperties?ComponentVersionId=" & request("CloneID") & "&app=" & request("app")
    end if


    Dim AppRoot : AppRoot = Session("ApplicationRoot")
    
    Response.Buffer = True
    Response.ExpiresAbsolute = Now() - 1
    Response.Expires = 0
    Response.CacheControl = "no-cache"
	%>

