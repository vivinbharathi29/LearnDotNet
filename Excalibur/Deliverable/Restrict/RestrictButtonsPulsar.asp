<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function ButtonClicked(ButtonID) {
    if (ButtonID == 1)
        if ('<%=Request("pulsarplusDivId")%>' != undefined && '<%=Request("pulsarplusDivId")%>' != "") {
            parent.window.parent.closeExternalPopup();
        }
      else
        window.parent.Cancel();
    else {
        window.parent.frames["UpperWindow"].ButtonClicked(ButtonID);
    }
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory>



<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<td><input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" LANGUAGE="javascript" onclick="return ButtonClicked(1)"></td> <!-- style="BORDER-LEFT-COLOR: tan; BORDER-BOTTOM-COLOR: tan; BORDER-TOP-COLOR: tan; BACKGROUND-COLOR: wheat; BORDER-RIGHT-COLOR: tan"-->
		<td width="10"></td>
		<td><input type="button" value="&lt;&lt; Previous" id="cmdPrevious" name="cmdPrevious" LANGUAGE="javascript" onclick="return ButtonClicked(2)" disabled></td>
		<td><input type="button" value="Next &gt;&gt;" id="cmdNext" name="cmdNext" LANGUAGE="javascript" onclick="return ButtonClicked(3)"></td>
		<td width="10"></td>
		<td><input type="button" value="Finish" id="cmdFinish" name="cmdFinish" disabled LANGUAGE="javascript" onclick="return ButtonClicked(4)"></td>
	</tr>
</table>

</BODY>
</HTML>
