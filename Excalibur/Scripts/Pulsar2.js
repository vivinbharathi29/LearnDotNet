//Verifiy If called from Pulsar2
function isFromPulsar2() {
    return (parent.window.frameElement != undefined && (parent.window.frameElement.id == "FrameElementInPulsar2" || parent.parent.window.frameElement.id == "FrameElementInPulsar2"));
}

//Close Pulsar2 popup
function closePulsar2Popup(canReloadTileGrid) {
    if (parent.window.frameElement.id == "FrameElementInPulsar2") {
        window.parent.parent.closePopupExternal(canReloadTileGrid);
    }
    else {
        parent.window.parent.parent.closePopupExternal(canReloadTileGrid);
    }
    return;
}

//Close Pulsar2 popup and call callAfterClose method
function closePulsar2PopupAndCallDelegate(callDelegateAfterClose) {
    if (parent.window.frameElement.id == "FrameElementInPulsar2") {
        window.parent.parent.closePulsar2PopupAndCallDelegate(callDelegateAfterClose);
    }
    else {
        parent.window.parent.parent.closePulsar2PopupAndCallDelegate(callDelegateAfterClose);
    }

    return;
}

function intialOfferings(returnValue) {
    if ($('#preferredLayout').val() == 'pulsar2') {
        if (returnValue != undefined) {
            var returnValue = returnValue.split(",");
        }
        if (returnValue[0] != 0) {
            location.href = "/Excalibur/InitialOffering.aspx?Business=" + returnValue[0] + "&Category=" + returnValue[1];
        }
        else {
            location.href = "/Excalibur/InitialOffering.aspx?ProductProgram=" + returnValue[1] + "&Category=" + returnValue[2] + "&ProgramText=" + returnValue[3];
        }
    }
}