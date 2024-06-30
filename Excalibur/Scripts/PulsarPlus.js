
//To return the query string value.
function GetQueryStringValue(theParameter) {
    var queryString = "";
    if (window.parent.parent.parent.document.getElementById('externalIframe') != null) {
        queryString = window.parent.parent.parent.document.getElementById('externalIframe').src;
    }else if (window.parent.parent.document.getElementById('externalIframe') != null) {
        queryString = window.parent.parent.document.getElementById('externalIframe').src;
    }

    if (queryString != "") {
        //var params = string.substr(1).split('&');
        if (queryString.indexOf('?') > 0) {
            var params = queryString.split('?')[1].split('&');
            for (var i = 0; i < params.length; i++) {
                var p = params[i].split('=');
                if (p[0] == theParameter) {
                    return decodeURIComponent(p[1]);
                }
            }
        }
        else if (queryString.indexOf('/pulsarplus/') > 0) {
            return 'pulsarplus'
        }
    }
    return "";
}

//To test whether the pop up is called from pulsarplus
function IsFromPulsarPlus() {    
    var data = GetQueryStringValue('app');
    if (data.toLowerCase() == 'pulsarplus') {
        return true;
    }
    else {
        return false;
    }
}
//To close the pop up called from pulsarplus.
function ClosePulsarPlusPopup() {
    if (window.parent.parent.document.getElementById('externalIframe') != null) {
        window.parent.parent.document.getElementById('externalIframe').src = "";
    }
    else if (window.parent.parent.parent.document.getElementById('externalIframe') != null) {
        window.parent.parent.parent.document.getElementById('externalIframe').src = "";
    }
    if (window.parent.parent.document.getElementById("idexternalclose") != null) {
        window.parent.parent.document.getElementById("idexternalclose").click();
    } else if (window.parent.parent.parent.document.getElementById("idexternalclose") != null) {
        window.parent.parent.parent.document.getElementById("idexternalclose").click();
    }
}
