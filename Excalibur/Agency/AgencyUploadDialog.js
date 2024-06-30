$(function () {
    UploadDialog.Setup();
});

var UploadDialog = function () { }
UploadDialog.Setup = function () {
    $("#AgencyUploadDocument").dialog({
        modal: true,
        autoOpen: false,
        width: 600,
        height: 350,
        show: "clip",
        buttons: {
            "Ok": function () {
                //var response = UploadDialog.Save();
                if (response) {
                    $(this).dialog("close");
                    location.reload();
                } else {
                    alert("Something went wrong.");
                }
            },
            "Cancel": function () {
                $(this).dialog("close");
            }
        }
    });

    $('#documentUpload').uploadify({
        //'hideButton': true,
        //'wmode': 'transparent',
        //'buttonText': 'Upload Documents',
        'uploader': '/uploadify/uploadify.swf',
        'script': '/uploadify/uploadify.ashx',
        'cancelImg': '/uploadify/cancel.png',
        'folder': '/uploads',
        'auto': true,
        'multi': true,
        'removeCompleted': true,
        'onComplete': function (event, ID, fileObj, response, data) {
            var docText = $('#uploadedDocuments').val();
            $('#uploadedDocuments').val(docText + ',' + fileObj.filePath);
            $("#documentList").append("<li><a href=\"#\">" + fileObj.name + "</a></li>");
        },
        'onAllComplete': function (event, data) {
            $('#startFileUpload').hide();
        },
        'onSelectOnce': function (event, data) {
            $('#startFileUpload').show();
        }

    });
}

UploadDialog.Show = function () {
    $("#AgencyUploadDocument").dialog("open");
};