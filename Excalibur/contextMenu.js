
$(document).on('click', '.contextCss', function () {
    event.preventDefault();
    leftClickContextMenu(this)
});

$(document).on('contextmenu', '.contextCss', function () {
    event.preventDefault();
    rightClickContextMenu(this)
});

$(document).bind("mousedown", function (e) {
    if (!$(e.target).parents(".custom-menu").length > 0) {
        $(".custom-menu").hide(100);
    }
});

function hideContextMenu() {
    $(".custom-menu").css('display', 'none');
}


function rightClickContextMenu(obj) {
    // Add Your Custom Menu display function here
    var node = window.event.srcElement;
    while (node.nodeName.toLowerCase() != "tr") {
        node = node.parentElement;
    }

    _productVersionID = node.PVID;
    _deliverableRootID = node.DRID;
    _serviceFamilyPn = node.SFPN;
    _spareKitID = node.SKID;
    _categoryID = node.CID;
    _productBrandID = node.PBID;
    _spareKitMapID = node.MAPID;
    showContextMenuImage(1, 2);
}

function leftClickContextMenu(obj) {
    // Add Your Custom Menu display function here
    if (obj.attributes["imgId"] != undefined) {
        var imageId = obj.attributes["imgId"].value;
        var prodDropId = obj.attributes["prodDrop"].value;
        showContextMenuImage(imageId, prodDropId);
    }
    else if (obj.attributes["dcrRootId"] != undefined) {
        var rootId = obj.attributes["dcrRootId"].value;
        var typeId = obj.attributes["dcrTypeId"].value;
        changePropertiesDcr(rootId, typeId);
    }
    else if (obj.attributes["delRootId"] != undefined) {
        var productId = $('#productId').val();
        var rootId = obj.attributes["delrootid"].value;
        var versionId = obj.attributes["versionId"].value;
        var targeted = obj.attributes["targeted"].value;
        var inImage = obj.attributes["inImage"].value;
        var categoryId = obj.attributes["CategoryID"].value;
        var typeId = obj.attributes["typeID"].value;
        var workflowComplete = obj.attributes["workflowComplete"].value;
        if (obj.attributes["accessGroup"] == undefined) {
            var accessGroup = '';
        } else {
            var accessGroup = obj.attributes["accessGroup"].value;
        }
        var seTestLead = obj.attributes["seTestLead"].value;
        var odmTestLead = obj.attributes["odmTestLead"].value;
        var wwanTestLead = obj.attributes["wwanTestLead"].value;
        var servicePM = obj.attributes["servicePM"].value;
        var fusion = obj.attributes["fusion"].value;
        var fusionRequirements = obj.attributes["fusionRequirements"].value;
        var active = obj.attributes["active"].value;
        if (obj.attributes["imagePath"] == undefined) {
            var imagePath = '';
        }
        else {
            var imagePath = obj.attributes["imagePath"].value;
        }
        if (obj.attributes["images"] == undefined) {
            var images = '';
        }
        else {
            var images = obj.attributes["images"].value;
        }
        var releaseId = $('#releaseId').val();
        var bsId = obj.attributes["BSID"].value;
        var delFilter = $('#delFilter').val()

        showDelContextMenu(productId, rootId, versionId, targeted, inImage, categoryId, typeId, workflowComplete, accessGroup, seTestLead, odmTestLead, wwanTestLead, servicePM, fusion, fusionRequirements, active, imagePath, images, releaseId, bsId, delFilter);
    }
}

function showContextMenuImage(imageId, prodDropId) {
    if (customMenuOrigImage == "") {
        customMenuOrigImage = document.getElementById("menu").innerHTML;
    }
    var customMenu = customMenuOrigImage.replace(/\[ImageID\]/g, imageId);
    customMenu = customMenu.replace(/\[ProductDropID\]/g, prodDropId);
    if (prodDropId == 0) {
        customMenu = customMenu.replace(/\[DisplayOption1\]/g, "display:none");
    }
    else {
        customMenu = customMenu.replace(/\[DisplayOption1\]/g, "");
    }

    document.getElementById("menu").innerHTML = customMenu;

    $(".custom-menu").finish().toggle(100).
    css({
        top: event.pageY + "px",
        left: event.pageX + "px"
    });

}
