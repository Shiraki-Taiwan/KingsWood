function showSubMenu() {
    var objThis = this;
    for (var i = 0; i < objThis.childNodes.length; i++) {
        if (objThis.childNodes.item(i).nodeName == "ul") {
            objThis.childNodes.item(i).style.display = "block";
        }
    }
}

function hideSubMenu() {
    var objThis = this;

    for (var i = 0; i < objThis.childNodes.length; i++) {
        if (objThis.childNodes.item(i).nodeName == "ul") {
            objThis.childNodes.item(i).style.display = "none";
            return;
        }
    }
}

function initialiseMenu() {
    var objLiCollection = document.body.getElementsByTagName("li");
    for (var i = 0; i < objLiCollection.length; i++) {
        var objLi = objLiCollection[i];
        var j;
        for (j = 0; j < objLi.childNodes.length; j++) {
            if (objLi.childNodes.item(j).nodeName == "ul") {
                objLi.onmouseover = showSubMenu;
                objLi.onmouseout = hideSubMenu;

                for (j = 0; j < objLi.childNodes.length; j++) {
                    if (objLi.childNodes.item(j).nodeName == "a") {
                        objLi.childNodes.item(j).className = "hassubmenu";
                    }
                }
            }
        }
    }
}