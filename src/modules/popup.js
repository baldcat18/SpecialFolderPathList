/// <reference path="decl-userdef.d.ts" />

Element.prototype.remove = function() { this.parentNode.removeChild(this); };

/** @type {Key.escape} */
var VK_ESCAPE = 27;

/** @type {HTMLDialog} */
var dialog = this;
/** @type {DialogArgument} */
var dlgargs = dialog.dialogArguments;

document.onclick = document.oncontextmenu = function() {
	var target = event.srcElement;
	if (target.tagName == "U") target = /** @type {Element} */ (target.parentNode);
	if (target.tagName == "LI") dlgargs.sendItem(target.id);
};

var items = dlgargs.items;
document.onkeyup = function() {
	var evt = /** @type {KeyboardEvent} */ (event);
	
	if (evt.keyCode == VK_ESCAPE) window.close();
	
	var code = String.fromCharCode(evt.keyCode);
	/** @type {DialogItem} */
	var item = null;
	for (var i = 0; i < items.length; i++) {
		if (items[i].key == code) {
			item = items[i];
			break;
		}
	}
	if (item) dlgargs.sendItem(item.id);
};

window.onload = function() {
	var menu = document.querySelector("menu");
	
	for (var i = 0; i < items.length; i++) {
		var item = items[i];
		var li = document.createElement("li");
		li.id = item.id;
		if (item.isVisible) li.className = "visible";
		li.innerHTML = item.caption + "(<u>" + item.key + "</u>)";
		menu.appendChild(li);
	}
	
	dialog.dialogWidth = menu.offsetWidth + "px";
	dialog.dialogHeight = menu.offsetHeight + "px";
};
window.onbeforeunload = function() {
	dialog = null;
	dlgargs = null;
};
