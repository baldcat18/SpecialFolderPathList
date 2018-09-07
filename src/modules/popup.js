/// <reference path="decl-userdef.d.ts" />

Element.prototype.remove = function() { this.parentNode.removeChild(this); };

/** @type {HTMLDialog} */
var dialog = this;
/** @type {DialogArgument} */
var dlgargs = dialog.dialogArguments;

document.onclick = document.oncontextmenu = function() {
	var target = event.srcElement;
	if (target.tagName == "U") target = /** @type {Element} */ (target.parentNode);
	if (target.tagName == "LI") dlgargs.sendItem(target.id);
};

/** @type {{[x: number]: string}} */
var items = {};
items[/** @type {Key.o} */ (79)] = "openFolder";
items[/** @type {Key.a} */ (65)] = "copyAsPath";
items[/** @type {Key.x} */ (88)] = "execExplorer";
items[/** @type {Key.e} */ (69)] = "execExplorerElevated";
items[/** @type {Key.p} */ (80)] = "execCmd";
items[/** @type {Key.w} */ (87)] = "execCmdElevated";
items[/** @type {Key.s} */ (83)] = "execPowershell";
items[/** @type {Key.h} */ (72)] = "execPowershellElevated";
items[/** @type {Key.l} */ (76)] = "execWsl";
items[/** @type {Key.i} */ (73)] = "execWslElevated";
items[/** @type {Key.r} */ (82)] = "showProperty";
document.onkeyup = function() {
	var evt = /** @type {KeyboardEvent} */ (event);
	
	if (evt.keyCode == /** @type {Key.escape} */ (27)) window.close();
	 
	var item = items[evt.keyCode];
	if (item) dlgargs.sendItem(item);
};

window.onload = function() {
	if (!dlgargs.explorerRunasLaunchingUser) document.getElementById("execExplorerElevated").remove();
	if (!dlgargs.isWslEnabled) removeElements(".wsl");
	if (!dlgargs.isPropertiesEnabled) document.getElementById("showProperty").remove();
	if (!dlgargs.isFileFolder) removeElements(".console");
	if (!dlgargs.extended) removeElements(".extended");
	
	var menu = document.querySelector("menu");
	dialog.dialogWidth = menu.offsetWidth + "px";
	dialog.dialogHeight = menu.offsetHeight + "px";
};
window.onbeforeunload = function() {
	dialog = null;
	dlgargs = null;
};

/** @param {string} v */
function removeElements(v) {
	var nodelist = document.querySelectorAll(v);
	
	for (var i = nodelist.length - 1; i >= 0; i--) nodelist[i].remove();
}
