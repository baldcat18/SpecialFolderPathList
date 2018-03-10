/// <reference path="decl-userdef.d.ts" />

Element.prototype.remove = function() { this.parentNode.removeChild(this); };

/** @type {typeof Key} */ 
// @ts-ignore
var Key = {
	escape: 27,
	a: 65,
	e: 69,
	h: 72,
	i: 73,
	l: 76,
	o: 79,
	p: 80,
	r: 82,
	s: 83,
	w: 87,
	x: 88
};

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
items[Key.o] = "openFolder";
items[Key.a] = "copyAsPath";
items[Key.x] = "execExplorer";
items[Key.e] = "execExplorerElevated";
items[Key.p] = "execCmd";
items[Key.w] = "execCmdElevated";
items[Key.s] = "execPowershell";
items[Key.h] = "execPowershellElevated";
items[Key.l] = "execWsl";
items[Key.i] = "execWslElevated";
items[Key.r] = "showProperty";
document.onkeyup = function() {
	var evt = /** @type {KeyboardEvent} */ (event);
	
	if (evt.keyCode == Key.escape) window.close();
	 
	/** @type {string} */
	var item = items[evt.keyCode];
	if (item) dlgargs.sendItem(item);
};

window.onload = function() {
	if (!dlgargs.explorerRunasLaunchingUser) document.getElementById("execExplorerElevated").remove();
	if (!dlgargs.isWslEnabled) removeElements(".wsl");
	if (!dlgargs.isPropertiesEnabled) document.getElementById("showProperty").remove();
	if (!dlgargs.isDirectory) removeElements(".console");
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
