/// <reference path="htamain.js" />

"use strict";

const VK_RETURN = 13;

const isExplorerRunasLaunchingUser = !getRegValue("HKCR\\AppID\\{CDCBCFCA-3CDC-436f-A4E2-0E02075250C2}\\RunAs");
const isWslEnabled = fso.FileExists(getSystemPath() + "\\wsl.exe");

const command = {
	/** @param {HTMLAnchorElement} target */
	openFolder: function(target) { target.xFolder.open(); },
	/** @param {HTMLAnchorElement} target */
	copyAsPath: function(target) { clipboardData.setData("Text", target.innerHTML); },
	/** @param {HTMLAnchorElement} target */
	execExplorer: function(target) { target.xFolder.execExplorer(); },
	/** @param {HTMLAnchorElement} target */
	execExplorerElevated: isExplorerRunasLaunchingUser ? function(target) { target.xFolder.execExplorer("runas"); } :
		function() { writeError("エクスプローラーを管理者として実行できる設定になっていません。"); },
	/** @param {HTMLAnchorElement} target */
	execCmd: function(target) { target.xFolder.execCmd(); },
	/** @param {HTMLAnchorElement} target */
	execCmdElevated: function(target) { target.xFolder.execCmd("runas"); },
	/** @param {HTMLAnchorElement} target */
	execPowershell: function(target) { target.xFolder.execPowershell(); },
	/** @param {HTMLAnchorElement} target */
	execPowershellElevated: function(target) { target.xFolder.execPowershell("runas"); },
	/** @param {HTMLAnchorElement} target */
	execWsl: function(target) { target.xFolder.execWsl(); },
	/** @param {HTMLAnchorElement} target */
	execWslElevated: function(target) { target.xFolder.execWsl("runas"); },
	/** @param {HTMLAnchorElement} target */
	showProperty: function(target) { target.xFolder.showProperties(); },
};
if (!isWslEnabled) {
	command.execWsl = command.execWslElevated =
		function() { writeError("WSL がサポートされていないか有効化されていません。"); };
}

const popup = (function() {
	/** @type {HTMLAnchorElement} */
	let target = null;
	/** @type {Window} */
	let dialog = null;
	/** @type {DialogItem[]} */
	const items = [
		{ id: "openFolder", caption: "開く(&O)", isAlwaysVisible: true, },
		{ id: "copyAsPath", caption: "パスのコピー(&A)", isAlwaysVisible: true, },
		{ id: "execExplorer", caption: "エクスプローラー(&X)", isAlwaysVisible: false, },
		{ id: "execExplorerElevated", caption: "エクスプローラーを管理者として開く(&E)", isExtended: true, isAlwaysVisible: (isExplorerRunasLaunchingUser ? null : false)},
		{ id: "execCmd", caption: "コマンドプロンプトを開く(&P)", isConsole:true, },
		{ id: "execCmdElevated", caption: "コマンドプロンプトを管理者として開く(&W)", isConsole:true, isExtended: true, },
		{ id: "execPowershell", caption: "Windows PowerShell を開く(&S)", isConsole:true, },
		{ id: "execPowershellElevated", caption: "Windows PowerShell を管理者として開く(&H)", isConsole:true, isExtended: true, },
		{ id: "execWsl", caption: "Linux シェルを開く(&L)", isConsole:true, isAlwaysVisible: (isWslEnabled ? null : false), },
		{ id: "execWslElevated", caption: "Linux シェルを管理者として開く(&I)", isConsole:true, isExtended: true, isAlwaysVisible: (isWslEnabled ? null : false), },
		{ id: "showProperty", caption: "プロパティ(&R)", },
	];
	/** @type {DialogArgument} */
	const dlgargs = {
		items: items,
		/** @param {string} item */
		sendItem: function(item) {
			popup.hide();
			
			// setTimeoutを使うことで、ちゃんとポップアップが閉じてからコマンドが実行される
			setTimeout(function() {
				command[item](target);
				target.focus();
			}, 0);
		},
	};
	
	return {
		get isClosed() { return !dialog || dialog.closed; },
		hide: function() {
			if (this.isClosed) return;
			dialog.close();
			dialog = null;
		},
		/** @param {PointerEvent} evt */
		show: function(evt) {
			this.hide();
			
			target = /** @type {HTMLAnchorElement} */ (evt.target);
			const folder = target.xFolder;
			
			items.forEach(function(item) {
				item.isVisible =
					item.isAlwaysVisible != null ? item.isAlwaysVisible :
					item.isConsole && !folder.isFileFolder ? false :
					item.id == "showProperty" ? folder.hasProperties() :
					item.isExtended ? evt.shiftKey : true;
			});
			
			let left = evt.screenX;
			let top = evt.screenY;
			// 右クリックではなくshift+F10等を使った場合
			if (evt.clientX == 0 && evt.clientY == 0) {
				const rect = target.getBoundingClientRect();
				left = rect.left + window.screenLeft + 8;
				top = rect.top + window.screenTop + 8;
			}
			
			const dlgopts = "dialogWidth: 0px; dialogHeight: 0px; unadorned: 1; " +
				"dialogLeft: " + left + "; dialogTop: " + top;
			
			dialog = window.showModelessDialog("modules/popup.html", dlgargs, dlgopts);
		},
	};
})();


function libMain() {
	loadScript("data", function() {
		// 起動した時にウインドウが後ろに表示されてしまうことがあるので、このメソッドで最前面に持ってくる
		window.focus();
		
		newListHtml();
		addEventHandler();
	});
}

function newListHtml() {
	const fragment = document.createDocumentFragment();
	
	if (!State.OS.isSuppoertedVersion) {
		const p = document.createElement("p");
		p.id = "warning";
		p.innerText = unsupportMessage.warning;
		fragment.appendChild(p);
	}
	if (Setting.debug) fragment.appendChild(getAppInfo());
	
	let dl = document.createElement("dl");
	
	const it = SpecialFolders.iterator();
	for (;;) {
		const result = it.next();
		if (result.done) break;
		const folder = result.value;
		
		if (folder.category) {
			fragment.appendChild(dl);
			
			if (Setting.viewCategory) {
				const h2 = document.createElement("h2");
				h2.innerText = "[" + folder.category +"]";
				fragment.appendChild(h2);
			}
			
			dl = document.createElement("dl");
		}
		
		if (!folder.folderItem) continue;
		
		const dt = document.createElement("dt");
		dt.innerText = folder.title || " ";
		
		const a = document.createElement("a");
		a.href = "#";
		a.innerText = folder.path;
		a.xFolder = folder;
		
		const dd = document.createElement("dd");
		dd.appendChild(a);
		
		dl.appendChild(dt);
		dl.appendChild(dd);
	}
	
	fragment.appendChild(dl);
	document.body.appendChild(fragment);
}

function getAppInfo() {
	const p = document.createElement("p");
	
	p.id = "info";
	p.innerText = "Version: {0}\nOS: {1}\nPlatform: {2} bit".
		xFormat(State.version, State.OS.caption, State.Host.platform);
	if (State.Host.isWow64) p.innerText += " (Wow64)";
	
	return p;
}

function addEventHandler() {
	/** @param {MouseEvent} evt */
	document.onclick = function(evt) {
		evt.preventDefault();
		
		const target = evt.target;
		if (!(target instanceof HTMLAnchorElement)) return;
		
		command[getVerb(evt)](target);
	};
	/** @param {PointerEvent} evt */
	document.oncontextmenu = function(evt) {
		evt.preventDefault();
		
		const target = evt.target;
		if (target instanceof HTMLAnchorElement) {
			if (evt.altKey) command.copyAsPath(target);
			else popup.show(evt);
		} else if (Setting.debug && evt.altKey) {
			htaDebug.breakpoint();
		}
		
	};
	/** @param {KeyboardEvent} evt */
	document.onkeyup =  function(evt) {
		evt.preventDefault();
		
		const target = evt.target;
		if (!(target instanceof HTMLAnchorElement)) return;
		if (evt.keyCode == VK_RETURN && evt.altKey) command.showProperty(target);
	};
	/** @param {DragEvent} evt */
	document.ondragstart = function(evt) { evt.preventDefault(); };
	
	window.onfocus = function() { popup.hide(); };
	/** @param {WheelEvent} evt */
	window.onmousewheel = function(evt) { if (!popup.isClosed) evt.preventDefault(); };
}

/**
 * @param {MouseEvent} evt
 * @returns {string}
 */
function getVerb(evt) {
	return evt.altKey ? "showProperty" :
		evt.ctrlKey ? "execPowershell" :
		evt.shiftKey ? "execCmd" : "openFolder";
}


libMain();
