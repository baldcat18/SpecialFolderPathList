/// <reference path="common.js" />

if (!Object.getOwnPropertyNames) {
	Object.getOwnPropertyNames = function(o) {
		/** @type {string[]} */
		var names = [];
		
		for (var name in o) {
			if (o.hasOwnProperty(name)) names.push(name);
		}
		
		return names;
	};
}

var htaDebug = (function() {
	/** @type {Window} */
	var dbgbox = null;
	/** @type {HTMLDivElement} */
	var logBox = null;
	
	/**
	 * @param {any} value 
	 * @returns {string}
	 */
	function getType(value) {
		if (value === null) return "null";
		var type = typeof value;
		if (type != "object") return type;
		/** @type {string} */
		var name = Object.prototype.toString.call(value).replace(/^\[object (.+)\]$/,"$1");
		return (name != "Object" || value.construcor == Object) ? name : type;
	}
	
	var retobj = {
		/**
		 * @param {string} message 
		 * @param {string} [title]
		 */
		write: function(message, title) {
			if (!dbgbox || dbgbox.closed) {
				dbgbox =
					window.showModelessDialog("modules/dbgbox.html", null, "dialogWidth: 700px; dialogHeight: 300px");
				logBox = dbgbox.document.getElementsByTagName("div")[0];
			}
			
			var text = "< " + message;
			if (title) text = "> " + title + "\n" + text;
			
			var p = dbgbox.document.createElement("p");
			p.innerText = text;
			logBox.appendChild(p);
			
			dbgbox.location.hash = "#bottom";
		},
		/**
		 * @param {{}} obj 
		 * @param {string} [title]
		 */
		dir: function(obj, title) {
			var names = Object.getOwnPropertyNames(obj);
			var dir = "\n";
			
			for (var i = 0; i < names.length; i++) {
				var name = names[i];
				var value = obj[name];
				var type = getType(value);
				if (value == null || type == "function") value = "";
				
				dir += "{0}: [{1}] {2}\n".xFormat(name, type, value);
			}
			
			this.write(dir, title);
		},
		/**
		 * @param {function(string): any} [callEval] 
		 * @param {string} [title]
		 */
		breakpoint: function(callEval, title) {
			var expr = "";
			var result = "";
			
			for (;;) {
				expr = prompt(title || "", expr);
				if (!expr) break;
				try {
					result = String((callEval || eval)(expr));
				} catch (err) {
					result = err instanceof Error ?
						"{0} (0x{1})\n{2}".xFormat(err.name, (err.number || 0).xToHex(), err.message) : err;
				}
				this.write(result, expr);
			}
		}
	};
	
	if (!Setting.debug) {
		var nop = function() {};
		for (var method in retobj) retobj[method] = nop;
	}
	
	return retobj;
})();

/**
 * @param {string} msg
 * @param {string} url
 * @param {number} line
*/
window.onerror = function(msg, url, line) {
	writeError("{0}\n\nurl: {1}\nline: {2}".xFormat(msg, url, line));
	return true;
};

function main() {
	if (Setting.abortIfOldOS && !State.OS.isSuppoertedVersion) {
		writeError(unsupportMessage.error);
		window.close();
		return;
	}
	
	if (getDocumentMode() < 11) {
		writeError("このツールを実行するには、Internet Explorer を\n11 にバージョンアップする必要があります。");
		window.close();
		return;
	}
	
	if (!Setting.htaFoeceWow64 && State.Host.isWow64) {
		// 32ビット版のウインドウが表示されないようにするため画面外に飛ばしてサイズも小さくする
		window.moveTo(-30000, -30000);
		window.resizeTo(1, 1);
		
		wShell.Run("\"{0}\\mshta.exe\" {1}".xFormat(getSysNativePath(), location.href));
		window.close();
		return;
	}
	
	var width = Math.floor(Setting.htaWidth);
	var height = Math.floor(Setting.htaHeight);
	if (width > 0 && height > 0) window.resizeTo(width, height);
	
	var left = Math.floor(Setting.htaLeft);
	var top = Math.floor(Setting.htaTop);
	if (!isNaN(left) && !isNaN(top)) window.moveTo(left, top);
	
	window.onload = function() { loadScript("htalib"); };
}

function getDocumentMode() {
	if (document.documentMode) return document.documentMode;
	
	var ieKey = "HKLM\\SOFTWARE\\Microsoft\\Internet Explorer\\";
	return parseFloat(getRegValue(ieKey + "svcVersion") || getRegValue(ieKey + "Version"));
}

/**
 * @param {string} name
 * @param {function(): void} [onload]
 */
function loadScript(name, onload) {
	var script = document.createElement("script");
	script.charset = "UTF-8";
	script.src = "modules/{0}.js".xFormat(name);
	if (onload) script.onload = onload;
	document.head.appendChild(script);
}


main();
