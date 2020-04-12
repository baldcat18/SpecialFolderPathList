/// <reference path="common.js" />

/**
 * @param {string} msg
 * @param {string} url
 * @param {number} line
*/
window.onerror = function(msg, url, line) {
	writeError("{0}\n\nurl: {1}\nline: {2}".xFormat(msg, url, line));
	return true;
};

if (getDocumentMode() < 9) {
	/** @type {(arg: any) => boolean} */
	// @ts-ignore
	Array.isArray = function(arg) {
		return Object.prototype.toString.call(arg) == "[object Array]";
	};
	/** @type {(callbackfn: (value: any, index: number, array: any[]) => void, thisArg?: any) => void} */
	Array.prototype.forEach = function(callbackfn, thisArg) {
		for (var i = 0; i < this.length; i++) {
			if (i in this) callbackfn.call(thisArg || this, this[i], this);
		}
	};
	Object.getOwnPropertyNames = function(o) {
		/** @type {string[]} */
		var names = [];

		for (var name in o) {
			if (o.hasOwnProperty(name)) names.push(name);
		}

		return names;
	};
}

G.ssfPROFILE = 40;
G.TemporaryFolder = 2;

var htaDebug = (function() {
	/** 0x800A01B6: オブジェクトでサポートされていないプロパティまたはメソッドです */
	var E_NO_PROPERTY = -2146827850;

	/** @type {Window} */
	var dbgbox = null;
	/** @type {HTMLDivElement} */
	var logBox = null;

	/** @type {(value: any) => string} */
	function getType(value) {
		if (value === null) return "null";
		var type = typeof value;
		if (type != "object") return type;
		var name = /** @type {string} */ (Object.prototype.toString.call(value)).replace(/^\[object (.+)\]$/, "$1");
		return (name != "Object" || value.construcor == Object) ? name : type;
	}

	/** @type {(value: any) => string} */
	function getString(value) {
		try {
			return String(value);
		} catch (err) {
			if (/** @type {Error} */ (err).number != E_NO_PROPERTY) throw err;
			return "unprintable object";
		}
	}

	var retobj = {
		/**
		 * @type {(expression: string, callEval: (expr: string) => any, message?: string) => void}
		 * @example htaDebug.assert("str === 'foo'", function(x){return eval(x);});
		 */
		assert: function(expression, callEval, message) {
			if (!message) {
				message = "Assertion failed: " + expression;
				if (getDocumentMode() >= 10) {
					try {
						throw new Error(message);
					} catch (err) {
						message = /** @type {Error} */ (err).stack;
					}
				}
			}

			if (!callEval(expression)) {
				writeError(message);
				window.close();
			}
		},
		/** @type {(message: string, title?: string) => void} */
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
		/** @type {(obj: {}, title?: string) => void} */
		list: function(obj, title) {
			var names = Object.getOwnPropertyNames(obj);
			var list = "\n";

			for (var i = 0; i < names.length; i++) {
				var name = names[i];
				var value = obj[name];
				var type = getType(value);
				if (value == null || type == "function") value = "";

				list += "{0}: [{1}] {2}\n".xFormat(name, type, getString(value));
			}

			this.write(list, title);
		},
		/** @type {(obj: {}, depth = 4) => void} */
		dir: function(obj, depth) {
			if (depth === undefined) depth = 4;

			this.write("\n" + createDir(obj, 0));

			/** @type {(value: any, currentDepth: number) => string} */
			function createDir(value, currentDepth) {
				var type = getType(value);
				if (type == "function") return "function";
				if (type == "null" || typeof value != "object") return String(value);
				if (currentDepth >= depth) return Array.isArray(value) ? "[{0}]".xFormat(value) : getString(value);

				var tabs = "";
				for (var i = 0; i < currentDepth; i++) tabs += "    ";
				var tmp = "";
				if (Array.isArray(value)) {
					value.forEach(function(val) {
						tmp += "{0}    {1}\n".xFormat(tabs, createDir(val, currentDepth + 1));
					}, "");
					return "[\n{0}{1}]".xFormat(tmp, tabs);
				}
				Object.getOwnPropertyNames(value).forEach(function(name) {
					tmp += "{0}    {1}: {2}\n".xFormat(tabs, name, createDir(value[name], currentDepth + 1));
				}, "");
				return "{\n{0}{1}}".xFormat(tmp, tabs);
			}
		},
		/** @type {(callEval = eval, title = "") => void} */
		breakpoint: function(callEval, title) {
			if (callEval === undefined) callEval = eval;
			if (title === undefined) title = "";

			var expr = "";
			var result = "";

			for (; ;) {
				expr = prompt(title, expr);
				if (!expr) break;
				try {
					result = getString((callEval)(expr));
				} catch (err) {
					result = err instanceof Error ?
						"{0} (0x{1})\n{2}".xFormat(err.name, (err.number || 0).xToHex(), err.message) : err;
				}
				this.write(result, expr);
			}
		}
	};

	if (!Setting.debug) {
		var nop = function() { };
		for (var method in retobj) retobj[method] = nop;
	}

	return retobj;
})();

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
 * @param {() => void} [onload]
 */
function loadScript(name, onload) {
	var script = document.createElement("script");
	script.charset = "UTF-8";
	script.src = "modules/{0}.js".xFormat(name);
	if (onload) script.onload = onload;
	document.head.appendChild(script);
}


main();
