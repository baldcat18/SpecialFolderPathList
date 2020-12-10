﻿/// <reference path="decl-userdef.d.ts" />

var G = this;

if (!Object.create) {
	/**
	 * @param {object} o
	 * @returns {any}
	 */
	Object.create = function(o) {
		var _ = function() { };
		_.prototype = o;
		return new _();
	};
}

/** 0x80070002 */
var E_NOTFOUND = -2147024894;
/** 0x800700E8: パイプを閉じています */
var E_NODATA = -2147024664;

// https://docs.microsoft.com/en-us/windows/desktop/api/winuser/nf-winuser-messagebox

/** 感嘆符（!）アイコンを表示します。 */
var MB_ICONWARNING = 0x30;

// https://docs.microsoft.com/ja-jp/windows/desktop/api/sysinfoapi/ns-sysinfoapi-_system_info

/** x86 */
var PROCESSOR_ARCHITECTURE_INTEL = 0;
/** ARM */
var PROCESSOR_ARCHITECTURE_ARM = 5;
/** Intel Itanium-based */
var PROCESSOR_ARCHITECTURE_IA64 = 6;
/** x64 (AMD or Intel) */
var PROCESSOR_ARCHITECTURE_AMD64 = 9;
/** ARM64 */
var PROCESSOR_ARCHITECTURE_ARM64 = 12;
/** Unknown architecture. */
var PROCESSOR_ARCHITECTURE_UNKNOWN = 0xFFFF;

var wShell = new ActiveXObject("WScript.Shell");
var shell = new ActiveXObject("Shell.Application");
var fso = new ActiveXObject("Scripting.FileSystemObject");

Number.prototype.xToHex = function() {
	return toUint32(this).toString(16).toUpperCase();
};

String.prototype.xExpand = function() {
	return wShell.ExpandEnvironmentStrings("" + this);
};
String.prototype.xFormat = function() {
	var args = arguments;

	/**
	 * @param {string} matched
	 * @param {string} index
	 * @param {string} fmt
	 * @returns {string}
	 */
	function replacer(matched, index, fmt) {
		if (matched == "{{") return "{";
		if (matched == "}}") return "}";

		if (!(index in args)) throw new RangeError("インデックスが不正: " + index);
		var value = args[index];
		return fmt ? value.toString(fmt) : value;
	}

	return this.replace(/\{\{|\}\}|\{(\d+)(?::(.+?))?\}/g, replacer);
};

G.Setting = { debug: false };

G.Version = (function() {
	var unspecified = -1;
	var errmsg = "引数 {0} がマイナス: {1}";

	/**
	 * @constructor
	 * @param {number} major
	 * @param {number} minor
	 * @param {number} [build]
	 * @param {number} [revision]
	 */
	function Version(major, minor, build, revision) {
		if (major < 0) throw new RangeError(errmsg.xFormat("major", major));
		this.major = toInt32(major);

		if (minor < 0) throw new RangeError(errmsg.xFormat("minor", minor));
		this.minor = toInt32(minor);

		this.build = unspecified;
		this.revision = unspecified;

		if (build == null) return;
		if (build < 0) throw new RangeError(errmsg.xFormat("build", build));
		this.build = toInt32(build);

		if (revision == null) return;
		if (revision < 0) throw new RangeError(errmsg.xFormat("revision", revision));
		this.revision = toInt32(revision);
	}

	/**
	 * @param {Version} value
	 * @returns {number}
	 */
	Version.prototype.compareTo = function(value) {
		if (!value) return 1;

		if (this.major != value.major) return this.major > value.major ? 1 : -1;
		if (this.minor != value.minor) return this.minor > value.minor ? 1 : -1;
		if (this.build != value.build) return this.build > value.build ? 1 : -1;
		if (this.revision != value.revision) return this.revision > value.revision ? 1 : -1;

		return 0;
	};
	/**
	 * @param {Version} obj
	 * @returns {boolean}
	 */
	Version.prototype.equals = function(obj) {
		return this.compareTo(obj) == 0;
	};
	/**
	 * @param {Version} obj
	 * @returns {boolean}
	 */
	Version.prototype.isGreaterThan = function(obj) {
		return !!obj && this.compareTo(obj) > 0;
	};
	/**
	 * @param {number} [fieldCount]
	 * @returns {string}
	 */
	Version.prototype.toString = function(fieldCount) {
		if (fieldCount == undefined) {
			fieldCount = this.build == unspecified ? 2 : this.revision == unspecified ? 3 : 4;
		}

		/** @type {number[]} */
		var tmp = [];
		if (fieldCount >= 1) tmp.push(this.major);
		if (fieldCount >= 2) tmp.push(this.minor);
		if (fieldCount >= 3 && this.build != unspecified) tmp.push(this.build);
		if (fieldCount >= 4 && this.revision != unspecified) tmp.push(this.revision);
		return tmp.join(".");
	};

	return Version;
})();

var State = (function() {
	var appVersion = "1.3.7 alpha";

	/**
	 * @template T
	 * @param {string} name
	 * @param {T} [defaultValue]
	 * @returns {T}
	 */
	function getWinNTCurrentVersionValue(name, defaultValue) {
		return getRegValue("HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\" + name, defaultValue);
	}

	var major = getWinNTCurrentVersionValue("CurrentMajorVersionNumber", -1);
	var minor = getWinNTCurrentVersionValue("CurrentMinorVersionNumber", -1);
	if (major < 0 || minor < 0) {
		var currentVersion = getWinNTCurrentVersionValue("CurrentVersion", "0.0").split(".");
		major = +currentVersion[0];
		minor = +currentVersion[1];
	}

	var buildLab = getWinNTCurrentVersionValue("BuildLabEx", "") || getWinNTCurrentVersionValue("BuildLab", "");
	var build = +getWinNTCurrentVersionValue("CurrentBuild") || parseInt(buildLab, 10) || null;
	var revision = getWinNTCurrentVersionValue("UBR", -1);
	if (revision < 0) revision = +buildLab.split(".", 2)[1] || null;

	var osVersion = new Version(major, minor, build, revision);
	var osVersionString = osVersion.toString(3);

	var dispVersion =
		getWinNTCurrentVersionValue("DisplayVersion", "") || // Win10 20H2から
		getWinNTCurrentVersionValue("ReleaseId", "") || // Win10 1511から
		getWinNTCurrentVersionValue("CSDVersion", "");

	var OS = {
		version: osVersion,
		caption: "{0}{1} ({2})\n    {3}".xFormat(
			getWinNTCurrentVersionValue("ProductName"), dispVersion ? " " + dispVersion : "", osVersion, buildLab),
		isSuppoertedVersion:
			osVersion.isGreaterThan(new Version(10, 0, 18363)) || // Win10 1909以降
			osVersionString == "10.0.17763" || // Win10 1809 Enterprise
			osVersionString == "10.0.17134" || // Win10 1803 Enterprise
			osVersionString == "6.3.9600" && osVersion.revision >= 17031 || // Win8.1 Update
			osVersion.toString(2) == "6.1" && osVersion.build >= 7601 // Win7 ESU
	};

	/** @returns {string} */
	function getHostType() {
		if (G.window && /\.hta$/.test(location.href)) return "mshta";
		if (G.WScript) return fso.GetBaseName(WScript.FullName).toLowerCase();
		throw new Error("対応していない実行環境です。");
	}

	function getPlatform() {
		switch (shell.GetSystemInformation("ProcessorArchitecture")) {
			// case PROCESSOR_ARCHITECTURE_IA64:
			case PROCESSOR_ARCHITECTURE_AMD64:
			case PROCESSOR_ARCHITECTURE_ARM64:
				return 64;
			case PROCESSOR_ARCHITECTURE_INTEL:
			case PROCESSOR_ARCHITECTURE_ARM:
				return 32;
			default:
				throw new Error("対応していないプラットフォームです。");
		}
	}

	var Host = {
		type: getHostType(),
		platform: getPlatform(),
		isWow64: !!wShell.Environment("Process").Item("PROCESSOR_ARCHITEW6432")
	};

	return { version: appVersion, OS: OS, Host: Host };
})();

var unsupportMessage = (function() {
	var errmsg = "このバージョンの Windows には対応していません。";

	return {
		error: errmsg,
		warning: errmsg + "\n誤った情報が出力されたり、正しく動作しない可能性があります。"
	};
})();

/**
 * @param {any} value
 * @returns {number}
 */
function toInt32(value) {
	// value | 0 や value >> 0 でもいい
	return ~~value;
}

/**
 * @param {any} value
 * @returns {number}
 */
function toUint32(value) {
	return value >>> 0;
}

/**
 * @template T
 * @param {string} name
 * @param {T} [defaultValue=null]
 * @param {boolean} [expand=false]
 * @returns {T}
 */
// @ts-ignore: オーバーロードあり
function getRegValue(name, defaultValue, expand) {
	if (defaultValue === undefined) defaultValue = null;
	if (expand === undefined) expand = false;

	var returnValue;
	try {
		returnValue = wShell.RegRead(name);
	} catch (err) {
		// E_NOTFOUND: レジストリのキーや値が見つからない
		if (/** @type {Error} */ (err).number != E_NOTFOUND) throw err;
		returnValue = defaultValue;
	}
	return expand && typeof returnValue == "string" ? returnValue.xExpand() : returnValue;
}

/** @returns {string} */
function getSystemPath() {
	return fso.GetSpecialFolder(/** @type {SpecialFolderConst.SystemFolder} */(1)).Path;
}

/** @returns {string} */
function getSysNativePath() {
	return "%windir%\\SysNative".xExpand();
}

/** @type {(text: string) => void} */
var writeError =
	State.Host.type == "mshta" ? function(text) {
		// ウインドウが後ろに表示されてしまうことがあるので、このメソッドで最前面に持ってくる
		if (!document.hasFocus()) window.focus();

		// WshShell#Popup()はモードレスダイアログで
		// ダイアログを残したままメインウインドウを閉じるとプロセスがリークしてしまうので
		// モーダルダイアログであるalert()を使う
		alert(text);
	} :
		State.Host.type == "cscript" ? function(text) { WScript.StdErr.WriteLine(text); } :
			function(text) { wShell.Popup(text, 0, "SpecialFolderPathList", MB_ICONWARNING); };
