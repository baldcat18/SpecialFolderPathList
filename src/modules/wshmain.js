/// <reference path="common.js" />

var STDOUT = WScript.StdOut;

function updateSetting() {
	/** @type {WshNamed} */
	// @ts-ignore
	var wshNamed = WScript.Arguments.Named;
	var properties = {
		bom: "wshWriteBom",
		dbg: "debug",
		dir: "wshWriteDir",
		c: "viewCategory",
		d: "wshWriteDisplayName",
		f: "wshForceWriteAllData",
		o: "fileFolderOnly",
		t: "wshWriteType"
	};

	for (var opt in properties) {
		if (wshNamed.Exists(opt)) Setting[properties[opt]] = wshNamed.Item(opt) !== false;
	}

	if (!Setting.debug) Setting.wshWriteDir = false;
}

function main() {
	if (State.Host.type != "cscript") {
		writeError("このスクリプトは cscript.exe で実行してください。");
		WScript.Quit(1);
	}

	updateSetting();

	if (!State.OS.isSuppoertedVersion) {
		if (Setting.abortIfOldOS) {
			writeError(unsupportMessage.error);
			WScript.Quit(1);
		} else {
			writeError(unsupportMessage.warning);
		}
	}

	if (Setting.wshWriteBom) STDOUT.Write("\uFEFF");
	if (Setting.debug) writeAppInfo();
	writeList();
}

function writeAppInfo() {
	STDOUT.WriteLine("Version: " + State.version);
	STDOUT.WriteLine("OS: " + State.OS.caption);
	STDOUT.Write("Platform: " + State.Host.platform + " bit");
	if (State.Host.isWow64) STDOUT.Write(" (Wow64)");
	STDOUT.WriteLine("\n");
}

function writeList() {
	var it = SpecialFolders.iterator();
	for (; ;) {
		var result = it.next();
		if (result.done) break;
		var folder = result.value;

		try {
			if (folder.category) {
				STDOUT.WriteLine("");
				if (Setting.viewCategory) STDOUT.WriteLine("Category: {0}\n".xFormat(folder.category));
			}

			if (!folder.folderItem && !Setting.wshForceWriteAllData) continue;

			STDOUT.WriteLine("Title: " + folder.title);
			if (Setting.wshWriteDir) STDOUT.WriteLine("Dir: " + (folder.dir || ""));
			if (folder.folderItem) {
				STDOUT.WriteLine("Path: " + folder.path);
				if (Setting.wshWriteDisplayName) STDOUT.WriteLine("DisplayName: " + folder.folderItem.Name);
			}
			if (Setting.wshWriteType) STDOUT.WriteLine("Type: " + folder.getType());
			STDOUT.WriteLine("");
		} catch (err) {
			if (/** @type {Error} */ (err).number == E_NODATA) {
				writeError(newErrorMessage(err, "パイプを閉じています。"));
				WScript.Quit(1);
			} else {
				var errmsg = newErrorMessage(err, "[{0}] のパス情報を出力できませんでした。".xFormat(folder.title));
				writeError(errmsg);
			}
		}
	}
}

/**
 * @param {any} err
 * @param {string} description
 * @returns {string}
 */
function newErrorMessage(err, description) {
	return err instanceof Error ?
		"{0}[0x{1}] {2}\n{3}".xFormat(err.name, err.number.xToHex(), err.message, description) :
		"{0}\n{1}".xFormat(String(err), description);
}


main();
