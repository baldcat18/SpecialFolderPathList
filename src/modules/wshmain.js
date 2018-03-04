/// <reference path="common.js" />

var stdout = WScript.StdOut;
var stderr = WScript.StdErr;

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
		o: "directoryOnly",
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
	
	if (Setting.wshWriteBom) stdout.Write("\uFEFF");
	if (Setting.debug) writeAppInfo();
	writeList();
}

function writeAppInfo() {
	stdout.WriteLine("Version: " + State.version);
	stdout.WriteLine("OS: " + State.OS.caption);
	stdout.Write("Platform: " + State.Host.platform + " bit");
	if (State.Host.isWow64) stdout.Write(" (Wow64)");
	stdout.WriteLine("\n");
}

function writeList() {
	for (;;) {
		var folder = SpecialFolder.getObject();
		if (!folder) break;
		
		try {
			if (folder.category) {
				stdout.WriteLine("");
				if (Setting.viewCategory) stdout.WriteLine("Category: {0}\n".xFormat(folder.category));
			}
			
			if (!folder.folderItem && !Setting.wshForceWriteAllData) continue;
			
			stdout.WriteLine("Title: " + folder.title);
			if (Setting.wshWriteDir) stdout.WriteLine("Dir: " + (folder.dir || ""));
			if (folder.folderItem) {
				stdout.WriteLine("Path: " + folder.path);
				if (Setting.wshWriteDisplayName) stdout.WriteLine("DisplayName: " + folder.folderItem.Name);
			}
			if (Setting.wshWriteType) stdout.WriteLine("Type: " + folder.getType());
			stdout.WriteLine("");
		} catch (err) {
			if (toUint32(/** @type {Error} */ (err).number) == 0x800700E8) { // パイプを閉じています
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
