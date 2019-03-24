/// <reference path="common.js" />

"use strict";

(function() {
	var OS = State.OS;
	
	/** Win8以降 */
	var WIN8  = OS.version.isGreaterThan(new Version(6, 2));
	/** Win8.1以降 */
	var WIN81 = OS.version.isGreaterThan(new Version(6, 3));
	/** Win10以降 */
	var WIN10 = OS.version.isGreaterThan(new Version(10, 0));
	/** Win10 1607以降 */
	var WIN10_1607 = OS.version.isGreaterThan(new Version(10, 0, 14393));
	/** Win10 1703以降 */
	var WIN10_1703 = OS.version.isGreaterThan(new Version(10, 0, 15063));
	/** Win10 1709以降 */
	var WIN10_1709 = OS.version.isGreaterThan(new Version(10, 0, 16299));
	/** Win10 1803以降 */
	var WIN10_1803 = OS.version.isGreaterThan(new Version(10, 0, 17134));
	/** WWin10IP Build18343以降 */
	var WIN10_1903 = OS.version.isGreaterThan(new Version(10, 0, 18343));
	
	var WIN10_1507_to_1511 = WIN10 && !WIN10_1607;
	
	var IS64BIT = State.Host.platform == 64;
	
	var USER_SHELL_FOLDERS_KEY =
		"HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders\\";
	var CURRENT_VERSION_KEY = "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\";
	
	var LIBRARIES_PATH = getRegValue(USER_SHELL_FOLDERS_KEY + "{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}",
			"%APPDATA%\\Microsoft\\Windows\\Libraries", true);
	var USER_ACCOUNT_PICTURES_PATH = "shell:Common AppData\\Microsoft\\User Account Pictures";
	
	var WINDOWS_APPS_TITLE = WIN10 ? "Windows アプリ" : "ストアアプリ";
	
	/** @type {PropertyTypes} */
	var PT_DEFAULT = 0;
	/**
	 * ShellExecuteメソッドを使う
	 * @type {PropertyTypes}
	 */
	var PT_SHELLEXECUTE = 1;
	/**
	 * コンテキストメニューの[プロパティ]を使う
	 * @type {PropertyTypes}
	 */
	var PT_VERB = 2;
	
	/** @type {SpecialFolderOption} */
	var DEFAULT_OPTION = {};
	
	if (Setting.directoryOnly !== undefined && Setting.fileFolderOnly === undefined) {
		Setting.fileFolderOnly = Setting.directoryOnly;
	}
	
	/**
	 * @param {string} path
	 * @returns {FolderItem}
	 */
	function getDirectoryFolderItem(path) {
		return shell.NameSpace(fso.GetParentFolderName(path)).Items().Item(fso.GetFileName(path));
	}
	
	
	/**
	 * @class
	 * @abstract
	 * @this {SpecialFolder}
	 * @param {SpecialFolderArgument} arg
	 */
	function SpecialFolder(arg) {
		this.title = arg.title;
		this.dir = arg.dir;
		this.folderItem = arg.folderItem;
		this.path = arg.path;
		this.category = arg.option.category || "";
		/** @private */
		this._folderItemForProperties = arg.option.folderItemForProperties;
	}
	/**
	 * @private
	 * @type {PropertyTypes}
	 */
	SpecialFolder.prototype._propertyTypes = null;
	/**
	 * @private
	 * @type {FolderItemVerb}
	 */
	SpecialFolder.prototype._properties = undefined;
	SpecialFolder.prototype.open = function() { this.folderItem.InvokeVerb(); };
	/** @param {string} [verb] */
	SpecialFolder.prototype.execExplorer = function(verb) {
		var path = typeof this.dir == "number" ? this.path : this.dir;
		shell.ShellExecute("explorer.exe", "\"{0}\"".xFormat(path), null, verb);
	};
	SpecialFolder.prototype.hasProperties = function() {
		if (this._propertyTypes == PT_SHELLEXECUTE) return true;
		if (this._folderItemForProperties === null) return false;
		
		if (this._properties === undefined) {
			this._properties = null;
			var verbs = (this._folderItemForProperties || this.folderItem).Verbs();
			if (verbs) {
				for (var i = verbs.Count - 1; i >= 0; i--) {
					var verb = verbs.Item(i);
					if (verb.Name == "プロパティ(&R)") {
						this._properties = verb;
						break;
					}
				}
			}
		}
		return !!this._properties;
	};
	SpecialFolder.prototype.showProperties = function() {
		if (this._propertyTypes == PT_SHELLEXECUTE) shell.ShellExecute(this.dir, "", "", "properties");
		else if (this.hasProperties()) this._properties.DoIt();
		else writeError("プロパティを表示できません。");
	};
	SpecialFolder.prototype.getType = function() { return this.folderItem.Type; };
	
	/**
	 * @class
	 * @this {SpecialFolder}
	 * @param {SpecialFolderArgument} arg
	 */
	function FileFolder(arg) {
		SpecialFolder.call(this, arg);
		this._propertyTypes = arg.option.propertyType ||
			(this._folderItemForProperties === undefined ? PT_SHELLEXECUTE : PT_VERB);
	}
	FileFolder.prototype = Object.create(SpecialFolder.prototype);
	FileFolder.prototype.constructor = FileFolder;
	FileFolder.prototype.isFileFolder = true;
	/** @param {string} [verb] */
	FileFolder.prototype.execCmd = function(verb) {
		shell.ShellExecute("cmd.exe", "/k pushd \"{0}\"".xFormat(this.path), null, verb);
	};
	/** @param {string} [verb] */
	FileFolder.prototype.execPowershell = function(verb) {
		var arg = "-NoExit -Command \"Push-Location -LiteralPath '{0}'\"".xFormat(this.path);
		shell.ShellExecute("powershell.exe", arg, null, verb);
	};
	/** @param {string} [verb] */
	FileFolder.prototype.execWsl = function(verb) {
		shell.ShellExecute("cmd.exe", "/c pushd \"{0}\" & wsl.exe".xFormat(this.path), null, verb);
	};
	
	// メソッドでthis.pathやthis._folderItemForPropertiesが未定義だと言われないようにするためのもの
	// 普通に実行するだけなら不要
	FileFolder.prototype.path = "";
	/** @type {FolderItem} */
	FileFolder.prototype._folderItemForProperties = undefined;
	
	/**
	 * @class
	 * @this {SpecialFolder}
	 * @param {SpecialFolderArgument} arg
	 */
	function VirtualFolder(arg) {
		SpecialFolder.call(this, arg);
		this._propertyTypes = arg.option.propertyType || PT_VERB;
	}
	VirtualFolder.prototype = Object.create(SpecialFolder.prototype);
	VirtualFolder.prototype.constructor = VirtualFolder;
	VirtualFolder.prototype.isFileFolder = false;
	VirtualFolder.prototype.execCmd = function(verb) {
		writeError("ファイル フォルダーではないのでコマンドプロンプトを実行できません。");
	};
	VirtualFolder.prototype.execPowershell = function() {
		writeError("ファイル フォルダーではないので PowerShell を実行できません。");
	};
	VirtualFolder.prototype.execWsl = function() {
		writeError("ファイル フォルダーではないので WSL を実行できません。");
	};
	
	/**
	 * @class
	 * @this {SpecialFolder}
	 * @param {SpecialFolderArgument} arg
	 */
	function InvalidFolder(arg) {
		SpecialFolder.call(this, arg);
	}
	InvalidFolder.prototype = Object.create(SpecialFolder.prototype);
	InvalidFolder.prototype.constructor = InvalidFolder;
	InvalidFolder.prototype.isFileFolder = false;
	/** @override */
	InvalidFolder.prototype.getType = function() { return "使用不可"; };
	
	/** @type {SpecialFolderArgument} */
	var sfArg = {
		title: "",
		dir: "",
		folderItem: null,
		path: "",
		option: null
	};
	
	/**
	 * @param {string} title
	 * @param {string} dir
	 * @param {SpecialFolderOption} [option]
	 * @returns {SpecialFolder}
	 */
	function createSpecialFolder(title, dir, option) {
		sfArg.title = title;
		if (dir && dir.slice(0, 6) != "shell:") dir = "file:" + dir;
		sfArg.dir = dir;
		sfArg.option = option || DEFAULT_OPTION;
		
		try { sfArg.folderItem = shell.NameSpace(dir).Self }
		catch (err) { sfArg.folderItem = null; }
		
		sfArg.path = "";
		if (sfArg.folderItem) {
			sfArg.path = sfArg.option.path || sfArg.folderItem.Path;
			if (sfArg.path.charAt(0) == ":") sfArg.path = "shell:" + sfArg.path;
		}
		
		// @ts-ignore: FileFolder、VirtualFolder、InvalidFolder は SpecialFolder のサブクラス
		if (fso.FolderExists(sfArg.path)) return new FileFolder(sfArg);
		if (Setting.fileFolderOnly) {
			sfArg.folderItem = null;
			sfArg.path = "";
		}
		// @ts-ignore
		return new (sfArg.folderItem ? VirtualFolder : InvalidFolder)(sfArg);
	}
	
	var DONE_ITERATION = Infinity;
	
	global.SpecialFolders = {
		item: (function() {
			/** @type {FolderIteratorIndex} */
			var index = null;
			return item;
			
			/**
			 * @param {number} itemIndex
			 * @returns {SpecialFolder}
			 */
			function item(itemIndex) {
				if (!index) index = { current: 0 };
				
				index.current = itemIndex;
				return getSpecialFolder(index);
			};
		})(),
		iterator: function() {
			return {
				/** @returns {IteratorResult<SpecialFolder>} */
				next: function() {
					return ++this._index.current == DONE_ITERATION ?
						{ done: true, value: undefined } : { done: false, value: getSpecialFolder(this._index) };
				},
				/**
				 * @private
				 * @type {FolderIteratorIndex}
				 */
				_index: { current: -1 }
			}
		}
	};
	
	/**
	 * @param {FolderIteratorIndex} index
	 */
	function getSpecialFolder(index) {
		switch (index.current) {
		case 0:
			// shell:Profile
			// shell:::{59031A47-3F72-44A7-89C5-5595FE6B30EE}
			// shell:ThisDeviceFolder / shell:::{F8278C54-A712-415B-B593-B77A2BE0DDA9} ([このデバイス]) (Win10 1703から)
			// %USERPROFILE%
			// %HOMEDRIVE%%HOMEPATH%
			return createSpecialFolder("個人用フォルダー", "shell:UsersFilesFolder", { category: "UserProfile", folderItemForProperties: shell.NameSpace(ssfPROFILE).Self });
		case 1:
			// Win10からサポート
			// shell:UsersFilesFolder\3D Objects
			// shell:MyComputerFolder\::{0DB7E03F-FC29-4DC6-9020-FF41B59E513A} (Win10 1709から)
			// Win10 1507から1703では3D Builderを起動した時に自動生成される
			return createSpecialFolder("3D オブジェクト", "shell:3D Objects");
		case 2:
			// shell:MyComputerFolder\::{B4BFCC3A-DB2C-424C-B029-7FE99A87C641} (Win8.1から)
			return createSpecialFolder("デスクトップ ディレクトリ", WIN81 ? "shell:ThisPCDesktopFolder" : wShell.SpecialFolders.Item("Desktop"));
		case 3:
			// shell:Local Documents / shell:MyComputerFolder\::{D3162B92-9365-467A-956B-92703ACA08AF} (Win10から)
			// shell:::{450D8FBA-AD25-11D0-98A8-0800361B1103} ([マイ ドキュメント] (Win8.1から))
			// shell:MyComputerFolder\::{A8CDFF1C-4878-43BE-B5FD-F8091C1C60D0} (Win8.1から)
			return createSpecialFolder(WIN81 ? "ドキュメント" : "マイ ドキュメント", "shell:Personal");
		case 4:
			// shell:Local Downloads / shell:MyComputerFolder\::{088E3905-0323-4B02-9826-5D99428E115F} (Win10から)
			// shell:MyComputerFolder\::{374DE290-123F-4565-9164-39C4925E467B} (Win8.1から)
			return createSpecialFolder("ダウンロード", "shell:Downloads");
		
		case 5:
			// shell:Local Music / shell:MyComputerFolder\::{3DFDF296-DBEC-4FB4-81D1-6A3438BCF4DE} (Win10から)
			// shell:MyComputerFolder\::{1CF1260C-4DD0-4EBB-811F-33C572699FDE} (Win8.1から)
			return createSpecialFolder(WIN81 ? "ミュージック" : "マイ ミュージック", "shell:My Music");
		case 6:
			// shell:My Music\Playlists
			// WMPやGroove ミュージックで再生リストを作成する時に自動生成される
			return createSpecialFolder("プレイリスト", "shell:Playlists");
		
		case 7:
			// shell:Local Pictures / shell:MyComputerFolder\::{24AD3AD4-A569-4530-98E1-AB02F9417AA8} (Win10から)
			// shell:MyComputerFolder\::{3ADD1653-EB32-4CB0-BBD7-DFA0ABB5ACCA} (Win8.1から)
			return createSpecialFolder(WIN81 ? "ピクチャ" : "マイ ピクチャ", "shell:My Pictures");
		case 8:
			// Win8.1からサポート
			// shell:My Pictures\Camera Roll
			// カメラアプリで写真や動画を撮影する時に自動生成される
			return createSpecialFolder("カメラ ロール", "shell:Camera Roll");
		case 9:
			// Win10からサポート
			// shell:My Pictures\Saved Pictures
			return createSpecialFolder("保存済みの写真", "shell:SavedPictures");
		case 10:
			// Win8からサポート
			// shell:My Pictures\Screenshots
			// Win＋PrtScrでスクリーンショットを保存する時に自動生成される
			return createSpecialFolder("スクリーンショット", "shell:Screenshots");
		case 11:
			// shell:My Pictures\Slide Shows
			// 手動でフォルダーを作成しても使用可
			// フォルダー名は大文字・小文字を一致させる必要あり(以下同)
			return createSpecialFolder("スライド ショー", "shell:PhotoAlbums");
		
		case 12:
			// shell:Local Videos / shell:MyComputerFolder\::{F86FA3AB-70D2-4FC7-9C99-FCBF05467F3A} (Win10から)
			// shell:MyComputerFolder\::{A0953C92-50DC-43BF-BE83-3742FED03C9C} (Win8.1から)
			return createSpecialFolder(WIN81 ? "ビデオ" : "マイ ビデオ", "shell:My Video");
		case 13:
			// Win10からサポート
			// shell:My Video\Captures
			// ゲームバーで動画やスクリーンショットを保存する時に自動生成される
			return createSpecialFolder("キャプチャ", "shell:Captures");
		
		case 14:
			// Win10 1703からサポート
			// shell:UsersFilesFolder\AppMods
			return createSpecialFolder("アプリケーションの修正", "shell:AppMods");
		case 15:
			// shell:UsersFilesFolder\{56784854-C6CB-462B-8169-88E350ACB882}
			return createSpecialFolder("アドレス帳", "shell:Contacts");
		case 16:
			return createSpecialFolder("お気に入り", "shell:Favorites");
		case 17:
			// shell:::{323CA680-C24D-4099-B94D-446DD2D7249E} ([お気に入り])
			// shell:::{D34A6CA6-62C2-4C34-8A7C-14709C1AD938} ([Common Places FS Folder])
			// shell:UsersFilesFolder\{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}
			return createSpecialFolder("リンク" + (WIN10 ? "" : " (エクスプローラーのお気に入り)"), "shell:Links");
		case 18:
			// Win10からサポート
			// shell:UsersFilesFolder\Recorded Calls
			return createSpecialFolder("録音した通話", "shell:Recorded Calls");
		case 19:
			// shell:UsersFilesFolder\{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}
			return createSpecialFolder("保存したゲーム", "shell:SavedGames");
		case 20:
			// shell:UsersFilesFolder\{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}
			return createSpecialFolder("検索", "shell:Searches");
		
		// OneDriveカテゴリのフォルダーはすべてWin8.1からサポート
		case 21:
			// shell:UsersFilesFolder\OneDrive
			// Win8.1ではMicrosoftアカウントでサインインする時に自動生成される
			// shell:::{59031A47-3F72-44A7-89C5-5595FE6B30EE}\::{8E74D236-7F35-4720-B138-1FED0B85EA75} (Win8.1のみ)
			// shell:::{59031A47-3F72-44A7-89C5-5595FE6B30EE}\::{018D5C66-4533-4307-9B53-224DE2ED1FE6} (Win10から)
			// %OneDrive% (Win10 1607から)
			return createSpecialFolder("OneDrive", "shell:OneDrive", { category: "OneDrive" });
		case 22:
			// shell:OneDrive\Documents
			return createSpecialFolder("OneDrive のドキュメント", WIN10 ? "shell:OneDriveDocuments" : "shell:SkyDriveDocuments");
		case 23:
			// shell:OneDrive\Music
			return createSpecialFolder("OneDrive のミュージック", WIN10 ? "shell:OneDriveMusic" : "shell:SkyDriveMusic");
		case 24:
			// shell:OneDrive\Pictures
			return createSpecialFolder("OneDrive の画像", WIN10 ? "shell:OneDrivePictures" : "shell:SkyDrivePictures");
		case 25:
			// shell:OneDrive\Pictures\Camera Roll
			return createSpecialFolder("OneDrive のカメラ ロール", WIN10 ? "shell:OneDriveCameraRoll" : "shell:SkyDriveCameraRoll");
		
		case 26:
			// %APPDATA%
			return createSpecialFolder("Application Data", "shell:AppData", { category: "AppData" });
		case 27:
			return createSpecialFolder("Credentials", "shell:CredentialManager");
		case 28:
			return createSpecialFolder("Crypto", "shell:CryptoKeys");
		case 29:
			return createSpecialFolder("Protect", "shell:DpapiKeys");
		case 30:
			return createSpecialFolder("SystemCertificates", "shell:SystemCertificates");
		
		case 31:
			return createSpecialFolder("クイック起動", "shell:Quick Launch");
		case 32:
			// shell:::{1F3427C8-5C10-4210-AA03-2EE45287D668}
			return createSpecialFolder("User Pinned", "shell:User Pinned");
		case 33:
			return createSpecialFolder("ImplicitAppShortcuts", "shell:ImplicitAppShortcuts");
		
		case 34:
			// Win8からサポート
			return createSpecialFolder("アカウントの画像", "shell:AccountPictures");
		case 35:
			// Win8.1以降では[LocalAppData]カテゴリになるので非表示に
			return createSpecialFolder("Cookies", WIN81 ? null : "shell:Cookies");
		case 36:
			return createSpecialFolder("Network Shortcuts", "shell:NetHood");
		case 37:
			// shell:::{ED50FC29-B964-48A9-AFB3-15EBB9B97F36} ([printhood delegate folder])
			return createSpecialFolder("Printer Shortcuts", "shell:PrintHood");
		case 38:
			return createSpecialFolder("最近使った項目", "shell:Recent");
		case 39:
			return createSpecialFolder("SendTo (送る)", "shell:SendTo");
		case 40:
			return createSpecialFolder("テンプレート", "shell:Templates");
		
		case 41:
			// shell:UsersLibrariesFolder
			// shell:::{031E4825-7B94-4DC3-B131-E946B44C8DD5}
			return createSpecialFolder("ライブラリ", "shell:Libraries", { category: "Libraries", path: LIBRARIES_PATH, folderItemForProperties: getDirectoryFolderItem(LIBRARIES_PATH) });
		case 42:
			// Win10からサポート
			// shell:Libraries\CameraRoll.library-ms
			// shell:Libraries\{2B20DF75-1EDA-4039-8097-38798227D5B7}
			return createSpecialFolder("カメラ ロール ライブラリ", "shell:CameraRollLibrary", { path: getRegValue(USER_SHELL_FOLDERS_KEY + "{2B20DF75-1EDA-4039-8097-38798227D5B7}", LIBRARIES_PATH + "\\CameraRoll.library-ms", true), propertyType: PT_SHELLEXECUTE });
		case 43:
			// shell:Libraries\{7B0DB17D-9CD2-4A93-9733-46CC89022E7C}
			return createSpecialFolder("ドキュメント ライブラリ", "shell:DocumentsLibrary", { path: getRegValue(USER_SHELL_FOLDERS_KEY + "{7B0DB17D-9CD2-4A93-9733-46CC89022E7C}", LIBRARIES_PATH + "\\Documents.library-ms", true), propertyType: PT_SHELLEXECUTE });
		case 44:
			// shell:Libraries\{2112AB0A-C86A-4FFE-A368-0DE96E47012E}
			return createSpecialFolder("ミュージック ライブラリ", "shell:MusicLibrary", { path: getRegValue(USER_SHELL_FOLDERS_KEY + "{2112AB0A-C86A-4FFE-A368-0DE96E47012E}", LIBRARIES_PATH + "\\Music.library-ms", true), propertyType: PT_SHELLEXECUTE });
		case 45:
			// shell:Libraries\{A990AE9F-A03B-4E80-94BC-9912D7504104}
			return createSpecialFolder("ピクチャ ライブラリ", "shell:PicturesLibrary", { path: getRegValue(USER_SHELL_FOLDERS_KEY + "{A990AE9F-A03B-4E80-94BC-9912D7504104}", LIBRARIES_PATH + "\\Pictures.library-ms", true), propertyType: PT_SHELLEXECUTE });
		case 46:
			// Win10からサポート
			// shell:Libraries\SavedPictures.library-ms
			// shell:Libraries\{E25B5812-BE88-4BD9-94B0-29233477B6C3}
			return createSpecialFolder("保存済みの写真 ライブラリ", "shell:SavedPicturesLibrary", { path: getRegValue(USER_SHELL_FOLDERS_KEY + "{E25B5812-BE88-4BD9-94B0-29233477B6C3}", LIBRARIES_PATH + "\\SavedPictures.library-ms", true), propertyType: PT_SHELLEXECUTE });
		case 47:
			// shell:::{031E4825-7B94-4DC3-B131-E946B44C8DD5}\{491E922F-5643-4AF4-A7EB-4E7A138D8174}
			return createSpecialFolder("ビデオ ライブラリ", "shell:VideosLibrary", { path: getRegValue(USER_SHELL_FOLDERS_KEY + "{491E922F-5643-4AF4-A7EB-4E7A138D8174}", LIBRARIES_PATH + "\\Videos.library-ms", true), propertyType: PT_SHELLEXECUTE });
		
		case 48:
			return createSpecialFolder("スタート メニュー", "shell:Start Menu", { category: "StartMenu" });
		case 49:
			return createSpecialFolder("プログラム", "shell:Programs");
		case 50:
			return createSpecialFolder(WIN10 ? "Windows 管理ツール" : "管理ツール", "shell:Administrative Tools");
		case 51:
			return createSpecialFolder("スタートアップ", "shell:Startup");
		
		case 52:
			// %LOCALAPPDATA%
			return createSpecialFolder("Local Application Data", "shell:Local AppData", { category: "LocalAppData" });
		case 53:
			return createSpecialFolder("Local Application Data Low", "shell:LocalAppDataLow");
		
		case 54:
			// Win10 1709からサポート
			// shell:Local AppData\Desktop
			return createSpecialFolder("AppDataDesktop", "shell:AppDataDesktop");
		case 55:
			// Win10からサポート
			// shell:Local AppData\DevelopmentFiles
			return createSpecialFolder("Development Files", "shell:Development Files");
		case 56:
			// Win10 1709からサポート
			// shell:Local AppData\Documents
			return createSpecialFolder("AppDataDocuments", "shell:AppDataDocuments");
		case 57:
			// Win10 1709からサポート
			// shell:Local AppData\Favorites
			return createSpecialFolder("AppDataFavorites", "shell:AppDataFavorites");
		case 58:
			// Win8からサポート
			return createSpecialFolder(WINDOWS_APPS_TITLE + "の設定", "shell:Local AppData\\Packages");
		case 59:
			// Win10 1709からサポート
			// shell:Local AppData\ProgramData
			return createSpecialFolder("AppDataProgramData", "shell:AppDataProgramData");
		case 60:
			// %TEMP%
			// %TMP%
			return createSpecialFolder("一時ファイル", fso.GetSpecialFolder(TemporaryFolder).Path);
		case 61:
			return createSpecialFolder("VirtualStore", "shell:Local AppData\\VirtualStore");
		
		case 62:
			// Win8からサポート
			return createSpecialFolder("アプリケーションのショートカット", "shell:Application Shortcuts");
		case 63:
			return createSpecialFolder("一時書き込みフォルダー", "shell:CD Burning");
		case 64:
			// Win10 1809からサポート
			// 標準ユーザー権限でフォントをインストールした時に自動生成される
			return createSpecialFolder("フォント (Per User)", "shell:Local AppData\\Microsoft\\Windows\\Fonts");
		case 65:
			return createSpecialFolder("GameExplorer", "shell:GameTasks");
		case 66:
			return createSpecialFolder("履歴", "shell:History");
		case 67:
			return createSpecialFolder("インターネット一時ファイル", "shell:Cache");
		case 68:
			// Win8.1でこのカテゴリに移動
			return createSpecialFolder("Cookies", WIN81 ? "shell:Cookies" : null);
		case 69:
			return createSpecialFolder("Ringtones", "shell:Ringtones");
		case 70:
			// Win8からサポート
			// shell:Local AppData\Microsoft\Windows\RoamedTileImages
			return createSpecialFolder("Roamed Tile Images", "shell:Roamed Tile Images");
		case 71:
			// Win8からサポート
			return createSpecialFolder("Roaming Tiles", "shell:Roaming Tiles");
		case 72:
			// Win8からサポート
			return createSpecialFolder("WinX (クイックアクセスメニュー)", WIN8 ? "shell:Local AppData\\Microsoft\\Windows\\WinX" : null);
		
		case 73:
			// Win8.1からサポート
			// shell:Local AppData\Microsoft\Windows\ConnectedSearch\History
			return createSpecialFolder("検索履歴", "shell:SearchHistoryFolder");
		case 74:
			// Win8.1からサポート
			// shell:Local AppData\Microsoft\Windows\ConnectedSearch\Templates
			return createSpecialFolder("検索テンプレート", "shell:SearchTemplatesFolder");
		
		case 75:
			return createSpecialFolder("ガジェット", WIN8 ? "shell:Local AppData\\Microsoft\\Windows Sidebar\\Gadgets" :"shell:Gadgets");
		case 76:
			// shell:Local AppData\Microsoft\Windows Photo Gallery\Original Images
			// フォトギャラリーでファイルを編集する時に自動生成される
			return createSpecialFolder("Original Images", "shell:Original Images");
		
		case 77:
			// shell:Local AppData\Programs
			return createSpecialFolder("Program Files (Per User)", "shell:UserProgramFiles");
		case 78:
			// shell:Local AppData\Programs\Common
			return createSpecialFolder("Common Program Files (Per User)", "shell:UserProgramFilesCommon");
		
		case 79:
			// shell:::{4336A54D-038B-4685-AB02-99BB52D3FB8B}
			// shell:ThisDeviceFolder ([このデバイス]) (Win10 1507から1607まで)
			// shell:::{5B934B42-522B-4C34-BBFE-37A3EF7B9C90} ([このデバイス]) (Win10から)
			// %PUBLIC%
			return createSpecialFolder("パブリック", "shell:Public", { category: "Public" });
		case 80:
			// Win8からサポート
			// shell:Public\AccountPictures
			return createSpecialFolder("パブリック アカウントの画像", "shell:PublicAccountPictures");
		case 81:
			return createSpecialFolder("パブリック デスクトップ", "shell:Common Desktop");
		case 82:
			return createSpecialFolder("パブリックのドキュメント", "shell:Common Documents");
		case 83:
			return createSpecialFolder("パブリックのダウンロード", "shell:CommonDownloads");
		case 84:
			return createSpecialFolder("ライブラリ (All Users)", "shell:PublicLibraries");
		case 85:
			// Win8.1までサポート
			// メディアセンターが使えるエディションで使用可
			return createSpecialFolder("パブリックの録画一覧", getRegValue(CURRENT_VERSION_KEY + "Media Center\\Service\\Recording\\RecordPath"));
		case 86:
			return createSpecialFolder("パブリックのミュージック", "shell:CommonMusic");
		case 87:
			// shell:CommonMusic\Sample Music
			return createSpecialFolder("サンプル ミュージック", "shell:SampleMusic");
		case 88:
			// Win7までサポート
			// shell:CommonMusic\Sample Playlists
			return createSpecialFolder("サンプル プレイリスト", "shell:SamplePlaylists");
		case 89:
			return createSpecialFolder("パブリックのピクチャ", "shell:CommonPictures");
		case 90:
			// shell:CommonPictures\Sample Pictures
			return createSpecialFolder("サンプル ピクチャ", "shell:SamplePictures");
		case 91:
			return createSpecialFolder("パブリックのビデオ", "shell:CommonVideo");
		case 92:
			// shell:CommonVideo\Sample Videos
			return createSpecialFolder("サンプル ビデオ", "shell:SampleVideos");
		
		case 93:
			// %ALLUSERSPROFILE%
			// %ProgramData%
			return createSpecialFolder("ProgramData", "shell:Common AppData", { category: "ProgramData" });
		case 94:
			// %ALLUSERSPROFILE%\OEM Links
			return createSpecialFolder("OEM Links", "shell:OEM Links");
		
		case 95:
			return createSpecialFolder("Windows Search のインデックス", getRegValue("HKLM\\SOFTWARE\\Microsoft\\Windows Search\\DataDirectory", null, true));
		case 96:
			return createSpecialFolder("既定のアカウントの画像", WIN8 ? USER_ACCOUNT_PICTURES_PATH : USER_ACCOUNT_PICTURES_PATH + "\\Default Pictures");
		
		case 97:
			// Win8からサポート
			return createSpecialFolder(WINDOWS_APPS_TITLE + "のリポジトリ", getRegValue(CURRENT_VERSION_KEY + "Appx\\PackageRepositoryRoot"));
		case 98:
			return createSpecialFolder("Device Metadata Store", "shell:Device Metadata Store");
		case 99:
			return createSpecialFolder("GameExplorer (All Users)", "shell:PublicGameTasks");
		case 100:
			// Win10からサポート
			// shell:Common AppData\Microsoft\Windows\RetailDemo
			// 市販デモ モードで使用される
			return createSpecialFolder("RetailDemo", "shell:Retail Demo");
		case 101:
			return createSpecialFolder("Ringtones (All Users)", "shell:CommonRingtones");
		case 102:
			return createSpecialFolder("テンプレート (All Users)", "shell:Common Templates");
		
		case 103:
			return createSpecialFolder("スタート メニュー (All Users)", "shell:Common Start Menu", { category: "CommonStartMenu" });
		case 104:
			return createSpecialFolder("プログラム (All Users)", "shell:Common Programs");
		case 105:
			// shell:ControlPanelFolder\::{D20EA4E1-3957-11D2-A40B-0C5020524153}
			return createSpecialFolder(WIN10 ? "Windows 管理ツール (All Users)" : "管理ツール (All Users)", "shell:Common Administrative Tools");
		case 106:
			return createSpecialFolder("スタートアップ (All Users)", "shell:Common Startup");
		case 107:
			// Win10からサポート
			return createSpecialFolder("Start Menu Places", "shell:Common Start Menu Places");
		
		case 108:
			// %SystemRoot%
			// %windir%
			return createSpecialFolder("Windows ディレクトリ", "shell:Windows", { category: "Windows" });
		case 109:
			// shell:::{1D2680C9-0E2A-469D-B787-065558BC7D43} ([Fusion Cache]) (.NET3.5まで)
			// CLSIDを使ってアクセスするとエクスプローラーがクラッシュする
			return createSpecialFolder(".NET Framework Assemblies", "shell:Windows\\assembly");
		case 110:
			return createSpecialFolder("ActiveX Cache Folder", getRegValue(CURRENT_VERSION_KEY + "Internet Settings\\ActiveXCache"));
		case 111:
			// shell:ControlPanelFolder\::{BD84B380-8CA2-1069-AB1D-08000948F534}
			return createSpecialFolder("フォント", "shell:Fonts");
		case 112:
			return createSpecialFolder("既定のサウンド", getRegValue(CURRENT_VERSION_KEY + "MediaPathUnexpanded", null, true));
		case 113:
			return createSpecialFolder("Subscription Folder", "shell:Windows\\Offline Web Pages");
		case 114:
			// HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers\knownfolders\0\Windows Wallpapers\MergeFolders
			return createSpecialFolder("壁紙", "shell:Windows\\Web");
		
		case 115:
			return createSpecialFolder("テーマのリソース", "shell:ResourceDir");
		case 116:
			// shell:ResourceDir\xxxx (xxxxはロケールIDの16進数4桁 日本語では0411)
			return createSpecialFolder("テーマのローカライズ リソース", "shell:LocalizedResourcesDir");
		
		case 117:
			return createSpecialFolder("システム ディレクトリ", IS64BIT ? "shell:System" : "shell:SystemX86");
		case 118:
			return IS64BIT ?
				createSpecialFolder("システム ディレクトリ (32 ビット)", "shell:SystemX86") :
				createSpecialFolder("システム ディレクトリ (64 ビット)", getSysNativePath());
		
		case 119:
			return createSpecialFolder("ユーザー", "shell:UserProfiles", { category: "Users" });
		case 120:
			return createSpecialFolder("既定のプロファイル", getRegValue("HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList\\Default", null, true));
		
		case 121:
			// shell:ProgramFilesX64 (64ビットアプリのみ)
			// %ProgramFiles%
			return createSpecialFolder("Program Files", "shell:ProgramFiles", { category: "ProgramFiles" });
		case 122:
			return IS64BIT ?
				createSpecialFolder("Program Files (32 ビット)", "shell:ProgramFilesX86") :
				createSpecialFolder("Program Files (64 ビット)", getRegValue(CURRENT_VERSION_KEY + "ProgramW6432Dir"));
		case 123:
			// shell:ProgramFilesCommonX64 (64ビットアプリのみ)
			// %CommonProgramFiles%
			return createSpecialFolder("Common Program Files", "shell:ProgramFilesCommon");
		case 124:
			return IS64BIT ?
				createSpecialFolder("Common Program Files (32 ビット)", "shell:ProgramFilesCommonX86") :
				createSpecialFolder("Common Program Files (64 ビット)", getRegValue(CURRENT_VERSION_KEY + "CommonW6432Dir"));
		case 125:
			// Win8からサポート
			return createSpecialFolder(WINDOWS_APPS_TITLE, getRegValue(CURRENT_VERSION_KEY + "Appx\\PackageRoot"));
		case 126:
			return createSpecialFolder("既定のガジェット", WIN8 ? "shell:ProgramFiles\\Windows Sidebar\\Gadgets" : "shell:Default Gadgets");
		case 127:
			return createSpecialFolder("ガジェット (All Users)", "shell:ProgramFiles\\Windows Sidebar\\Shared Gadgets");
		
		case 128:
			return createSpecialFolder("デスクトップ", "shell:Desktop", { category: "Desktop / MyComputer" });
		case 129:
			// shell:MyComputerFolderはWin10 1507/1511だとなぜかデスクトップになってしまう
			return createSpecialFolder(WIN81 ? "PC" : "コンピューター", !WIN10_1507_to_1511 ? "shell:MyComputerFolder" : "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}");
		case 130:
			return createSpecialFolder(WIN10 ? "最近使ったフォルダー" : "最近表示した場所", "shell:::{22877A6D-37A1-461A-91B0-DBDA5AAEBC99}");
		case 131:
			// Win10からサポート
			// shell:::{4564B25E-30CD-4787-82BA-39E73A750B14} ([Recent Items Instance Folder])
			return createSpecialFolder("最近使用したファイル", "shell:::{3134EF9C-6B18-4996-AD04-ED5912E00EB5}");
		case 132:
			return createSpecialFolder("ポータブル メディア デバイス", "shell:::{35786D3C-B075-49B9-88DD-029876E11C01}");
		case 133:
			// Win10からサポート
			return createSpecialFolder("よく使用するフォルダー", "shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}");
		case 134:
			return createSpecialFolder("ごみ箱", "shell:RecycleBinFolder");
		case 135:
			// Win10からサポート
			return createSpecialFolder("クイック アクセス", "shell:::{679F85CB-0220-4080-B29B-5540CC05AAB6}");
		case 136:
			// Win8からサポート
			// Win8/8.1では[PC]と同じなので非表示に
			return createSpecialFolder("Removable Storage Devices", WIN10 ? "shell:::{A6482830-08EB-41E2-84C1-73920C2BADB9}": null);
		case 137:
			return createSpecialFolder("ホームグループ", "shell:HomeGroupFolder");
		case 138:
			return createSpecialFolder("ネットワーク", "shell:NetworkPlacesFolder");
		case 139:
			// Win10からサポート
			return createSpecialFolder("Removable Drives", "shell:::{F5FB2C77-0E2F-4A16-A381-3E560C68BC83}");
		
		case 140:
			return createSpecialFolder("コントロール パネル (カテゴリ表示)", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}", { category: "ControlPanel" });
		case 141:
			return createSpecialFolder("デスクトップのカスタマイズ", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\1");
		case 142:
			// shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\4
			return createSpecialFolder("ハードウェアとサウンド", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\2");
		case 143:
			return createSpecialFolder("ネットワークとインターネット", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\3");
		case 144:
			// shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\10
			return createSpecialFolder("システムとセキュリティ", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\5");
		case 145:
			return createSpecialFolder(WIN10_1803 ? "時計と地域" : "時計、言語、および地域", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\6");
		case 146:
			return createSpecialFolder("コンピューターの簡単操作", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\7");
		case 147:
			return createSpecialFolder("プログラム", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\8");
		case 148:
			return createSpecialFolder(WIN10 ? "ユーザー アカウント" : WIN8 ? "ユーザー アカウントとファミリー セーフティ" : "ユーザー アカウントと家族のための安全設定", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\9");
		
		case 149:
			// shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}
			// shell:::{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}
			// shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\11
			return createSpecialFolder("すべてのコントロール パネル項目", "shell:ControlPanelFolder");
		
		// コントロールパネル内の項目はCLSIDだけを指定してもアクセス可能
		// 例えば[電源オプション]なら shell:::{025A5937-A6BE-4686-A844-36FE4BEC8B6D}
		// ただしその場合はアドレスバーからコントロールパネルに移動できない
		case 150:
			// Win7/8のみサポート
			return createSpecialFolder("既定の位置", "shell:ControlPanelFolder\\::{00C6D95F-329C-409A-81D7-C46C66EA7F33}");
		case 151:
			// Win7/8のみサポート
			return createSpecialFolder("生体認証デバイス", "shell:ControlPanelFolder\\::{0142E4D0-FB7A-11DC-BA4A-000FFE7AB428}");
		case 152:
			return createSpecialFolder("電源オプション", "shell:ControlPanelFolder\\::{025A5937-A6BE-4686-A844-36FE4BEC8B6D}");
		case 153:
			return createSpecialFolder("資格情報マネージャー", "shell:ControlPanelFolder\\::{1206F5F1-0569-412C-8FEC-3204630DFB70}");
		case 154:
			return createSpecialFolder("プログラムの取得", "shell:AddNewProgramsFolder");
		case 155:
			// shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{E44E5D18-0652-4508-A4E2-8A090067BCB0}
			return createSpecialFolder("既定のプログラム", "shell:ControlPanelFolder\\::{17CD9488-1228-4B2F-88CE-4298E93E0966}");
		case 156:
			return createSpecialFolder("RemoteApp とデスクトップ接続", "shell:ControlPanelFolder\\::{241D7C96-F8BF-4F85-B01F-E2B043341A4B}");
		case 157:
			// Win8.1までサポート
			return createSpecialFolder("Windows Update", "shell:ControlPanelFolder\\::{36EEF7DB-88AD-4E81-AD49-0E313F0C35F8}");
		case 158:
			return createSpecialFolder(WIN10_1709 ? "Windows Defender ファイアウォール" : "Windows ファイアウォール", "shell:ControlPanelFolder\\::{4026492F-2F69-46B8-B9BF-5654FC07E423}");
		case 159:
			return createSpecialFolder("音声認識", "shell:ControlPanelFolder\\::{58E3C745-D971-4081-9034-86E34B30836A}");
		case 160:
			return createSpecialFolder("ユーザー アカウント", "shell:ControlPanelFolder\\::{60632754-C523-4B62-B45C-4172DA012619}");
		case 161:
			// Win10 1709までサポート
			return createSpecialFolder("ホームグループ", "shell:ControlPanelFolder\\::{67CA7650-96E6-4FDD-BB43-A8E774F73A57}");
		case 162:
			// Win8までサポート
			return createSpecialFolder("パフォーマンスの情報とツール", "shell:ControlPanelFolder\\::{78F3955E-3B90-4184-BD14-5397C15F1EFC}");
		case 163:
			return createSpecialFolder("ネットワークと共有センター", "shell:ControlPanelFolder\\::{8E908FC9-BECC-40F6-915B-F4CA0E70D03D}");
		case 164:
			return createSpecialFolder(WIN8 ? "ファミリー セーフティ" : "保護者による制限", "shell:ControlPanelFolder\\::{96AE8D84-A250-4520-95A5-A47A7E3C548B}");
		case 165:
			return createSpecialFolder("自動再生", "shell:ControlPanelFolder\\::{9C60DE1E-E5FC-40F4-A487-460851A8D915}");
		case 166:
			return createSpecialFolder("回復", "shell:ControlPanelFolder\\::{9FE63AFD-59CF-4419-9775-ABCC3849F861}");
		case 167:
			return createSpecialFolder("デバイスとプリンター", "shell:ControlPanelFolder\\::{A8A91A66-3A7D-4424-8D24-04E180695C7A}");
		case 168:
			// Win8.1以外でサポート
			return createSpecialFolder(WIN10 ? "バックアップと復元 (Windows 7)" : "バックアップと復元", "shell:ControlPanelFolder\\::{B98A2BEA-7D42-4558-8BD1-832F41BAC6FD}");
		case 169:
			return createSpecialFolder("システム", "shell:ControlPanelFolder\\::{BB06C0E4-D293-4F75-8A90-CB05B6477EEE}");
		case 170:
			return createSpecialFolder(WIN10 ? "セキュリティとメンテナンス" : "アクション センター", "shell:ControlPanelFolder\\::{BB64F8A7-BEE7-4E1A-AB8D-7D8273F7FDB6}");
		case 171:
			// shell:Fonts
			return createSpecialFolder("フォント", "shell:ControlPanelFolder\\::{BD84B380-8CA2-1069-AB1D-08000948F534}", { path: "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\0\\::{BD84B380-8CA2-1069-AB1D-08000948F534}" });
		case 172:
			// Win7からWin10 1803までサポート
			return createSpecialFolder("言語", "shell:ControlPanelFolder\\::{BF782CC9-5A52-4A17-806C-2A894FFEEAC5}");
		case 173:
			// Win10 1607までサポート
			return createSpecialFolder("ディスプレイ", "shell:ControlPanelFolder\\::{C555438B-3C23-4769-A71F-B6D3D9B6053A}");
		case 174:
			return createSpecialFolder("トラブルシューティング", "shell:ControlPanelFolder\\::{C58C4893-3BE0-4B45-ABB5-A63E4B8C8651}");
		case 175:
			// Win7までサポート
			return createSpecialFolder("はじめに", "shell:ControlPanelFolder\\::{CB1B7F8C-C50A-4176-B604-9E24DEE8D4D1}");
		case 176:
			// shell:Common Administrative Tools
			return createSpecialFolder("管理ツール", "shell:ControlPanelFolder\\::{D20EA4E1-3957-11D2-A40B-0C5020524153}", { path: "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\0\\::{D20EA4E1-3957-11D2-A40B-0C5020524153}" });
		case 177:
			return createSpecialFolder("コンピューターの簡単操作センター", "shell:ControlPanelFolder\\::{D555645E-D4F8-4C29-A827-D93C859C4F2A}");
		case 178:
			// Enterprise/Ultimateで使用可
			// Win8からはProでも使用可
			// Win8.1からはCore/Homeでも使用可
			return createSpecialFolder(shell.GetSystemInformation("IsOS_Personal") ? "デバイスの暗号化" : "BitLocker ドライブ暗号化", "shell:ControlPanelFolder\\::{D9EF8727-CAC2-4E60-809E-86F80A666C91}");
		case 179:
			// Win7までサポート
			return createSpecialFolder("ネットワーク マップ", "shell:ControlPanelFolder\\::{E7DE9B1A-7533-4556-9484-B26FB486475E}");
		case 180:
			// Win8までサポート
			return createSpecialFolder("Windows SideShow", "shell:ControlPanelFolder\\::{E95A4861-D57A-4BE1-AD0F-35267E261739}");
		case 181:
			// Win8.1までサポート
			return createSpecialFolder(WIN8 ? "位置情報の設定" : "位置センサーとその他のセンサー", "shell:ControlPanelFolder\\::{E9950154-C418-419E-A90A-20C5287AE24B}");
		case 182:
			// Win8.1からサポート
			// Win7ではKB2891638をインストールすれば使用可
			return createSpecialFolder("ワーク フォルダー", "shell:ControlPanelFolder\\::{ECDB0924-4208-451E-8EE0-373C0956DE16}");
		case 183:
			return createSpecialFolder(WIN10_1607 ? "個人用設定" : "個人設定", "shell:ControlPanelFolder\\::{ED834ED6-4B5A-4BFE-8F11-A626DCB6A921}");
		case 184:
			// Win8からサポート
			return createSpecialFolder("ファイル履歴", "shell:ControlPanelFolder\\::{F6B6E965-E9B2-444B-9286-10C9152EDBC5}");
		case 185:
			// Win8からサポート
			return createSpecialFolder("記憶域", "shell:ControlPanelFolder\\::{F942C606-0914-47AB-BE56-1321B8035096}");
		
		case 186:
			return createSpecialFolder("プログラムと機能", "shell:ChangeRemoveProgramsFolder");
		case 187:
			return createSpecialFolder("インストールされた更新プログラム", "shell:AppUpdatesFolder");
		
		case 188:
			return createSpecialFolder("同期センター", "shell:SyncCenterFolder");
		case 189:
			// shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{2E9E59C0-B437-4981-A647-9C34B9B90891} ([Sync Setup Folder])
			return createSpecialFolder("同期のセットアップ", "shell:SyncSetupFolder");
		case 190:
			return createSpecialFolder("同期の競合", "shell:ConflictFolder");
		case 191:
			return createSpecialFolder("同期結果", "shell:SyncResultsFolder");
		
		case 192:
			return createSpecialFolder("通知領域アイコン", "shell:" + (WIN10 ? "::{21EC2020-3AEA-1069-A2DD-08002B30309D}" : "ControlPanelFolder") + "\\::{05D7B0F4-2121-4EFF-BF6B-ED3F69B894D9}");
		case 193:
			// Win8.1 Update以降ではフォルダーを開けないので非表示に
			return createSpecialFolder("ワイヤレス ネットワークの管理", WIN81 ? null : "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\\::{1FA9085F-25A2-489B-85D4-86326EEDCD87}");
		case 194:
			// shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{863AA9FD-42DF-457B-8E4D-0DE1B8015C60}
			return createSpecialFolder("プリンター", "shell:PrintersFolder");
		case 195:
			return createSpecialFolder("Bluetooth デバイス", "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\\::{28803F59-3A75-4058-995F-4EE5503B023C}");
		case 196:
			// shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{992CFFA0-F557-101A-88EC-00DD010CCC48}
			return createSpecialFolder("ネットワーク接続", "shell:ConnectionsFolder");
		case 197:
			return createSpecialFolder("フォント設定", "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\\::{93412589-74D4-4E4E-AD0E-E0CB621440FD}");
		case 198:
			return createSpecialFolder("すべてのタスク (God Mode)", "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\\::{ED7BA470-8E54-465E-825C-99712043E01C}");
		
		case 199:
			// クライアントHyper-Vを有効にすると利用可
			return createSpecialFolder("リモート ファイル ブラウザー", "shell:::{0907616E-F5E6-48D8-9D61-A91C3D28106D}", { category: "OtherFolders" });
		case 200:
			return createSpecialFolder("キャビネット ファイル ビューアー", "shell:::{0CD7A5C0-9F37-11CE-AE65-08002B2E1262}");
		case 201:
			return createSpecialFolder("ネットワーク (WORKGROUP)", "shell:::{208D2C60-3AEA-1069-A2D7-08002B30309D}");
		case 202:
			// Win8からサポート
			return createSpecialFolder("メディア サーバー", "shell:::{289AF617-1CC3-42A6-926C-E6A863F0E3BA}");
		case 203:
			return createSpecialFolder("Results Folder", "shell:::{2965E715-EB66-4719-B53F-1672673BBEFA}");
		case 204:
			// Explorer Browser Results Folder
			return createSpecialFolder("", "shell:::{418C8B64-5463-461D-88E0-75E2AFA3C6FA},");
		case 205:
			// Win8からサポート
			return createSpecialFolder("Applications (すべてのアプリ)", "shell:AppsFolder");
		case 206:
			return createSpecialFolder("Command Folder", "shell:::{437FF9C0-A07F-4FA0-AF80-84B6C6440A16}");
		case 207:
			return createSpecialFolder("ホームグループ", "shell:::{6785BFAC-9D2D-4BE5-B7E2-59937E8FB80A}");
		case 208:
			// Win8.1 Update以降ではフォルダーを開けないので非表示に
			return createSpecialFolder("Programs Folder (すべてのプログラム)", WIN81 ? null : "shell:::{7BE9D83C-A729-4D97-B5A7-1B7313C39E0A}");
		case 209:
			// Win8.1 Update以降ではフォルダーを開けないので非表示に
			return createSpecialFolder("Programs Folder and Fast Items (すべてのプログラム (ファイルを先に表示))", WIN81 ? null : "shell:::{865E5E76-AD83-4DCA-A109-50DC2113CE9A}");
		case 210:
			// search:
			// search-ms:
			return createSpecialFolder("検索結果", "shell:SearchHomeFolder");
		case 211:
			// Win8.1 UpdateからWin10 1511までサポート
			return createSpecialFolder("StartMenuAllPrograms", "shell:StartMenuAllPrograms");
		case 212:
			// 企業向けエディションで使用可
			return createSpecialFolder("オフライン ファイル フォルダー", "shell:::{AFDB1F70-2A4C-11D2-9039-00C04F8EEB3E}");
		case 213:
			return createSpecialFolder("delegate folder that appears in Computer", "shell:::{B155BDF8-02F0-451E-9A26-AE317CFD7779}");
		case 214:
			return createSpecialFolder("AppSuggestedLocations", "shell:::{C57A6066-66A3-4D91-9EB9-41532179F0A5}");
		case 215:
			// Win10 1709までサポート
			return createSpecialFolder("ゲーム", "shell:Games");
		case 216:
			if (!Setting.debug) index.current = DONE_ITERATION;
			
			return createSpecialFolder("Previous Versions Results Folder", "shell:::{F8C2AB3B-17BC-41DA-9758-339D7DBF2D88}");
		
		// 通常とは違う名前がエクスプローラーのタイトルバーに表示されるフォルダー
		case 217:
			// Win10から
			// shell:::{5B934B42-522B-4C34-BBFE-37A3EF7B9C90} (Win10 1507から1607まで)
			// shell:::{F8278C54-A712-415B-B593-B77A2BE0DDA9} (Win10 1703から)
			return createSpecialFolder("このデバイス ({0})".xFormat(WIN10_1703 ? "個人用フォルダー" : "パブリック"), "shell:ThisDeviceFolder", { category: "OtherNames", propertyType: PT_VERB });
		case 218:
			// Win8までだと別名にならないので非表示に
			return createSpecialFolder("マイ ドキュメント (ドキュメント)", WIN81 ? "shell:::{450D8FBA-AD25-11D0-98A8-0800361B1103}" : null);
		case 219:
			return createSpecialFolder("お気に入り (リンク)", "shell:::{323CA680-C24D-4099-B94D-446DD2D7249E}");
		case 220:
			return createSpecialFolder("Common Places FS Folder (リンク)", "shell:::{D34A6CA6-62C2-4C34-8A7C-14709C1AD938}", { propertyType: PT_VERB });
		case 221:
			return createSpecialFolder("printhood delegate folder (Printer Shortcuts)", "shell:::{ED50FC29-B964-48A9-AFB3-15EBB9B97F36}");
		case 222:
			// .NET3.5まで
			// CLSIDを使ってアクセスするとエクスプローラーがクラッシュする
			return createSpecialFolder("Fusion Cache (.NET Framework Assemblies)", "shell:::{1D2680C9-0E2A-469D-B787-065558BC7D43}");
		case 223:
			// Win10から
			return createSpecialFolder("Recent Items Instance Folder (最近使用したファイル)", "shell:::{4564B25E-30CD-4787-82BA-39E73A750B14}");
		case 224:
			return createSpecialFolder("Sync Setup Folder (同期のセットアップ)", "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\\::{2E9E59C0-B437-4981-A647-9C34B9B90891}");

		// エクスプローラーで開けないフォルダー
		case 225:
			// 検索
			return createSpecialFolder("検索", "shell:::{04731B67-D933-450A-90E6-4ACD2E9408FE}", { category: "CantOpen" });
		case 226:
			// Win8.1 Updateでこのカテゴリに移動
			return createSpecialFolder("ワイヤレス ネットワークの管理", WIN81 ? "shell:::{1FA9085F-25A2-489B-85D4-86326EEDCD87}" : null);
		case 227:
			return createSpecialFolder("Sync Center Conflict Folder", "shell:::{289978AC-A101-4341-A817-21EBA7FD046D}");
		case 228:
			return createSpecialFolder("LayoutFolder", "shell:::{328B0346-7EAF-4BBE-A479-7CB88A095F5B}");
		case 229:
			return createSpecialFolder("Explorer Browser Results Folder", "shell:::{418C8B64-5463-461D-88E0-75E2AFA3C6FA}");
		case 230:
			// Win8.1から
			return createSpecialFolder("すべての設定", "shell:::{5ED4F38C-D3FF-4D61-B506-6820320AEBFE}");
		case 231:
			return createSpecialFolder("Microsoft FTP Folder", "shell:::{63DA6EC0-2E98-11CF-8D82-444553540000}");
		case 232:
			// Win8から
			return createSpecialFolder("CLSID_AppInstanceFolder", "shell:::{64693913-1C21-4F30-A98F-4E52906D3B56}");
		case 233:
			return createSpecialFolder("Sync Results Folder", "shell:::{71D99464-3B6B-475C-B241-E15883207529}");
		case 234:
			// Win8.1 Updateでこのカテゴリに移動
			// Win10 1511まで
			return createSpecialFolder("Programs Folder", WIN81 ? "shell:::{7BE9D83C-A729-4D97-B5A7-1B7313C39E0A}" : null);
		case 235:
			// Win8.1 Updateでこのカテゴリに移動
			// Win10 1511まで
			return createSpecialFolder("Programs Folder and Fast Items", WIN81 ? "shell:::{865E5E76-AD83-4DCA-A109-50DC2113CE9A}" : null);
		case 236:
			// Win10でこのカテゴリに移動
			return createSpecialFolder("インターネット", WIN10 ? "shell:InternetFolder" : null);
		case 237:
			return createSpecialFolder("File Backup Index", "shell:::{877CA5AC-CB41-4842-9C69-9136E42D47E2}");
		case 238:
			return createSpecialFolder("Microsoft Office Outlook", "shell:::{89D83576-6BD1-4C86-9454-BEB04E94C819}");
		case 239:
			return createSpecialFolder("DXP", "shell:::{8FD8B88D-30E1-4F25-AC2B-553D3D65F0EA}");
		case 240:
			return createSpecialFolder("Enhanced Storage Data Source", "shell:::{9113A02D-00A3-46B9-BC5F-9C04DADDD5D7}");
		case 241:
			// Win8から
			return createSpecialFolder("CLSID_StartMenuLauncherProviderFolder", "shell:::{98F275B4-4FFF-11E0-89E2-7B86DFD72085}");
		case 242:
			return createSpecialFolder("IE RSS Feeds Folder", "shell:::{9A096BB5-9DC3-4D1C-8526-C3CBF991EA4E}");
		case 243:
			// Win8から
			return createSpecialFolder("CLSID_StartMenuCommandingProviderFolder", "shell:::{A00EE528-EBD9-48B8-944A-8942113D46AC}");
		case 244:
			return createSpecialFolder("Previous Versions Results Delegate Folder", "shell:::{A3C3D402-E56C-4033-95F7-4885E80B0111}");
		case 245:
			return createSpecialFolder("Library Folder", "shell:::{A5A3563A-5755-4A6F-854E-AFA3230B199F}");
		case 246:
			// Win8からサポート
			return createSpecialFolder("ホームグループ内の現在のユーザー", "shell:HomeGroupCurrentUserFolder");
		case 247:
			return createSpecialFolder("Sync Results Delegate Folder", "shell:::{BC48B32F-5910-47F5-8570-5074A8A5636A}");
		case 248:
			return createSpecialFolder("オフライン ファイル", "shell:::{BD7A2E7B-21CB-41B2-A086-B309680C6B7E}");
		case 249:
			// Win8から
			return createSpecialFolder("DLNA Content Directory Data Source", "shell:::{D2035EDF-75CB-4EF1-95A7-410D9EE17170}");
		case 250:
			return createSpecialFolder("CLSID_StartMenuProviderFolder", "shell:::{DAF95313-E44D-46AF-BE1B-CBACEA2C3065}");
		case 251:
			return createSpecialFolder("CLSID_StartMenuPathCompleteProviderFolder", "shell:::{E345F35F-9397-435C-8F95-4E922C26259E}");
		case 252:
			return createSpecialFolder("Sync Center Conflict Delegate Folder", "shell:::{E413D040-6788-4C22-957E-175D1C513A34}");
		case 253:
			return createSpecialFolder("Shell DocObject Viewer", "shell:::{E7E4BC40-E76A-11CE-A9BB-00AA004AE837}");
		case 254:
			// Win8から
			return createSpecialFolder("StreamBackedFolder", "shell:::{EDC978D6-4D53-4B2F-A265-5805674BE568}");
		case 255:
			return createSpecialFolder("Sync Setup Delegate Folder", "shell:::{F1390A9A-A3F4-4E5D-9C5F-98F3BD8D935C}");
		case 256:
			return createSpecialFolder("オフライン ファイル", "shell:CSCFolder");
		
		case 257:
			if (true) ++index.current;
			else {
				// Win8からサポート
				// ファイル履歴を有効にすると利用可
				// スクリプトのホストプログラムのプロセスが残り続ける
				// @ts-ignore
				return createSpecialFolder("FileHistoryDataSource", "shell:::{2F6CE85C-F9EE-43CA-90C7-8A9BD53A2467}");
			}
		
		// フォルダー以外のshellコマンド
		case 258:
			return createSpecialFolder(WIN8 ? "タスク バーとナビゲーション" : "タスク バーと[スタート]メニュー", "shell:::{0DF44EAA-FF21-4412-828E-260A8728E7F1}", { category: "OtherShellCommands" });
		case 259:
			// Win10 1511まで
			return createSpecialFolder(WIN8 ? "検索 - ファイル" : "検索", "shell:::{2559A1F0-21D7-11D4-BDAF-00C04F60B9F0}");
		case 260:
			// Win8.1まで
			return createSpecialFolder("ヘルプとサポート", "shell:::{2559A1F1-21D7-11D4-BDAF-00C04F60B9F0}");
		case 261:
			// Windows Security
			// Ctrl+Alt+Delと同じ
			// Win10 1903(?)でこのカテゴリに移動
			return createSpecialFolder("Windows Security", "shell:::{2559A1F2-21D7-11D4-BDAF-00C04F60B9F0}");
		case 262:
			return createSpecialFolder("ファイル名を指定して実行", "shell:::{2559A1F3-21D7-11D4-BDAF-00C04F60B9F0}");
		case 263:
			return createSpecialFolder("電子メール", "shell:::{2559A1F5-21D7-11D4-BDAF-00C04F60B9F0}");
		case 264:
			return createSpecialFolder("プログラムのアクセスとコンピューターの既定の設定", "shell:::{2559A1F7-21D7-11D4-BDAF-00C04F60B9F0}");
		case 265:
			// Win8から
			return createSpecialFolder(WIN10 ? "Cortana" : "検索", "shell:::{2559A1F8-21D7-11D4-BDAF-00C04F60B9F0}");
		case 266:
			// Win+Dと同じ
			return createSpecialFolder("デスクトップの表示", "shell:::{3080F90D-D7AD-11D9-BD98-0000947B0257}");
		case 267:
			// Win7ではCtrl+Win+Tab、Win8/8.1ではCtrl+Alt+Tab、Win10 1607以降ではWin+Tabと同じ (Win10 1507/1511では使用不可)
			return createSpecialFolder("ウィンドウを切り替える", "shell:::{3080F90E-D7AD-11D9-BD98-0000947B0257}");
		case 268:
			// Win7まで
			return createSpecialFolder("ガジェット ギャラリー", "shell:::{37EFD44D-EF8D-41B1-940D-96973A50E9E0}");
		case 269:
			return createSpecialFolder("接続先", "shell:::{38A98528-6CBF-4CA9-8DC0-B1E1D10F7B1B}");
		case 270:
			return createSpecialFolder("電話とモデム", "shell:::{40419485-C444-4567-851A-2DD7BFA1684D}");
		case 271:
			// Win8.1から
			return createSpecialFolder("新しいウィンドウで開く", "shell:::{52205FD8-5DFB-447D-801A-D0B52F2E83E1}");
		case 272:
			return createSpecialFolder("Windows モビリティ センター", "shell:::{5EA4F148-308C-46D7-98A9-49041B1DD468}");
		case 273:
			return createSpecialFolder(WIN8 ? "地域" : "地域と言語", "shell:::{62D8ED13-C9D0-4CE8-A914-47DD628FB1B0}");
		case 274:
			return createSpecialFolder("Windows の機能", "shell:::{67718415-C450-4F3C-BF8A-B487642DC39B}");
		case 275:
			return createSpecialFolder("マウス", "shell:::{6C8EEC18-8D75-41B2-A177-8831D59D2D50}");
		case 276:
			return createSpecialFolder(WIN10 ? "エクスプローラーのオプション" : "フォルダー オプション", "shell:::{6DFD7C5C-2451-11D3-A299-00C04F8EF6AF}");
		case 277:
			return createSpecialFolder("キーボード", "shell:::{725BE8F7-668E-4C7B-8F90-46BDB0936430}");
		case 278:
			return createSpecialFolder("デバイス マネージャー", "shell:::{74246BFC-4C96-11D0-ABEF-0020AF6B0B7A}");
		case 279:
			// Win8まで
			return createSpecialFolder("Windows CardSpace ", "shell:::{78CB147A-98EA-4AA6-B0DF-C8681F69341C}");
		case 280:
			// netplwiz.exe / control.exe userpasswords2
			return createSpecialFolder("ユーザー アカウント", "shell:::{7A9D77BD-5403-11D2-8785-2E0420524153}");
		case 281:
			return createSpecialFolder(WIN8 ? "タブレット PC 設定" : "Tablet PC 設定", "shell:::{80F3F1D5-FECA-45F3-BC32-752C152E456E}");
		case 282:
			// Win10以降では開けないので非表示に
			return createSpecialFolder("インターネット", WIN10 ? null : "shell:InternetFolder");
		case 283:
			return createSpecialFolder("インデックスのオプション", "shell:::{87D66A43-7B11-4A28-9811-C86EE395ACF7}");
		case 284:
			// Win8から
			// Enterpriseで使用可
			// Win10 1607以降ではProでも使用可
			return createSpecialFolder("Windows To Go ワークスペースの作成", "shell:::{8E0C279D-0BD1-43C3-9EBD-31C3DC5B8A77}");
		case 285:
			// Win8まで
			return createSpecialFolder("生体認証デバイスへようこそ", "shell:::{8E35B548-F174-4C7D-81E2-8ED33126F6FD}");
		case 286:
			// Win10 1607から1809まで(?)
			return createSpecialFolder("赤外線", "shell:::{A0275511-0E86-4ECA-97C2-ECD8F1221D08}");
		case 287:
			return createSpecialFolder("インターネット オプション", "shell:::{A3DD4F92-658A-410F-84FD-6FBBBEF2FFFE}");
		case 288:
			return createSpecialFolder("色の管理", "shell:::{B2C761C6-29BC-4F19-9251-E6195265BAF1}");
		case 289:
			// Win8.1まで
			return createSpecialFolder(WIN8 ? "Windows への機能の追加" : "Windows Anytime Upgrade", "shell:::{BE122A0E-4503-11DA-8BDE-F66BAD1E3F3A}");
		case 290:
			// Win8まで
			// shell:::{CCFB7955-B4DC-42CE-893D-884D72DD6B19}
			return createSpecialFolder("生体認証デバイス メッセージ", "shell:::{CBC84B69-69EA-439B-B791-DF15F60333CF}");
		case 291:
			return createSpecialFolder("音声合成", "shell:::{D17D1D6D-CC3F-4815-8FE3-607E7D5D10B3}");
		case 292:
			return createSpecialFolder("ネットワークの場所の追加", "shell:::{D4480A50-BA28-11D1-8E75-00C04FA31A86}");
		case 293:
			// Win10 1607まで
			return createSpecialFolder("Windows Defender", "shell:::{D8559EB9-20C0-410E-BEDA-7ED416AECC2A}");
		case 294:
			return createSpecialFolder("日付と時刻", "shell:::{E2E7934B-DCE5-43C4-9576-7FE4F75E7480}");
		case 295:
			return createSpecialFolder("サウンド", "shell:::{F2DDFC82-8F12-4CDD-B7DC-D4FE1425AA4D}");
		case 296:
			return createSpecialFolder("ペンとタッチ", "shell:::{F82DF8F7-8B9F-442E-A48C-818EA735FF9B}");
		
		// 上にあるのとは違うデータでフォルダーの情報を取得する
		// CSIDLは扱わない
		case 297:
			return createSpecialFolder("shell:Profile", "shell:Profile", { category: "OtherDirs" });
		case 298:
			return createSpecialFolder("shell:Local Documents", "shell:Local Documents");
		case 299:
			return createSpecialFolder("shell:Local Downloads", "shell:Local Downloads");
		case 300:
			return createSpecialFolder("shell:Local Music", "shell:Local Music");
		case 301:
			return createSpecialFolder("shell:Local Pictures", "shell:Local Pictures");
		case 302:
			return createSpecialFolder("shell:Local Videos", "shell:Local Videos");
		case 303:
			return createSpecialFolder("アドレス帳", "shell:UsersFilesFolder\\{56784854-C6CB-462B-8169-88E350ACB882}");
		case 304:
			return createSpecialFolder("リンク", "shell:UsersFilesFolder\\{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}");
		case 305:
			return createSpecialFolder("保存したゲーム", "shell:UsersFilesFolder\\{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}");
		case 306:
			return createSpecialFolder("検索", "shell:UsersFilesFolder\\{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}");
		case 307:
			return createSpecialFolder("shell:UsersLibrariesFolder", "shell:UsersLibrariesFolder");
		case 308:
			return createSpecialFolder("カメラ ロール ライブラリ", "shell:Libraries\\{2B20DF75-1EDA-4039-8097-38798227D5B7}");
		case 309:
			return createSpecialFolder("ドキュメント ライブラリ", "shell:Libraries\\{7B0DB17D-9CD2-4A93-9733-46CC89022E7C}");
		case 310:
			return createSpecialFolder("ミュージック ライブラリ", "shell:Libraries\\{2112AB0A-C86A-4FFE-A368-0DE96E47012E}");
		case 311:
			return createSpecialFolder("ピクチャ ライブラリ", "shell:Libraries\\{A990AE9F-A03B-4E80-94BC-9912D7504104}");
		case 312:
			return createSpecialFolder("保存済みの写真 ライブラリ", "shell:Libraries\\{E25B5812-BE88-4BD9-94B0-29233477B6C3}");
		case 313:
			return createSpecialFolder("ビデオ ライブラリ", "shell:Libraries\\{491E922F-5643-4AF4-A7EB-4E7A138D8174}");
		case 314:
			// 64ビットアプリのみ
			return createSpecialFolder("shell:ProgramFilesX64", "shell:ProgramFilesX64");
		case 315:
			// 64ビットアプリのみ
			return createSpecialFolder("shell:ProgramFilesCommonX64", "shell:ProgramFilesCommonX64");
		case 316:
			return createSpecialFolder("shell:MyComputerFolder", "shell:MyComputerFolder");
		
		case 317:
			return createSpecialFolder("%USERPROFILE%", "%USERPROFILE%".xExpand());
		case 318:
			return createSpecialFolder("%HOMEDRIVE%%HOMEPATH%", "%HOMEDRIVE%%HOMEPATH%".xExpand());
		case 319:
			return createSpecialFolder("%OneDrive%", "%OneDrive%".xExpand());
		case 320:
			return createSpecialFolder("%APPDATA%", "%APPDATA%".xExpand());
		case 321:
			return createSpecialFolder("%LOCALAPPDATA%", "%LOCALAPPDATA%".xExpand());
		case 322:
			return createSpecialFolder("%PUBLIC%", "%PUBLIC%".xExpand());
		case 323:
			return createSpecialFolder("%ALLUSERSPROFILE%", "%ALLUSERSPROFILE%".xExpand());
		case 324:
			return createSpecialFolder("%ProgramData%", "%ProgramData%".xExpand());
		case 325:
			return createSpecialFolder("%SystemRoot%", "%SystemRoot%".xExpand());
		case 326:
			return createSpecialFolder("%windir%", "%windir%".xExpand());
		case 327:
			return createSpecialFolder("%ProgramFiles%", "%ProgramFiles%".xExpand());
		case 328:
			return createSpecialFolder("%CommonProgramFiles%", "%CommonProgramFiles%".xExpand());
		
		case 329:
			// Win10から
			return createSpecialFolder("OneDrive", "shell:::{018D5C66-4533-4307-9B53-224DE2ED1FE6}");
		case 330:
			// ライブラリ
			return createSpecialFolder("UsersLibraries", "shell:::{031E4825-7B94-4DC3-B131-E946B44C8DD5}");
		case 331:
			// Win10から
			return createSpecialFolder("Local Downloads", "shell:::{088E3905-0323-4B02-9826-5D99428E115F}");
		case 332:
			// Win10 1709から
			return createSpecialFolder("3D Object", "shell:::{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}");
		case 333:
			// プログラムの取得
			return createSpecialFolder("Install New Programs", "shell:::{15EAE92E-F17A-4431-9F28-805E482DAFD4}");
		case 334:
			// Win8.1から
			return createSpecialFolder("My Music", "shell:::{1CF1260C-4DD0-4EBB-811F-33C572699FDE}");
		case 335:
			return createSpecialFolder("User Pinned", "shell:::{1F3427C8-5C10-4210-AA03-2EE45287D668}");
		case 336:
			return createSpecialFolder("This PC", "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}");
		case 337:
			// すべてのコントロール パネル項目
			return createSpecialFolder("All Control Panel Items", "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}");
		case 338:
			// プリンター
			return createSpecialFolder("Printers", "shell:::{2227A280-3AEA-1069-A2DE-08002B30309D}");
		case 339:
			// Win10から
			return createSpecialFolder("Local Pictures", "shell:::{24AD3AD4-A569-4530-98E1-AB02F9417AA8}");
		case 340:
			return createSpecialFolder("All Control Panel Items", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\0");
		case 341:
			return createSpecialFolder("Hardware and Sound", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\4");
		case 342:
			return createSpecialFolder("System and Security", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\10");
		case 343:
			return createSpecialFolder("All Control Panel Items", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\11");
		case 344:
			// Win8.1から
			return createSpecialFolder("Downloads", "shell:::{374DE290-123F-4565-9164-39C4925E467B}");
		case 345:
			// Win8.1から
			return createSpecialFolder("My Pictures", "shell:::{3ADD1653-EB32-4CB0-BBD7-DFA0ABB5ACCA}");
		case 346:
			// Win10から
			return createSpecialFolder("Local Music", "shell:::{3DFDF296-DBEC-4FB4-81D1-6A3438BCF4DE}");
		case 347:
			return createSpecialFolder("Applications", "shell:::{4234D49B-0245-4DF3-B780-3893943456E1}");
		case 348:
			// パブリック
			return createSpecialFolder("Public Folder", "shell:::{4336A54D-038B-4685-AB02-99BB52D3FB8B}");
		case 349:
			return createSpecialFolder("UsersFiles", "shell:::{59031A47-3F72-44A7-89C5-5595FE6B30EE}");
		case 350:
			// このデバイス
			return createSpecialFolder("This Device", "shell:::{5B934B42-522B-4C34-BBFE-37A3EF7B9C90}");
		case 351:
			// ごみ箱
			return createSpecialFolder("Recycle Bin", "shell:::{645FF040-5081-101B-9F08-00AA002F954E}");
		case 352:
			// プログラムと機能
			return createSpecialFolder("Programs and Features", "shell:::{7B81BE6A-CE2B-4676-A29E-EB907A5126C5}");
		case 353:
			// ネットワーク接続
			return createSpecialFolder("Network Connections", "shell:::{7007ACC7-3202-11D1-AAD2-00805FC1270E}");
		case 354:
			// プリンター
			return createSpecialFolder("Remote Printers", "shell:::{863AA9FD-42DF-457B-8E4D-0DE1B8015C60}");
		case 355:
			// インターネット
			return createSpecialFolder("Internet Folder", "shell:::{871C5380-42A0-1069-A2EA-08002B30309D}");
		case 356:
			// Win8.1のみ
			return createSpecialFolder("OneDrive", "shell:::{8E74D236-7F35-4720-B138-1FED0B85EA75}");
		case 357:
			// 検索結果
			return createSpecialFolder("CLSID_SearchHome", "shell:::{9343812E-1C37-4A49-A12E-4B2D810D956B}");
		case 358:
			// 同期センター
			return createSpecialFolder("Sync Center Folder", "shell:::{9C73F5E5-7AE7-4E32-A8E8-8D23B85255BF}");
		case 359:
			// ネットワーク接続
			return createSpecialFolder("Network Connections", "shell:::{992CFFA0-F557-101A-88EC-00DD010CCC48}");
		case 360:
			// Win8.1から
			return createSpecialFolder("My Video", "shell:::{A0953C92-50DC-43BF-BE83-3742FED03C9C}");
		case 361:
			// Win8.1から
			return createSpecialFolder("Personal", "shell:::{A8CDFF1C-4878-43BE-B5FD-F8091C1C60D0}");
		case 362:
			// Win8.1 UpdateからWin10 1511まで
			return createSpecialFolder("StartMenuAllPrograms", "shell:::{ADFA80E7-9769-4AD9-992C-55DC57E1008C}");
		case 363:
			// Win8.1から
			return createSpecialFolder("ThisPCDesktopFolder", "shell:::{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}");
		case 364:
			// ホームグループ
			return createSpecialFolder("Other Users Folder", "shell:::{B4FB3F98-C1EA-428D-A78A-D1F5659CBA93}");
		case 365:
			// フォント
			return createSpecialFolder("Microsoft Windows Font Folder", "shell:ControlPanelFolder\\::{BD84B380-8CA2-1069-AB1D-08000948F534}");
		case 366:
			// 管理ツール (All Users)
			return createSpecialFolder("Administrative Tools", "shell:::{D20EA4E1-3957-11D2-A40B-0C5020524153}");
		case 367:
			// Win10から
			return createSpecialFolder("Local Documents", "shell:::{D3162B92-9365-467A-956B-92703ACA08AF}");
		case 368:
			// インストールされた更新プログラム
			return createSpecialFolder("Installed Updates", "shell:::{D450A8A1-9568-45C7-9C0E-B4F9FB4537BD}");
		case 369:
			// ゲーム
			// Win10 1709までサポート
			return createSpecialFolder("Games Explorer", "shell:::{ED228FDF-9EA8-4870-83B1-96B02CFE0D52}");
		case 370:
			// ネットワーク
			return createSpecialFolder("Computers and Devices", "shell:::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}");
		case 371:
			// このデバイス (Win10 1703から)
			return createSpecialFolder("This Device", "shell:::{F8278C54-A712-415B-B593-B77A2BE0DDA9}");
		case 372:
			// Win10から
			return createSpecialFolder("Local Videos", "shell:::{F86FA3AB-70D2-4FC7-9C99-FCBF05467F3A}");
		
		case 373:
			// すべてのコントロール パネル項目
			return createSpecialFolder("Control Panel command object for Start menu and desktop", "shell:::{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}");
		case 374:
			// 既定のプログラム
			return createSpecialFolder("Default Programs command object for Start menu", "shell:::{E44E5D18-0652-4508-A4E2-8A090067BCB0}");
		
		// フォルダーとして使えないshellコマンド
		// Get-Clsid.ps1やGet-ShellCommand.ps1でヒットしないようにするためのもの
		case 375:
			return createSpecialFolder("shell:MAPIFolder", "shell:MAPIFolder", { category: "Unusable" });
		case 376:
			return createSpecialFolder("shell:RecordedTVLibrary", "shell:RecordedTVLibrary");
		
		case 377:
			return createSpecialFolder("", "shell:::{00020D75-0000-0000-C000-000000000046}");
		case 378:
			return createSpecialFolder("Desktop", "shell:::{00021400-0000-0000-C000-000000000046}");
		case 379:
			return createSpecialFolder("Shortcut", "shell:::{00021401-0000-0000-C000-000000000046}");
		case 380:
			// Win10 1507から1703まで
			return createSpecialFolder("", "shell:::{047EA9A0-93BB-415F-A1C3-D7AEB3DD5087}");
		case 381:
			return createSpecialFolder("Open With Context Menu Handler", "shell:::{09799AFB-AD67-11D1-ABCD-00C04FC30936}");
		case 382:
			return createSpecialFolder("Folder Shortcut", "shell:::{0AFACED1-E828-11D1-9187-B532F1E9575D}");
		case 383:
			return createSpecialFolder("", "shell:::{0C39A5CF-1A7A-40C8-BA74-8900E6DF5FCD}");
		case 384:
			return createSpecialFolder("", "shell:::{0D45D530-764B-11D0-A1CA-00AA00C16E65}");
		case 385:
			// Win8から
			return createSpecialFolder("Shell File System Folder", "shell:::{0E5AAE11-A475-4C5B-AB00-C66DE400274E}");
		case 386:
			return createSpecialFolder("Device Center Print Context Menu Extension", "shell:::{0E6DAA63-DD4E-47CE-BF9D-FDB72ECE4A0D}");
		case 387:
			return createSpecialFolder("IE History and Feeds Shell Data Source for Windows Search", "shell:::{11016101-E366-4D22-BC06-4ADA335C892B}");
		case 388:
			return createSpecialFolder("OpenMediaSharing", "shell:::{17FC1A80-140E-4290-A64F-4A29A951A867}");
		case 389:
			return createSpecialFolder("CLSID_DBFolderBoth", "shell:::{1BEF2128-2F96-4500-BA7C-098DC0049CB2}");
		case 390:
			return createSpecialFolder("CompatContextMenu Class", "shell:::{1D27F844-3A1F-4410-85AC-14651078412D}");
		case 391:
			// Win10 1903(?)以降では[OtherShellCommands]カテゴリになるので非表示に
			return createSpecialFolder("Windows Security", WIN10_1903 ? null : "shell:::{2559A1F2-21D7-11D4-BDAF-00C04F60B9F0}");
		case 392:
			return createSpecialFolder("Location Folder", "shell:::{267CF8A9-F4E3-41E6-95B1-AF881BE130FF}");
		case 393:
			return createSpecialFolder("Enhanced Storage Context Menu Handler Class", "shell:::{2854F705-3548-414C-A113-93E27C808C85}");
		case 394:
			return createSpecialFolder("System Restore", "shell:::{3F6BC534-DFA1-4AB4-AE54-EF25A74E0107}");
		case 395:
			return createSpecialFolder("Start Menu Folder", "shell:::{48E7CAAB-B918-4E58-A94D-505519C795DC}");
		case 396:
			return createSpecialFolder("IGD Property Page", "shell:::{4A1E5ACD-A108-4100-9E26-D2FAFA1BA486}");
		case 397:
			// Win10 1607まで
			return createSpecialFolder("LzhCompressedFolder2", "shell:::{4F289A46-2BBB-4AE8-9EDA-E5E034707A71}");
		case 398:
			// Win10から
			return createSpecialFolder("This PC", "shell:::{5E5F29CE-E0A8-49D3-AF32-7A7BDC173478}");
		case 399:
			return createSpecialFolder("", "shell:::{62AE1F9A-126A-11D0-A14B-0800361B1103}");
		case 400:
			return createSpecialFolder("Search Connector Folder", "shell:::{72B36E70-8700-42D6-A7F7-C9AB3323EE51}");
		case 401:
			return createSpecialFolder("CryptPKO Class", "shell:::{7444C717-39BF-11D1-8CD9-00C04FC29D45}");
		case 402:
			return createSpecialFolder("Temporary Internet Files", "shell:::{7BD29E00-76C1-11CF-9DD0-00A0C9034933}");
		case 403:
			return createSpecialFolder("Temporary Internet Files", "shell:::{7BD29E01-76C1-11CF-9DD0-00A0C9034933}");
		case 404:
			return createSpecialFolder(WIN10_1703 ? "" : "Briefcase", "shell:::{85BBD920-42A0-1069-A2E4-08002B30309D}");
		case 405:
			return createSpecialFolder("Shortcut", "shell:::{85CFCCAF-2D14-42B6-80B6-F40F65D016E7}");
		case 406:
			return createSpecialFolder("Mobile Broadband Profile Settings Editor", "shell:::{87630419-6216-4FF8-A1F0-143562D16D5C}");
		case 407:
			return createSpecialFolder("Compressed (zipped) Folder SendTo Target", "shell:::{888DCA60-FC0A-11CF-8F0F-00C04FD7D062}");
		case 408:
			return createSpecialFolder("ActiveX Cache Folder", "shell:::{88C6C381-2E85-11D0-94DE-444553540000}");
		case 409:
			return createSpecialFolder("Libraries delegate folder that appears in Users Files Folder", "shell:::{896664F7-12E1-490F-8782-C0835AFD98FC}");
		case 410:
			// Win10 1607まで
			return createSpecialFolder("Windows Search Service Media Center Namespace Extension Handler", "shell:::{98D99750-0B8A-4C59-9151-589053683D73}");
		case 411:
			return createSpecialFolder("MAPI Shell Context Menu", "shell:::{9D3C0751-A13F-46A6-B833-B46A43C30FE8}");
		case 412:
			return createSpecialFolder("Previous Versions", "shell:::{9DB7A13C-F208-4981-8353-73CC61AE2783}");
		case 413:
			return createSpecialFolder("Mail Service", "shell:::{9E56BE60-C50F-11CF-9A2C-00A0C90A90CE}");
		case 414:
			return createSpecialFolder("Desktop Shortcut", "shell:::{9E56BE61-C50F-11CF-9A2C-00A0C90A90CE}");
		case 415:
			return createSpecialFolder("DevicePairingFolder Initialization", "shell:::{AEE2420F-D50E-405C-8784-363C582BF45A}");
		case 416:
			return createSpecialFolder("CLSID_DBFolder", "shell:::{B2952B16-0E07-4E5A-B993-58C52CB94CAE}");
		case 417:
			return createSpecialFolder("Device Center Scan Context Menu Extension", "shell:::{B5A60A9E-A4C7-4A93-AC6E-0B76D1D87DC4}");
		case 418:
			return createSpecialFolder("DeviceCenter Initialization", "shell:::{C2B136E2-D50E-405C-8784-363C582BF43E}");
		case 419:
			// Win10 1507から1607まで
			return createSpecialFolder("", "shell:::{D9AC5E73-BB10-467B-B884-AA1E475C51F5}");
		case 420:
			return createSpecialFolder("delegate folder that appears in Users Files Folder", "shell:::{DFFACDC5-679F-4156-8947-C5C76BC0B67F}");
		case 421:
			return createSpecialFolder("CompressedFolder", "shell:::{E88DCCE0-B7B3-11D1-A9F0-00AA0060FA31}");
		case 422:
			return createSpecialFolder("MyDocs Drop Target", "shell:::{ECF03A32-103D-11D2-854D-006008059367}");
		case 423:
			return createSpecialFolder("Shell File System Folder", "shell:::{F3364BA0-65B9-11CE-A9BA-00AA004AE837}");
		case 424:
			// Win10 1607まで
			return createSpecialFolder("Sticky Notes Namespace Extension for Windows Desktop Search", "shell:::{F3F5824C-AD58-4728-AF59-A1EBE3392799}");
		case 425:
			return createSpecialFolder("Subscription Folder", "shell:::{F5175861-2688-11D0-9C5E-00AA00A45957}");
		case 426:
			return createSpecialFolder("Internet Shortcut", "shell:::{FBF23B40-E3F0-101B-8488-00AA003E56F8}");
		case 427:
			return createSpecialFolder("History", "shell:::{FF393560-C2A7-11CF-BFF4-444553540000}");
		case 428:
			index.current = DONE_ITERATION;
			
			return createSpecialFolder("Windows Photo Viewer Image Verbs", "shell:::{FFE2A43C-56B9-4BF5-9A79-CC6D4285608A}");
		
		case DONE_ITERATION:
			return null;
		
		default:
			throw new Error("indexが不正: " + index.current);
		}
	};
})();
