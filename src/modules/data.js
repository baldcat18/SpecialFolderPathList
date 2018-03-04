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
	
	var WIN10_1507_to_1511 = WIN10 && !WIN10_1607;
	
	var is64bit = State.Host.platform == 64;
	
	/** @type {ShellSpecialFolderConstants.ssfDESKTOPDIRECTORY} */
	var ssfDESKTOPDIRECTORY = 16;
	/** @type {ShellSpecialFolderConstants.ssfDRIVES} */
	var ssfDRIVES = 17;
	/** @type {ShellSpecialFolderConstants.ssfPROFILE} */
	var ssfPROFILE = 40;
	/** @type {SpecialFolderConst.TemporaryFolder} */
	var TemporaryFolder = 2;
	
	var USER_SHELL_FOLDERS_KEY =
		"HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders\\";
	var CURRENT_VERSION_KEY = "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\";
	
	var LIBRARIES_PATH = getRegValue(USER_SHELL_FOLDERS_KEY + "{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}",
			"%APPDATA%\\Microsoft\\Windows\\Libraries", true);
	var USER_ACCOUNT_PICTURES_PATH = "shell:Common AppData\\Microsoft\\User Account Pictures";
	
	var WINDOWS_APPS_TITLE = WIN10 ? "Windows アプリ" : "ストアアプリ";
	
	/** @type {SpecialFolderOption} */
	var defaultOption = {};
	/** @type {PropertyTypes.ptShellExecute} */
	var ptShellExecute = 1;
	/** @type {PropertyTypes.ptVerb} */
	var ptVerb = 2;
	
	var index = -1;
	var doneIteration = Infinity;
	
	/**
	 * @param {string} path
	 * @return {FolderItem}
	 */
	function getDirectoryFolderItem(path) {
		var items = shell.NameSpace(fso.GetParentFolderName(path)).Items();
		var name = fso.GetFileName(path);
		
		// @ts-ignore : 今の所 FolderItems.Item()に名前を渡しても機能する
		var item = items.Item(name);
		if (item) return item;
		
		/*
		for (var i = 0; i < items.Count; i++) {
			var item = items.Item(i);
			if (item.Name == name) return item;
		}
		*/
		
		if (!Setting.debug) return null;
		throw new Error(0x8000FFFF /* E_UNEXPECTED */, "フォルダーが見つからない: " + path);
	}
	
	
	/**
	 * @constructor
	 * @param {string} title
	 * @param {string | number} dir
	 * @param {SpecialFolderOption} [option]
	 */
	function SpecialFolderConstructor(title, dir, option) {
		this.title = title;
		this.dir = dir;
		
		try { this.folderItem = shell.NameSpace(dir).Self }
		catch (err) { this.folderItem = null; }
		
		if (!option) option = defaultOption;
		
		if (option.category) this.category = option.category;
		
		/** @type {string} */
		this.path = option.path || (this.folderItem ? this.folderItem.Path : "");
		if (this.path && this.path.charAt(0) == ":") this.path = "shell:" + this.path;
		
		this.isDirectory = fso.FolderExists(this.path);
		
		if (State.Host.type == "mshta") {
			this._folderItemForProperties = (option || defaultOption).folderItemForProperties;
			this._propertyTypes = option.propertyType ||
				(this.isDirectory && this._folderItemForProperties === undefined ? ptShellExecute : ptVerb);
			
			/** @type {FolderItemVerb} */
			this._properties = undefined;
		}
	};
	
	SpecialFolderConstructor.prototype.open = function() { this.folderItem.InvokeVerb(); };
	/** @param {string} [verb] */
	SpecialFolderConstructor.prototype.execCmd = function(verb) {
		if (this.isDirectory) shell.ShellExecute("cmd.exe", "/k pushd \"{0}\"".xFormat(this.path), null, verb);
		else writeError("ディレクトリではないのでコマンドプロンプトを実行できません。");
	};
	/** @param {string} [verb] */
	SpecialFolderConstructor.prototype.execExplorer = function(verb) {
		var path = typeof this.dir == "number" ? this.path : this.dir;
		shell.ShellExecute("explorer.exe", "\"{0}\"".xFormat(path), null, verb);
	};
	/** @param {string} [verb] */
	SpecialFolderConstructor.prototype.execPowershell = function(verb) {
		if (this.isDirectory) {
			var arg = "-NoExit -Command \"Push-Location -LiteralPath '{0}'\"".xFormat(this.path);
			shell.ShellExecute("powershell.exe", arg, null, verb);
		}
		else writeError("ディレクトリではないので PowerShell を実行できません。");
	};
	/** @return {boolean} */
	SpecialFolderConstructor.prototype.hasProperties = function() {
		if (this._propertyTypes == ptShellExecute) return true;
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
	SpecialFolderConstructor.prototype.showProperties = function() {
		if (this._propertyTypes == ptShellExecute) {
			var dir = /** @type {string} */ (this.dir);
			var path = dir.slice(0, 6) == "shell:" ? dir : "file:" + dir;
			shell.ShellExecute(path, "", "", "properties");
		}
		else if (this.hasProperties()) this._properties.DoIt();
		else writeError("プロパティを表示できません。");
	};
	SpecialFolderConstructor.prototype.getType = function() {
		return this.folderItem ? this.folderItem.Type : "使用不可";
	};
	
	global.SpecialFolder = SpecialFolderConstructor;
	
	SpecialFolder.getObject = function() {
		switch (++index) {
		case 0:
			// shell:Profile
			// shell:::{59031a47-3f72-44a7-89c5-5595fe6b30ee}
			// shell:ThisDeviceFolder / shell:::{f8278c54-a712-415b-b593-b77a2be0dda9} ([このデバイス]) (Win10 1703から)
			// %USERPROFILE%
			// %HOMEDRIVE%%HOMEPATH%
			return new SpecialFolder("個人用フォルダー", "shell:UsersFilesFolder", { category: "UserProfile", folderItemForProperties: shell.NameSpace(ssfPROFILE).Self });
		case 1:
			// Win10からサポート
			// shell:UsersFilesFolder\3D Objects
			// shell:MyComputerFolder\::{0DB7E03F-FC29-4DC6-9020-FF41B59E513A} (Win10 1709から)
			// Win10 1507から1703では3D Builderを起動した時に自動生成される
			return new SpecialFolder("3D オブジェクト", "shell:3D Objects");
		case 2:
			// shell:MyComputerFolder\::{B4BFCC3A-DB2C-424C-B029-7FE99A87C641} (Win8.1から)
			return new SpecialFolder("デスクトップ ディレクトリ", WIN81 ? "shell:ThisPCDesktopFolder" : ssfDESKTOPDIRECTORY, WIN81 ? null : { propertyType: ptVerb });
		case 3:
			// shell:Local Documents / shell:MyComputerFolder\::{d3162b92-9365-467a-956b-92703aca08af} (Win10から)
			// shell:::{450D8FBA-AD25-11D0-98A8-0800361B1103} ([マイ ドキュメント] (Win8.1から))
			// shell:MyComputerFolder\::{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0} (Win8.1から)
			return new SpecialFolder(WIN81 ? "ドキュメント" : "マイ ドキュメント", "shell:Personal");
		case 4:
			// shell:Local Downloads / shell:MyComputerFolder\::{088e3905-0323-4b02-9826-5d99428e115f} (Win10から)
			// shell:MyComputerFolder\::{374DE290-123F-4565-9164-39C4925E467B} (Win8.1から)
			return new SpecialFolder("ダウンロード", "shell:Downloads");
		
		case 5:
			// shell:Local Music / shell:MyComputerFolder\::{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de} (Win10から)
			// shell:MyComputerFolder\::{1CF1260C-4DD0-4ebb-811F-33C572699FDE} (Win8.1から)
			return new SpecialFolder(WIN81 ? "ミュージック" : "マイ ミュージック", "shell:My Music");
		case 6:
			// shell:My Music\Playlists
			// WMPやGroove ミュージックで再生リストを作成する時に自動生成される
			return new SpecialFolder("プレイリスト", "shell:Playlists");
		
		case 7:
			// shell:Local Pictures / shell:MyComputerFolder\::{24ad3ad4-a569-4530-98e1-ab02f9417aa8} (Win10から)
			// shell:MyComputerFolder\::{3ADD1653-EB32-4cb0-BBD7-DFA0ABB5ACCA} (Win8.1から)
			return new SpecialFolder(WIN81 ? "ピクチャ" : "マイ ピクチャ", "shell:My Pictures");
		case 8:
			// Win8.1からサポート
			// shell:My Pictures\Camera Roll
			// カメラアプリで写真や動画を撮影する時に自動生成される
			return new SpecialFolder("カメラ ロール", "shell:Camera Roll");
		case 9:
			// Win10からサポート
			// shell:My Pictures\Saved Pictures
			return new SpecialFolder("保存済みの写真", "shell:SavedPictures");
		case 10:
			// Win8からサポート
			// shell:My Pictures\Screenshots
			// Win＋PrtScrでスクリーンショットを保存する時に自動生成される
			return new SpecialFolder("スクリーンショット", "shell:Screenshots");
		case 11:
			// shell:My Pictures\Slide Shows
			// 手動でフォルダーを作成しても使用可
			// フォルダー名は大文字・小文字を一致させる必要あり(以下同)
			return new SpecialFolder("スライド ショー", "shell:PhotoAlbums");
		
		case 12:
			// shell:Local Videos / shell:MyComputerFolder\::{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a} (Win10から)
			// shell:MyComputerFolder\::{A0953C92-50DC-43bf-BE83-3742FED03C9C} (Win8.1から)
			return new SpecialFolder(WIN81 ? "ビデオ" : "マイ ビデオ", "shell:My Video");
		case 13:
			// Win10からサポート
			// shell:My Video\Captures
			return new SpecialFolder("キャプチャ", "shell:Captures");
		
		case 14:
			// Win10 1703からサポート
			// shell:UsersFilesFolder\AppMods
			return new SpecialFolder("アプリケーションの修正", "shell:AppMods");
		case 15:
			// shell:UsersFilesFolder\{56784854-C6CB-462B-8169-88E350ACB882}
			return new SpecialFolder("アドレス帳", "shell:Contacts");
		case 16:
			return new SpecialFolder("お気に入り", "shell:Favorites");
		case 17:
			// shell:::{323CA680-C24D-4099-B94D-446DD2D7249E} ([お気に入り])
			// shell:::{d34a6ca6-62c2-4c34-8a7c-14709c1ad938} ([Common Places FS Folder])
			// shell:UsersFilesFolder\{bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968}
			return new SpecialFolder("リンク" + (WIN10 ? "" : " (エクスプローラーのお気に入り)"), "shell:Links");
		case 18:
			// Win10からサポート
			// shell:UsersFilesFolder\Recorded Calls
			return new SpecialFolder("録音した通話", "shell:Recorded Calls");
		case 19:
			// shell:UsersFilesFolder\{4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4}
			return new SpecialFolder("保存したゲーム", "shell:SavedGames");
		case 20:
			// shell:UsersFilesFolder\{7d1d3a04-debb-4115-95cf-2f29da2920da}
			return new SpecialFolder("検索", "shell:Searches");
		
		// OneDriveカテゴリのフォルダーはすべてWin8.1からサポート
		case 21:
			// shell:UsersFilesFolder\OneDrive
			// Win8.1ではMicrosoftアカウントでサインインする時に自動生成される
			// shell:::{59031a47-3f72-44a7-89c5-5595fe6b30ee}\::{8E74D236-7F35-4720-B138-1FED0B85EA75} (Win8.1のみ)
			// shell:::{59031a47-3f72-44a7-89c5-5595fe6b30ee}\::{018D5C66-4533-4307-9B53-224DE2ED1FE6} (Win10から)
			// %OneDrive% (Win10 1607から)
			return new SpecialFolder("OneDrive", "shell:OneDrive", { category: "OneDrive" });
		case 22:
			// shell:OneDrive\Documents
			return new SpecialFolder("OneDrive のドキュメント", WIN10 ? "shell:OneDriveDocuments" : "shell:SkyDriveDocuments");
		case 23:
			// shell:OneDrive\Music
			return new SpecialFolder("OneDrive のミュージック", WIN10 ? "shell:OneDriveMusic" : "shell:SkyDriveMusic");
		case 24:
			// shell:OneDrive\Pictures
			return new SpecialFolder("OneDrive の画像", WIN10 ? "shell:OneDrivePictures" : "shell:SkyDrivePictures");
		case 25:
			// shell:OneDrive\Pictures\Camera Roll
			return new SpecialFolder("OneDrive のカメラ ロール", WIN10 ? "shell:OneDriveCameraRoll" : "shell:SkyDriveCameraRoll");
		
		case 26:
			// %APPDATA%
			return new SpecialFolder("Application Data", "shell:AppData", { category: "AppData" });
		case 27:
			return new SpecialFolder("Credentials", "shell:CredentialManager");
		case 28:
			return new SpecialFolder("Crypto", "shell:CryptoKeys");
		case 29:
			return new SpecialFolder("Protect", "shell:DpapiKeys");
		case 30:
			return new SpecialFolder("SystemCertificates", "shell:SystemCertificates");
		
		case 31:
			return new SpecialFolder("クイック起動", "shell:Quick Launch");
		case 32:
			// shell:::{1f3427c8-5c10-4210-aa03-2ee45287d668}
			return new SpecialFolder("User Pinned", "shell:User Pinned");
		case 33:
			return new SpecialFolder("ImplicitAppShortcuts", "shell:ImplicitAppShortcuts");
		
		case 34:
			// Win8からサポート
			return new SpecialFolder("アカウントの画像", "shell:AccountPictures");
		case 35:
			// Win8.1以降では[LocalAppData]カテゴリになるので非表示に
			return new SpecialFolder("Cookies", WIN81 ? null : "shell:Cookies");
		case 36:
			return new SpecialFolder("Network Shortcuts", "shell:NetHood");
		case 37:
			// shell:::{ed50fc29-b964-48a9-afb3-15ebb9b97f36} ([printhood delegate folder])
			return new SpecialFolder("Printer Shortcuts", "shell:PrintHood");
		case 38:
			return new SpecialFolder("最近使った項目", "shell:Recent");
		case 39:
			return new SpecialFolder("SendTo (送る)", "shell:SendTo");
		case 40:
			return new SpecialFolder("テンプレート", "shell:Templates");
		
		case 41:
			if (Setting.directoryOnly) index += 6;
			
			// shell:UsersLibrariesFolder
			// shell:::{031E4825-7B94-4dc3-B131-E946B44C8DD5}
			return new SpecialFolder("ライブラリ", "shell:Libraries", { category: "Libraries", path: LIBRARIES_PATH, folderItemForProperties: getDirectoryFolderItem(LIBRARIES_PATH) });
		case 42:
			// Win10からサポート
			// shell:Libraries\CameraRoll.library-ms
			// shell:Libraries\{2B20DF75-1EDA-4039-8097-38798227D5B7}
			return new SpecialFolder("カメラ ロール ライブラリ", "shell:CameraRollLibrary", { path: getRegValue(USER_SHELL_FOLDERS_KEY + "{2B20DF75-1EDA-4039-8097-38798227D5B7}", LIBRARIES_PATH + "\\CameraRoll.library-ms", true), propertyType: ptShellExecute });
		case 43:
			// shell:Libraries\{7b0db17d-9cd2-4a93-9733-46cc89022e7c}
			return new SpecialFolder("ドキュメント ライブラリ", "shell:DocumentsLibrary", { path: getRegValue(USER_SHELL_FOLDERS_KEY + "{7b0db17d-9cd2-4a93-9733-46cc89022e7c}", LIBRARIES_PATH + "\\Documents.library-ms", true), propertyType: ptShellExecute });
		case 44:
			// shell:Libraries\{2112AB0A-C86A-4ffe-A368-0DE96E47012E}
			return new SpecialFolder("ミュージック ライブラリ", "shell:MusicLibrary", { path: getRegValue(USER_SHELL_FOLDERS_KEY + "{2112AB0A-C86A-4ffe-A368-0DE96E47012E}", LIBRARIES_PATH + "\\Music.library-ms", true), propertyType: ptShellExecute });
		case 45:
			// shell:Libraries\{A990AE9F-A03B-4e80-94BC-9912D7504104}
			return new SpecialFolder("ピクチャ ライブラリ", "shell:PicturesLibrary", { path: getRegValue(USER_SHELL_FOLDERS_KEY + "{A990AE9F-A03B-4e80-94BC-9912D7504104}", LIBRARIES_PATH + "\\Pictures.library-ms", true), propertyType: ptShellExecute });
		case 46:
			// Win10からサポート
			// shell:Libraries\SavedPictures.library-ms
			// shell:Libraries\{E25B5812-BE88-4bd9-94B0-29233477B6C3}
			return new SpecialFolder("保存済みの写真 ライブラリ", "shell:SavedPicturesLibrary", { path: getRegValue(USER_SHELL_FOLDERS_KEY + "{E25B5812-BE88-4bd9-94B0-29233477B6C3}", LIBRARIES_PATH + "\\SavedPictures.library-ms", true), propertyType: ptShellExecute });
		case 47:
			// shell:::{031E4825-7B94-4dc3-B131-E946B44C8DD5}\{491E922F-5643-4af4-A7EB-4E7A138D8174}
			return new SpecialFolder("ビデオ ライブラリ", "shell:VideosLibrary", { path: getRegValue(USER_SHELL_FOLDERS_KEY + "{491E922F-5643-4af4-A7EB-4E7A138D8174}", LIBRARIES_PATH + "\\Videos.library-ms", true), propertyType: ptShellExecute });
		
		case 48:
			return new SpecialFolder("スタート メニュー", "shell:Start Menu", { category: "StartMenu" });
		case 49:
			return new SpecialFolder("プログラム", "shell:Programs");
		case 50:
			return new SpecialFolder(WIN10 ? "Windows 管理ツール" : "管理ツール", "shell:Administrative Tools");
		case 51:
			return new SpecialFolder("スタートアップ", "shell:Startup");
		
		case 52:
			// %LOCALAPPDATA%
			return new SpecialFolder("Local Application Data", "shell:Local AppData", { category: "LocalAppData" });
		case 53:
			return new SpecialFolder("Local Application Data Low", "shell:LocalAppDataLow");
		
		case 54:
			// Win10 1709からサポート
			// shell:Local AppData\Desktop
			return new SpecialFolder("AppDataDesktop", "shell:AppDataDesktop");
		case 55:
			// Win10からサポート
			// shell:Local AppData\DevelopmentFiles
			return new SpecialFolder("Development Files", "shell:Development Files");
		case 56:
			// Win10 1709からサポート
			// shell:Local AppData\Documents
			return new SpecialFolder("AppDataDocuments", "shell:AppDataDocuments");
		case 57:
			// Win10 1709からサポート
			// shell:Local AppData\Favorites
			return new SpecialFolder("AppDataFavorites", "shell:AppDataFavorites");
		case 58:
			// Win8からサポート
			return new SpecialFolder(WINDOWS_APPS_TITLE + "の設定", "shell:Local AppData\\Packages");
		case 59:
			// Win10 1709からサポート
			// shell:Local AppData\ProgramData
			return new SpecialFolder("AppDataProgramData", "shell:AppDataProgramData");
		case 60:
			// %TEMP%
			// %TMP%
			return new SpecialFolder("一時ファイル", fso.GetSpecialFolder(TemporaryFolder).Path);
		case 61:
			return new SpecialFolder("VirtualStore", "shell:Local AppData\\VirtualStore");
		
		case 62:
			// Win8からサポート
			return new SpecialFolder("アプリケーションのショートカット", "shell:Application Shortcuts");
		case 63:
			return new SpecialFolder("一時書き込みフォルダー", "shell:CD Burning");
		case 64:
			return new SpecialFolder("GameExplorer", "shell:GameTasks");
		case 65:
			return new SpecialFolder("履歴", "shell:History");
		case 66:
			return new SpecialFolder("インターネット一時ファイル", "shell:Cache");
		case 67:
			// Win8.1でこのカテゴリに移動
			return new SpecialFolder("Cookies", WIN81 ? "shell:Cookies" : null);
		case 68:
			return new SpecialFolder("Ringtones", "shell:Ringtones");
		case 69:
			// Win8からサポート
			// shell:Local AppData\Microsoft\Windows\RoamedTileImages
			return new SpecialFolder("Roamed Tile Images", "shell:Roamed Tile Images");
		case 70:
			// Win8からサポート
			return new SpecialFolder("Roaming Tiles", "shell:Roaming Tiles");
		case 71:
			// Win8からサポート
			return new SpecialFolder("WinX (クイックアクセスメニュー)", WIN8 ? "shell:Local AppData\\Microsoft\\Windows\\WinX" : null);
		
		case 72:
			// Win8.1からサポート
			// shell:Local AppData\Microsoft\Windows\ConnectedSearch\History
			return new SpecialFolder("検索履歴", "shell:SearchHistoryFolder");
		case 73:
			// Win8.1からサポート
			// shell:Local AppData\Microsoft\Windows\ConnectedSearch\Templates
			return new SpecialFolder("検索テンプレート", "shell:SearchTemplatesFolder");
		
		case 74:
			return new SpecialFolder("ガジェット", WIN8 ? "shell:Local AppData\\Microsoft\\Windows Sidebar\\Gadgets" :"shell:Gadgets");
		case 75:
			// shell:Local AppData\Microsoft\Windows Photo Gallery\Original Images
			// フォトギャラリーでファイルを編集する時に自動生成される
			return new SpecialFolder("Original Images", "shell:Original Images");
		
		case 76:
			// shell:Local AppData\Programs
			return new SpecialFolder("Program Files (Per User)", "shell:UserProgramFiles");
		case 77:
			// shell:Local AppData\Programs\Common
			return new SpecialFolder("Common Program Files (Per User)", "shell:UserProgramFilesCommon");
		
		case 78:
			// shell:::{4336a54d-038b-4685-ab02-99bb52d3fb8b}
			// shell:ThisDeviceFolder ([このデバイス]) (Win10 1507から1607まで)
			// shell:::{5b934b42-522b-4c34-bbfe-37a3ef7b9c90} ([このデバイス]) (Win10から)
			// %PUBLIC%
			return new SpecialFolder("パブリック", "shell:Public", { category: "Public" });
		case 79:
			// Win8からサポート
			// shell:Public\AccountPictures
			return new SpecialFolder("パブリック アカウントの画像", "shell:PublicAccountPictures");
		case 80:
			return new SpecialFolder("パブリック デスクトップ", "shell:Common Desktop");
		case 81:
			return new SpecialFolder("パブリックのドキュメント", "shell:Common Documents");
		case 82:
			return new SpecialFolder("パブリックのダウンロード", "shell:CommonDownloads");
		case 83:
			return new SpecialFolder("ライブラリ (All Users)", "shell:PublicLibraries");
		case 84:
			// Win8.1までサポート
			// メディアセンターが使えるエディションで使用可
			return new SpecialFolder("パブリックの録画一覧", getRegValue(CURRENT_VERSION_KEY + "Media Center\\Service\\Recording\\RecordPath"));
		case 85:
			return new SpecialFolder("パブリックのミュージック", "shell:CommonMusic");
		case 86:
			// shell:CommonMusic\Sample Music
			return new SpecialFolder("サンプル ミュージック", "shell:SampleMusic");
		case 87:
			// Win7までサポート
			// shell:CommonMusic\Sample Playlists
			return new SpecialFolder("サンプル プレイリスト", "shell:SamplePlaylists");
		case 88:
			return new SpecialFolder("パブリックのピクチャ", "shell:CommonPictures");
		case 89:
			// shell:CommonPictures\Sample Pictures
			return new SpecialFolder("サンプル ピクチャ", "shell:SamplePictures");
		case 90:
			return new SpecialFolder("パブリックのビデオ", "shell:CommonVideo");
		case 91:
			// shell:CommonVideo\Sample Videos
			return new SpecialFolder("サンプル ビデオ", "shell:SampleVideos");
		
		case 92:
			// %ALLUSERSPROFILE%
			// %ProgramData%
			return new SpecialFolder("ProgramData", "shell:Common AppData", { category: "ProgramData" });
		case 93:
			// %ALLUSERSPROFILE%\OEM Links
			return new SpecialFolder("OEM Links", "shell:OEM Links");
		
		case 94:
			return new SpecialFolder("Windows Search のインデックス", getRegValue("HKLM\\SOFTWARE\\Microsoft\\Windows Search\\DataDirectory", null, true));
		case 95:
			return new SpecialFolder("既定のアカウントの画像", WIN8 ? USER_ACCOUNT_PICTURES_PATH : USER_ACCOUNT_PICTURES_PATH + "\\Default Pictures");
		
		case 96:
			// Win8からサポート
			return new SpecialFolder(WINDOWS_APPS_TITLE + "のリポジトリ", getRegValue(CURRENT_VERSION_KEY + "Appx\\PackageRepositoryRoot"));
		case 97:
			return new SpecialFolder("Device Metadata Store", "shell:Device Metadata Store");
		case 98:
			return new SpecialFolder("GameExplorer (All Users)", "shell:PublicGameTasks");
		case 99:
			// Win10からサポート
			// shell:Common AppData\Microsoft\Windows\RetailDemo
			// 市販デモ モードで使用される
			return new SpecialFolder("RetailDemo", "shell:Retail Demo");
		case 100:
			return new SpecialFolder("Ringtones (All Users)", "shell:CommonRingtones");
		case 101:
			return new SpecialFolder("テンプレート (All Users)", "shell:Common Templates");
		
		case 102:
			return new SpecialFolder("スタート メニュー (All Users)", "shell:Common Start Menu", { category: "CommonStartMenu" });
		case 103:
			return new SpecialFolder("プログラム (All Users)", "shell:Common Programs");
		case 104:
			// shell:ControlPanelFolder\::{D20EA4E1-3957-11d2-A40B-0C5020524153}
			return new SpecialFolder(WIN10 ? "Windows 管理ツール (All Users)" : "管理ツール (All Users)", "shell:Common Administrative Tools");
		case 105:
			return new SpecialFolder("スタートアップ (All Users)", "shell:Common Startup");
		case 106:
			// Win10からサポート
			return new SpecialFolder("Start Menu Places", "shell:Common Start Menu Places");
		
		case 107:
			// %SystemRoot%
			// %windir%
			return new SpecialFolder("Windows ディレクトリ", "shell:Windows", { category: "Windows" });
		case 108:
			// shell:::{1D2680C9-0E2A-469d-B787-065558BC7D43} ([Fusion Cache]) (.NET3.5まで)
			// CLSIDを使ってアクセスするとエクスプローラーがクラッシュする
			return new SpecialFolder(".NET Framework Assemblies", "shell:Windows\\assembly");
		case 109:
			return new SpecialFolder("ActiveX Cache Folder", getRegValue(CURRENT_VERSION_KEY + "Internet Settings\\ActiveXCache"));
		case 110:
			// shell:ControlPanelFolder\::{BD84B380-8CA2-1069-AB1D-08000948F534}
			return new SpecialFolder("フォント", "shell:Fonts");
		case 111:
			return new SpecialFolder("既定のサウンド", getRegValue(CURRENT_VERSION_KEY + "MediaPathUnexpanded", null, true));
		case 112:
			return new SpecialFolder("Subscription Folder", "shell:Windows\\Offline Web Pages");
		case 113:
			// HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers\knownfolders\0\Windows Wallpapers\MergeFolders
			return new SpecialFolder("壁紙", "shell:Windows\\Web");
		
		case 114:
			return new SpecialFolder("テーマのリソース", "shell:ResourceDir");
		case 115:
			// shell:ResourceDir\xxxx (xxxxはロケールIDの16進数4桁 日本語では0411)
			return new SpecialFolder("テーマのローカライズ リソース", "shell:LocalizedResourcesDir");
		
		case 116:
			return new SpecialFolder("システム ディレクトリ", is64bit ? "shell:System" : "shell:SystemX86");
		case 117:
			return is64bit ?
				new SpecialFolder("システム ディレクトリ (32 ビット)", "shell:SystemX86") :
				new SpecialFolder("システム ディレクトリ (64 ビット)", getSysNativePath());
		
		case 118:
			return new SpecialFolder("ユーザー", "shell:UserProfiles", { category: "Users" });
		case 119:
			return new SpecialFolder("既定のプロファイル", getRegValue("HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList\\Default", null, true));
		
		case 120:
			// shell:ProgramFilesX64 (64ビットアプリのみ)
			// %ProgramFiles%
			return new SpecialFolder("Program Files", "shell:ProgramFiles", { category: "ProgramFiles" });
		case 121:
			return is64bit ?
				new SpecialFolder("Program Files (32 ビット)", "shell:ProgramFilesX86") :
				new SpecialFolder("Program Files (64 ビット)", getRegValue(CURRENT_VERSION_KEY + "ProgramW6432Dir"));
		case 122:
			// shell:ProgramFilesCommonX64 (64ビットアプリのみ)
			// %CommonProgramFiles%
			return new SpecialFolder("Common Program Files", "shell:ProgramFilesCommon");
		case 123:
			return is64bit ?
				new SpecialFolder("Common Program Files (32 ビット)", "shell:ProgramFilesCommonX86") :
				new SpecialFolder("Common Program Files (64 ビット)", getRegValue(CURRENT_VERSION_KEY + "CommonW6432Dir"));
		case 124:
			// Win8からサポート
			return new SpecialFolder(WINDOWS_APPS_TITLE, getRegValue(CURRENT_VERSION_KEY + "Appx\\PackageRoot"));
		case 125:
			return new SpecialFolder("既定のガジェット", WIN8 ? "shell:ProgramFiles\\Windows Sidebar\\Gadgets" : "shell:Default Gadgets");
		case 126:
			if (Setting.directoryOnly) index = doneIteration;
			
			return new SpecialFolder("ガジェット (All Users)", "shell:ProgramFiles\\Windows Sidebar\\Shared Gadgets");
		
		case 127:
			return new SpecialFolder("デスクトップ", "shell:Desktop", { category: "Desktop / " + (WIN81 ? "ThisPC" : "Computer") });
		case 128:
			// shell:MyComputerFolderはWin10 1507/1511だとなぜかデスクトップになってしまう
			return new SpecialFolder(WIN81 ? "PC" : "コンピューター", !WIN10_1507_to_1511 ? "shell:MyComputerFolder" : ssfDRIVES);
		case 129:
			return new SpecialFolder(WIN10 ? "最近使ったフォルダー" : "最近表示した場所", "shell:::{22877a6d-37a1-461a-91b0-dbda5aaebc99}");
		case 130:
			// Win10からサポート
			// shell:::{4564b25e-30cd-4787-82ba-39e73a750b14} ([Recent Items Instance Folder])
			return new SpecialFolder("最近使用したファイル", "shell:::{3134ef9c-6b18-4996-ad04-ed5912e00eb5}");
		case 131:
			return new SpecialFolder("ポータブル メディア デバイス", "shell:::{35786D3C-B075-49b9-88DD-029876E11C01}");
		case 132:
			// Win10からサポート
			return new SpecialFolder("よく使用するフォルダー", "shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}");
		case 133:
			return new SpecialFolder("ごみ箱", "shell:RecycleBinFolder");
		case 134:
			// Win10からサポート
			return new SpecialFolder("クイック アクセス", "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}");
		case 135:
			// Win8からサポート
			// Win8/8.1では[PC]と同じなので非表示に
			return new SpecialFolder("Removable Storage Devices", WIN10 ? "shell:::{a6482830-08eb-41e2-84c1-73920c2badb9}": null);
		case 136:
			return new SpecialFolder("ホームグループ", "shell:HomeGroupFolder");
		case 137:
			return new SpecialFolder("ネットワーク", "shell:NetworkPlacesFolder");
		case 138:
			// Win10からサポート
			return new SpecialFolder("Removable Drives", "shell:::{F5FB2C77-0E2F-4A16-A381-3E560C68BC83}");
		
		case 139:
			return new SpecialFolder("コントロール パネル (カテゴリ表示)", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}", { category: "ControlPanel" });
		case 140:
			return new SpecialFolder("デスクトップのカスタマイズ", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\1");
		case 141:
			// shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\4
			return new SpecialFolder("ハードウェアとサウンド", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\2");
		case 142:
			return new SpecialFolder("ネットワークとインターネット", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\3");
		case 143:
			// shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\10
			return new SpecialFolder("システムとセキュリティ", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\5");
		case 144:
			return new SpecialFolder("時計、言語、および地域", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\6");
		case 145:
			return new SpecialFolder("コンピューターの簡単操作", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\7");
		case 146:
			return new SpecialFolder("プログラム", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\8");
		case 147:
			return new SpecialFolder(WIN10 ? "ユーザー アカウント" : WIN8 ? "ユーザー アカウントとファミリー セーフティ" : "ユーザー アカウントと家族のための安全設定", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\9");
		
		case 148:
			// shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}
			// shell:::{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}
			// shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\11
			return new SpecialFolder("すべてのコントロール パネル項目", "shell:ControlPanelFolder");
		
		// コントロールパネル内の項目はCLSIDだけを指定してもアクセス可能
		// 例えば[電源オプション]なら shell:::{025A5937-A6BE-4686-A844-36FE4BEC8B6D}
		// ただしその場合はアドレスバーからコントロールパネルに移動できない
		case 149:
			// Win7/8のみサポート
			return new SpecialFolder("既定の位置", "shell:ControlPanelFolder\\::{00C6D95F-329C-409a-81D7-C46C66EA7F33}");
		case 150:
			// Win7/8のみサポート
			return new SpecialFolder("生体認証デバイス", "shell:ControlPanelFolder\\::{0142e4d0-fb7a-11dc-ba4a-000ffe7ab428}");
		case 151:
			return new SpecialFolder("電源オプション", "shell:ControlPanelFolder\\::{025A5937-A6BE-4686-A844-36FE4BEC8B6D}");
		case 152:
			return new SpecialFolder("資格情報マネージャー", "shell:ControlPanelFolder\\::{1206F5F1-0569-412C-8FEC-3204630DFB70}");
		case 153:
			return new SpecialFolder("プログラムの取得", "shell:AddNewProgramsFolder");
		case 154:
			// shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{E44E5D18-0652-4508-A4E2-8A090067BCB0}
			return new SpecialFolder("既定のプログラム", "shell:ControlPanelFolder\\::{17cd9488-1228-4b2f-88ce-4298e93e0966}");
		case 155:
			return new SpecialFolder("RemoteApp とデスクトップ接続", "shell:ControlPanelFolder\\::{241D7C96-F8BF-4F85-B01F-E2B043341A4B}");
		case 156:
			// Win8.1までサポート
			return new SpecialFolder("Windows Update", "shell:ControlPanelFolder\\::{36eef7db-88ad-4e81-ad49-0e313f0c35f8}");
		case 157:
			return new SpecialFolder(WIN10_1709 ? "Windows Defender ファイアウォール" : "Windows ファイアウォール", "shell:ControlPanelFolder\\::{4026492F-2F69-46B8-B9BF-5654FC07E423}");
		case 158:
			return new SpecialFolder("音声認識", "shell:ControlPanelFolder\\::{58E3C745-D971-4081-9034-86E34B30836A}");
		case 159:
			return new SpecialFolder("ユーザー アカウント", "shell:ControlPanelFolder\\::{60632754-c523-4b62-b45c-4172da012619}");
		case 160:
			// Win10 1709までサポート (?)
			return new SpecialFolder("ホームグループ", "shell:ControlPanelFolder\\::{67CA7650-96E6-4FDD-BB43-A8E774F73A57}");
		case 161:
			// Win8までサポート
			return new SpecialFolder("パフォーマンスの情報とツール", "shell:ControlPanelFolder\\::{78F3955E-3B90-4184-BD14-5397C15F1EFC}");
		case 162:
			return new SpecialFolder("ネットワークと共有センター", "shell:ControlPanelFolder\\::{8E908FC9-BECC-40f6-915B-F4CA0E70D03D}");
		case 163:
			return new SpecialFolder(WIN8 ? "ファミリー セーフティ" : "保護者による制限", "shell:ControlPanelFolder\\::{96AE8D84-A250-4520-95A5-A47A7E3C548B}");
		case 164:
			return new SpecialFolder("自動再生", "shell:ControlPanelFolder\\::{9C60DE1E-E5FC-40f4-A487-460851A8D915}");
		case 165:
			return new SpecialFolder("回復", "shell:ControlPanelFolder\\::{9FE63AFD-59CF-4419-9775-ABCC3849F861}");
		case 166:
			return new SpecialFolder("デバイスとプリンター", "shell:ControlPanelFolder\\::{A8A91A66-3A7D-4424-8D24-04E180695C7A}");
		case 167:
			// Win8.1以外でサポート
			return new SpecialFolder(WIN10 ? "バックアップと復元 (Windows 7)" : "バックアップと復元", "shell:ControlPanelFolder\\::{B98A2BEA-7D42-4558-8BD1-832F41BAC6FD}");
		case 168:
			return new SpecialFolder("システム", "shell:ControlPanelFolder\\::{BB06C0E4-D293-4f75-8A90-CB05B6477EEE}");
		case 169:
			return new SpecialFolder(WIN10 ? "セキュリティとメンテナンス" : "アクション センター", "shell:ControlPanelFolder\\::{BB64F8A7-BEE7-4E1A-AB8D-7D8273F7FDB6}");
		case 170:
			// shell:Fonts
			return new SpecialFolder("フォント", "shell:ControlPanelFolder\\::{BD84B380-8CA2-1069-AB1D-08000948F534}", { path: "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\0\\::{BD84B380-8CA2-1069-AB1D-08000948F534}" });
		case 171:
			return new SpecialFolder("言語", "shell:ControlPanelFolder\\::{BF782CC9-5A52-4A17-806C-2A894FFEEAC5}");
		case 172:
			// Win10 1607までサポート
			return new SpecialFolder("ディスプレイ", "shell:ControlPanelFolder\\::{C555438B-3C23-4769-A71F-B6D3D9B6053A}");
		case 173:
			return new SpecialFolder("トラブルシューティング", "shell:ControlPanelFolder\\::{C58C4893-3BE0-4B45-ABB5-A63E4B8C8651}");
		case 174:
			// Win7までサポート
			return new SpecialFolder("はじめに", "shell:ControlPanelFolder\\::{CB1B7F8C-C50A-4176-B604-9E24DEE8D4D1}");
		case 175:
			// shell:Common Administrative Tools
			return new SpecialFolder("管理ツール", "shell:ControlPanelFolder\\::{D20EA4E1-3957-11d2-A40B-0C5020524153}", { path: "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\0\\::{D20EA4E1-3957-11d2-A40B-0C5020524153}" });
		case 176:
			return new SpecialFolder("コンピューターの簡単操作センター", "shell:ControlPanelFolder\\::{D555645E-D4F8-4c29-A827-D93C859C4F2A}");
		case 177:
			// Enterprise/Ultimateで使用可
			// Win8からはProでも使用可
			// Win8.1からはCore/Homeでも使用可
			return new SpecialFolder(shell.GetSystemInformation("IsOS_Personal") ? "デバイスの暗号化" : "BitLocker ドライブ暗号化", "shell:ControlPanelFolder\\::{D9EF8727-CAC2-4e60-809E-86F80A666C91}");
		case 178:
			// Win7までサポート
			return new SpecialFolder("ネットワーク マップ", "shell:ControlPanelFolder\\::{E7DE9B1A-7533-4556-9484-B26FB486475E}");
		case 179:
			// Win8までサポート
			return new SpecialFolder("Windows SideShow", "shell:ControlPanelFolder\\::{E95A4861-D57A-4be1-AD0F-35267E261739}");
		case 180:
			// Win8.1までサポート
			return new SpecialFolder(WIN8 ? "位置情報の設定" : "位置センサーとその他のセンサー", "shell:ControlPanelFolder\\::{E9950154-C418-419e-A90A-20C5287AE24B}");
		case 181:
			// Win8.1からサポート
			// Win7ではKB2891638をインストールすれば使用可
			return new SpecialFolder("ワーク フォルダー", "shell:ControlPanelFolder\\::{ECDB0924-4208-451E-8EE0-373C0956DE16}");
		case 182:
			return new SpecialFolder(WIN10_1607 ? "個人用設定" : "個人設定", "shell:ControlPanelFolder\\::{ED834ED6-4B5A-4bfe-8F11-A626DCB6A921}");
		case 183:
			// Win8からサポート
			return new SpecialFolder("ファイル履歴", "shell:ControlPanelFolder\\::{F6B6E965-E9B2-444B-9286-10C9152EDBC5}");
		case 184:
			// Win8からサポート
			return new SpecialFolder("記憶域", "shell:ControlPanelFolder\\::{F942C606-0914-47AB-BE56-1321B8035096}");
		
		case 185:
			return new SpecialFolder("プログラムと機能", "shell:ChangeRemoveProgramsFolder");
		case 186:
			return new SpecialFolder("インストールされた更新プログラム", "shell:AppUpdatesFolder");
		
		case 187:
			return new SpecialFolder("同期センター", "shell:SyncCenterFolder");
		case 188:
			// shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{2E9E59C0-B437-4981-A647-9C34B9B90891} ([Sync Setup Folder])
			return new SpecialFolder("同期のセットアップ", "shell:SyncSetupFolder");
		case 189:
			return new SpecialFolder("同期の競合", "shell:ConflictFolder");
		case 190:
			return new SpecialFolder("同期結果", "shell:SyncResultsFolder");
		
		case 191:
			return new SpecialFolder("通知領域アイコン", "shell:" + (WIN10 ? "::{21EC2020-3AEA-1069-A2DD-08002B30309D}" : "ControlPanelFolder") + "\\::{05d7b0f4-2121-4eff-bf6b-ed3f69b894d9}");
		case 192:
			// Win8.1以降ではフォルダーを開けないので非表示に
			return new SpecialFolder("ワイヤレス ネットワークの管理", WIN81 ? null : "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\\::{1FA9085F-25A2-489B-85D4-86326EEDCD87}");
		case 193:
			// shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{863AA9FD-42DF-457B-8E4D-0DE1B8015C60}
			return new SpecialFolder("プリンター", "shell:PrintersFolder");
		case 194:
			return new SpecialFolder("Bluetooth デバイス", "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\\::{28803F59-3A75-4058-995F-4EE5503B023C}");
		case 195:
			// shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{992CFFA0-F557-101A-88EC-00DD010CCC48}
			return new SpecialFolder("ネットワーク接続", "shell:ConnectionsFolder");
		case 196:
			return new SpecialFolder("フォント設定", "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\\::{93412589-74D4-4E4E-AD0E-E0CB621440FD}");
		case 197:
			return new SpecialFolder("すべてのタスク (God Mode)", "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\\::{ED7BA470-8E54-465E-825C-99712043E01C}");
		
		case 198:
			// Win10 1703までサポート
			return new SpecialFolder("リモート ファイル ブラウザー", "shell:::{0907616E-F5E6-48D8-9D61-A91C3D28106D}", { category: "OtherFolders" });
		case 199:
			return new SpecialFolder("キャビネット ファイル ビューアー", "shell:::{0CD7A5C0-9F37-11CE-AE65-08002B2E1262}");
		case 200:
			return new SpecialFolder("ネットワーク (WORKGROUP)", "shell:::{208D2C60-3AEA-1069-A2D7-08002B30309D}");
		case 201:
			// Win8からサポート
			return new SpecialFolder("メディア サーバー", "shell:::{289AF617-1CC3-42A6-926C-E6A863F0E3BA}");
		case 202:
			return new SpecialFolder("Results Folder", "shell:::{2965e715-eb66-4719-b53f-1672673bbefa}");
		case 203:
			// Explorer Browser Results Folder
			return new SpecialFolder("", "shell:::{418c8b64-5463-461d-88e0-75e2afa3c6fa},");
		case 204:
			// Win8からサポート
			return new SpecialFolder("Applications (すべてのアプリ)", "shell:AppsFolder");
		case 205:
			return new SpecialFolder("Command Folder", "shell:::{437ff9c0-a07f-4fa0-af80-84b6c6440a16}");
		case 206:
			return new SpecialFolder("ホームグループ", "shell:::{6785BFAC-9D2D-4be5-B7E2-59937E8FB80A}");
		case 207:
			// Win8.1以降ではフォルダーを開けないので非表示に
			return new SpecialFolder("Programs Folder (すべてのプログラム)", WIN81 ? null : "shell:::{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}");
		case 208:
			// Win8.1以降ではフォルダーを開けないので非表示に
			return new SpecialFolder("Programs Folder and Fast Items (すべてのプログラム (ファイルを先に表示))", WIN81 ? null : "shell:::{865e5e76-ad83-4dca-a109-50dc2113ce9a}");
		case 209:
			// search:
			// search-ms:
			return new SpecialFolder("検索結果", "shell:SearchHomeFolder");
		case 210:
			// Win8.1からWin10 1511までサポート
			return new SpecialFolder("StartMenuAllPrograms", "shell:StartMenuAllPrograms");
		case 211:
			// 企業向けエディションで使用可
			return new SpecialFolder("オフライン ファイル フォルダー", "shell:::{AFDB1F70-2A4C-11d2-9039-00C04F8EEB3E}");
		case 212:
			return new SpecialFolder("delegate folder that appears in Computer", "shell:::{b155bdf8-02f0-451e-9a26-ae317cfd7779}");
		case 213:
			return new SpecialFolder("AppSuggestedLocations", "shell:::{c57a6066-66a3-4d91-9eb9-41532179f0a5}");
		case 214:
			// Win10 1709までサポート (?)
			return new SpecialFolder("ゲーム", "shell:Games");
		case 215:
			if (!Setting.debug) index = doneIteration;
			
			return new SpecialFolder("Previous Versions Results Folder", "shell:::{f8c2ab3b-17bc-41da-9758-339d7dbf2d88}");
		
		// 通常とは違う名前がエクスプローラーのタイトルバーに表示されるフォルダー
		case 216:
			// Win10から
			// shell:::{5b934b42-522b-4c34-bbfe-37a3ef7b9c90} (Win10 1507から1607まで)
			// shell:::{f8278c54-a712-415b-b593-b77a2be0dda9} (Win10 1703から)
			return new SpecialFolder("このデバイス ({0})".xFormat(WIN10_1703 ? "個人用フォルダー" : "パブリック"), "shell:ThisDeviceFolder", { category: "AnotherName", propertyType: ptVerb });
		case 217:
			// Win8までだと別名にならないので非表示に
			return new SpecialFolder("マイ ドキュメント (ドキュメント)", WIN81 ? "shell:::{450D8FBA-AD25-11D0-98A8-0800361B1103}" : null);
		case 218:
			return new SpecialFolder("お気に入り (リンク)", "shell:::{323CA680-C24D-4099-B94D-446DD2D7249E}");
		case 219:
			return new SpecialFolder("Common Places FS Folder (リンク)", "shell:::{d34a6ca6-62c2-4c34-8a7c-14709c1ad938}", { propertyType: ptVerb });
		case 220:
			return new SpecialFolder("printhood delegate folder (Printer Shortcuts)", "shell:::{ed50fc29-b964-48a9-afb3-15ebb9b97f36}");
		case 221:
			// .NET3.5まで
			// CLSIDを使ってアクセスするとエクスプローラーがクラッシュする
			return new SpecialFolder("Fusion Cache (.NET Framework Assemblies)", "shell:::{1D2680C9-0E2A-469d-B787-065558BC7D43}");
		case 222:
			// Win10から
			return new SpecialFolder("Recent Items Instance Folder (最近使用したファイル)", "shell:::{4564b25e-30cd-4787-82ba-39e73a750b14}");
		case 223:
			return new SpecialFolder("Sync Setup Folder (同期のセットアップ)", "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\\::{2E9E59C0-B437-4981-A647-9C34B9B90891}");

		// エクスプローラーで開けないフォルダー
		case 224:
			// 検索
			return new SpecialFolder("検索", "shell:::{04731B67-D933-450a-90E6-4ACD2E9408FE}", { category: "CantOpen" });
		case 225:
			// Win8.1でこのカテゴリに移動
			return new SpecialFolder("ワイヤレス ネットワークの管理", WIN81 ? "shell:::{1FA9085F-25A2-489B-85D4-86326EEDCD87}" : null);
		case 226:
			return new SpecialFolder("Sync Center Conflict Folder", "shell:::{289978AC-A101-4341-A817-21EBA7FD046D}");
		case 227:
			return new SpecialFolder("LayoutFolder", "shell:::{328B0346-7EAF-4BBE-A479-7CB88A095F5B}");
		case 228:
			return new SpecialFolder("Explorer Browser Results Folder", "shell:::{418c8b64-5463-461d-88e0-75e2afa3c6fa}");
		case 229:
			// Win8.1から
			return new SpecialFolder("すべての設定", "shell:::{5ED4F38C-D3FF-4D61-B506-6820320AEBFE}");
		case 230:
			return new SpecialFolder("Microsoft FTP Folder", "shell:::{63da6ec0-2e98-11cf-8d82-444553540000}");
		case 231:
			// Win8から
			return new SpecialFolder("CLSID_AppInstanceFolder", "shell:::{64693913-1c21-4f30-a98f-4e52906d3b56}");
		case 232:
			return new SpecialFolder("Sync Results Folder", "shell:::{71D99464-3B6B-475C-B241-E15883207529}");
		case 233:
			// Win8.1でこのカテゴリに移動
			// Win10 1511まで
			return new SpecialFolder("Programs Folder", WIN81 ? "shell:::{7be9d83c-a729-4d97-b5a7-1b7313c39e0a}" : null);
		case 234:
			// Win8.1でこのカテゴリに移動
			// Win10 1511まで
			return new SpecialFolder("Programs Folder and Fast Items", WIN81 ? "shell:::{865e5e76-ad83-4dca-a109-50dc2113ce9a}" : null);
		case 235:
			// Win10でこのカテゴリに移動
			return new SpecialFolder("インターネット", WIN10 ? "shell:InternetFolder" : null);
		case 236:
			return new SpecialFolder("File Backup Index", "shell:::{877ca5ac-cb41-4842-9c69-9136e42d47e2}");
		case 237:
			return new SpecialFolder("Microsoft Office Outlook", "shell:::{89D83576-6BD1-4c86-9454-BEB04E94C819}");
		case 238:
			return new SpecialFolder("DXP", "shell:::{8FD8B88D-30E1-4F25-AC2B-553D3D65F0EA}");
		case 239:
			return new SpecialFolder("Enhanced Storage Data Source", "shell:::{9113A02D-00A3-46B9-BC5F-9C04DADDD5D7}");
		case 240:
			// Win8から
			return new SpecialFolder("CLSID_StartMenuLauncherProviderFolder", "shell:::{98F275B4-4FFF-11E0-89E2-7B86DFD72085}");
		case 241:
			return new SpecialFolder("IE RSS Feeds Folder", "shell:::{9a096bb5-9dc3-4d1c-8526-c3cbf991ea4e}");
		case 242:
			// Win8から
			return new SpecialFolder("CLSID_StartMenuCommandingProviderFolder", "shell:::{a00ee528-ebd9-48b8-944a-8942113d46ac}");
		case 243:
			return new SpecialFolder("Previous Versions Results Delegate Folder", "shell:::{a3c3d402-e56c-4033-95f7-4885e80b0111}");
		case 244:
			return new SpecialFolder("Library Folder", "shell:::{a5a3563a-5755-4a6f-854e-afa3230b199f}");
		case 245:
			// Win8からサポート
			return new SpecialFolder("ホームグループ内の現在のユーザー", "shell:HomeGroupCurrentUserFolder");
		case 246:
			return new SpecialFolder("Sync Results Delegate Folder", "shell:::{BC48B32F-5910-47F5-8570-5074A8A5636A}");
		case 247:
			return new SpecialFolder("オフライン ファイル", "shell:::{BD7A2E7B-21CB-41b2-A086-B309680C6B7E}");
		case 248:
			// Win8から
			return new SpecialFolder("DLNA Content Directory Data Source", "shell:::{D2035EDF-75CB-4EF1-95A7-410D9EE17170}");
		case 249:
			return new SpecialFolder("CLSID_StartMenuProviderFolder", "shell:::{daf95313-e44d-46af-be1b-cbacea2c3065}");
		case 250:
			return new SpecialFolder("CLSID_StartMenuPathCompleteProviderFolder", "shell:::{e345f35f-9397-435c-8f95-4e922c26259e}");
		case 251:
			return new SpecialFolder("Sync Center Conflict Delegate Folder", "shell:::{E413D040-6788-4C22-957E-175D1C513A34}");
		case 252:
			return new SpecialFolder("Shell DocObject Viewer", "shell:::{E7E4BC40-E76A-11CE-A9BB-00AA004AE837}");
		case 253:
			// Win8から
			return new SpecialFolder("StreamBackedFolder", "shell:::{EDC978D6-4D53-4b2f-A265-5805674BE568}");
		case 254:
			return new SpecialFolder("Sync Setup Delegate Folder", "shell:::{F1390A9A-A3F4-4E5D-9C5F-98F3BD8D935C}");
		case 255:
			return new SpecialFolder("オフライン ファイル", "shell:CSCFolder");
		
		case 256:
			if (true) ++index;
			else {
				// Win8からサポート
				// ファイル履歴を有効にすると利用可
				// スクリプトのホストプログラムのプロセスが残り続ける
				// @ts-ignore
				return new SpecialFolder("FileHistoryDataSource", "shell:::{2F6CE85C-F9EE-43CA-90C7-8A9BD53A2467}");
			}
		
		// フォルダー以外のshellコマンド
		case 257:
			return new SpecialFolder(WIN8 ? "タスク バーとナビゲーション" : "タスク バーと[スタート]メニュー", "shell:::{0DF44EAA-FF21-4412-828E-260A8728E7F1}", { category: "OtherShellCommands" });
		case 258:
			// Win10 1511まで
			return new SpecialFolder(WIN8 ? "検索 - ファイル" : "検索", "shell:::{2559a1f0-21d7-11d4-bdaf-00c04f60b9f0}");
		case 259:
			// Win8.1まで
			return new SpecialFolder("ヘルプとサポート", "shell:::{2559a1f1-21d7-11d4-bdaf-00c04f60b9f0}");
		case 260:
			return new SpecialFolder("ファイル名を指定して実行", "shell:::{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}");
		case 261:
			return new SpecialFolder("電子メール", "shell:::{2559a1f5-21d7-11d4-bdaf-00c04f60b9f0}");
		case 262:
			return new SpecialFolder("プログラムのアクセスとコンピューターの既定の設定", "shell:::{2559a1f7-21d7-11d4-bdaf-00c04f60b9f0}");
		case 263:
			// Win8から
			return new SpecialFolder(WIN10 ? "Cortana" : "検索", "shell:::{2559a1f8-21d7-11d4-bdaf-00c04f60b9f0}");
		case 264:
			// Win+Dと同じ
			return new SpecialFolder("デスクトップの表示", "shell:::{3080F90D-D7AD-11D9-BD98-0000947B0257}");
		case 265:
			// Win7ではCtrl+Win+Tab、Win8/8.1ではCtrl+Alt+Tab、Win10 1607以降ではWin+Tabと同じ (Win10 1507/1511では使用不可)
			return new SpecialFolder("ウィンドウを切り替える", "shell:::{3080F90E-D7AD-11D9-BD98-0000947B0257}");
		case 266:
			// Win7まで
			return new SpecialFolder("ガジェット ギャラリー", "shell:::{37efd44d-ef8d-41b1-940d-96973a50e9e0}");
		case 267:
			return new SpecialFolder("接続先", "shell:::{38A98528-6CBF-4CA9-8DC0-B1E1D10F7B1B}");
		case 268:
			return new SpecialFolder("電話とモデム", "shell:::{40419485-C444-4567-851A-2DD7BFA1684D}");
		case 269:
			// Win8.1から
			return new SpecialFolder("新しいウィンドウで開く", "shell:::{52205fd8-5dfb-447d-801a-d0b52f2e83e1}");
		case 270:
			return new SpecialFolder("Windows モビリティ センター", "shell:::{5ea4f148-308c-46d7-98a9-49041b1dd468}");
		case 271:
			return new SpecialFolder(WIN8 ? "地域" : "地域と言語", "shell:::{62D8ED13-C9D0-4CE8-A914-47DD628FB1B0}");
		case 272:
			return new SpecialFolder("Windows の機能", "shell:::{67718415-c450-4f3c-bf8a-b487642dc39b}");
		case 273:
			return new SpecialFolder("マウス", "shell:::{6C8EEC18-8D75-41B2-A177-8831D59D2D50}");
		case 274:
			return new SpecialFolder(WIN10 ? "エクスプローラーのオプション" : "フォルダー オプション", "shell:::{6DFD7C5C-2451-11d3-A299-00C04F8EF6AF}");
		case 275:
			return new SpecialFolder("キーボード", "shell:::{725BE8F7-668E-4C7B-8F90-46BDB0936430}");
		case 276:
			return new SpecialFolder("デバイス マネージャー", "shell:::{74246bfc-4c96-11d0-abef-0020af6b0b7a}");
		case 277:
			// Win8まで
			return new SpecialFolder("Windows CardSpace ", "shell:::{78CB147A-98EA-4AA6-B0DF-C8681F69341C}");
		case 278:
			// netplwiz.exe / control.exe userpasswords2
			return new SpecialFolder("ユーザー アカウント", "shell:::{7A9D77BD-5403-11d2-8785-2E0420524153}");
		case 279:
			return new SpecialFolder(WIN8 ? "タブレット PC 設定" : "Tablet PC 設定", "shell:::{80F3F1D5-FECA-45F3-BC32-752C152E456E}");
		case 280:
			// Win10以降では開けないので非表示に
			return new SpecialFolder("インターネット", WIN10 ? null : "shell:InternetFolder");
		case 281:
			return new SpecialFolder("インデックスのオプション", "shell:::{87D66A43-7B11-4A28-9811-C86EE395ACF7}");
		case 282:
			// Win8から
			// Enterpriseで使用可
			// Win10 1607以降ではProでも使用可
			return new SpecialFolder("Windows To Go ワークスペースの作成", "shell:::{8E0C279D-0BD1-43C3-9EBD-31C3DC5B8A77}");
		case 283:
			// Win8まで
			return new SpecialFolder("生体認証デバイスへようこそ", "shell:::{8e35b548-f174-4c7d-81e2-8ed33126f6fd}");
		case 284:
			// Win10 1607から
			return new SpecialFolder("赤外線", "shell:::{A0275511-0E86-4ECA-97C2-ECD8F1221D08}");
		case 285:
			return new SpecialFolder("インターネット オプション", "shell:::{A3DD4F92-658A-410F-84FD-6FBBBEF2FFFE}");
		case 286:
			return new SpecialFolder("色の管理", "shell:::{B2C761C6-29BC-4f19-9251-E6195265BAF1}");
		case 287:
			// Win8.1まで
			return new SpecialFolder(WIN8 ? "Windows への機能の追加" : "Windows Anytime Upgrade", "shell:::{BE122A0E-4503-11DA-8BDE-F66BAD1E3F3A}");
		case 288:
			// Win8まで
			// shell:::{CCFB7955-B4DC-42CE-893D-884D72DD6B19}
			return new SpecialFolder("生体認証デバイス メッセージ", "shell:::{CBC84B69-69EA-439B-B791-DF15F60333CF}");
		case 289:
			return new SpecialFolder("音声合成", "shell:::{D17D1D6D-CC3F-4815-8FE3-607E7D5D10B3}");
		case 290:
			return new SpecialFolder("ネットワークの場所の追加", "shell:::{D4480A50-BA28-11d1-8E75-00C04FA31A86}");
		case 291:
			// Win10 1607まで
			return new SpecialFolder("Windows Defender", "shell:::{D8559EB9-20C0-410E-BEDA-7ED416AECC2A}");
		case 292:
			return new SpecialFolder("日付と時刻", "shell:::{E2E7934B-DCE5-43C4-9576-7FE4F75E7480}");
		case 293:
			return new SpecialFolder("サウンド", "shell:::{F2DDFC82-8F12-4CDD-B7DC-D4FE1425AA4D}");
		case 294:
			return new SpecialFolder("ペンとタッチ", "shell:::{F82DF8F7-8B9F-442E-A48C-818EA735FF9B}");
		
		// 上にあるのとは違うデータでフォルダーの情報を取得する
		// CSIDLは一部を除いて扱わない
		case 295:
			return new SpecialFolder("shell:Profile", "shell:Profile", { category: "OtherDir" });
		case 296:
			return new SpecialFolder("shell:Local Documents", "shell:Local Documents");
		case 297:
			return new SpecialFolder("shell:Local Downloads", "shell:Local Downloads");
		case 298:
			return new SpecialFolder("shell:Local Music", "shell:Local Music");
		case 299:
			return new SpecialFolder("shell:Local Pictures", "shell:Local Pictures");
		case 300:
			return new SpecialFolder("shell:Local Videos", "shell:Local Videos");
		case 301:
			return new SpecialFolder("アドレス帳", "shell:UsersFilesFolder\\{56784854-C6CB-462B-8169-88E350ACB882}");
		case 302:
			return new SpecialFolder("リンク", "shell:UsersFilesFolder\\{bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968}");
		case 303:
			return new SpecialFolder("保存したゲーム", "shell:UsersFilesFolder\\{4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4}");
		case 304:
			return new SpecialFolder("検索", "shell:UsersFilesFolder\\{7d1d3a04-debb-4115-95cf-2f29da2920da}");
		case 305:
			return new SpecialFolder("shell:UsersLibrariesFolder", "shell:UsersLibrariesFolder");
		case 306:
			return new SpecialFolder("カメラ ロール ライブラリ", "shell:Libraries\\{2B20DF75-1EDA-4039-8097-38798227D5B7}");
		case 307:
			return new SpecialFolder("ドキュメント ライブラリ", "shell:Libraries\\{7b0db17d-9cd2-4a93-9733-46cc89022e7c}");
		case 308:
			return new SpecialFolder("ミュージック ライブラリ", "shell:Libraries\\{2112AB0A-C86A-4ffe-A368-0DE96E47012E}");
		case 309:
			return new SpecialFolder("ピクチャ ライブラリ", "shell:Libraries\\{A990AE9F-A03B-4e80-94BC-9912D7504104}");
		case 310:
			return new SpecialFolder("保存済みの写真 ライブラリ", "shell:Libraries\\{E25B5812-BE88-4bd9-94B0-29233477B6C3}");
		case 311:
			return new SpecialFolder("ビデオ ライブラリ", "shell:Libraries\\{491E922F-5643-4af4-A7EB-4E7A138D8174}");
		case 312:
			// 64ビットアプリのみ
			return new SpecialFolder("shell:ProgramFilesX64", "shell:ProgramFilesX64");
		case 313:
			// 64ビットアプリのみ
			return new SpecialFolder("shell:ProgramFilesCommonX64", "shell:ProgramFilesCommonX64");
		case 314:
			return new SpecialFolder("shell:MyComputerFolder", "shell:MyComputerFolder");
		
		case 315:
			return new SpecialFolder("ssfPROFILE", ssfPROFILE, { propertyType: ptVerb });
		
		case 316:
			return new SpecialFolder("%USERPROFILE%", "%USERPROFILE%".xExpand());
		case 317:
			return new SpecialFolder("%HOMEDRIVE%%HOMEPATH%", "%HOMEDRIVE%%HOMEPATH%".xExpand());
		case 318:
			return new SpecialFolder("%OneDrive%", "%OneDrive%".xExpand());
		case 319:
			return new SpecialFolder("%APPDATA%", "%APPDATA%".xExpand());
		case 320:
			return new SpecialFolder("%LOCALAPPDATA%", "%LOCALAPPDATA%".xExpand());
		case 321:
			return new SpecialFolder("%PUBLIC%", "%PUBLIC%".xExpand());
		case 322:
			return new SpecialFolder("%ALLUSERSPROFILE%", "%ALLUSERSPROFILE%".xExpand());
		case 323:
			return new SpecialFolder("%ProgramData%", "%ProgramData%".xExpand());
		case 324:
			return new SpecialFolder("%SystemRoot%", "%SystemRoot%".xExpand());
		case 325:
			return new SpecialFolder("%windir%", "%windir%".xExpand());
		case 326:
			return new SpecialFolder("%ProgramFiles%", "%ProgramFiles%".xExpand());
		case 327:
			return new SpecialFolder("%CommonProgramFiles%", "%CommonProgramFiles%".xExpand());
		
		case 328:
			// Win10から
			return new SpecialFolder("OneDrive", "shell:::{018D5C66-4533-4307-9B53-224DE2ED1FE6}");
		case 329:
			// ライブラリ
			return new SpecialFolder("UsersLibraries", "shell:::{031E4825-7B94-4dc3-B131-E946B44C8DD5}");
		case 330:
			// Win10から
			return new SpecialFolder("Local Downloads", "shell:::{088e3905-0323-4b02-9826-5d99428e115f}");
		case 331:
			// Win10 1709から
			return new SpecialFolder("3D Object", "shell:::{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}");
		case 332:
			// プログラムの取得
			return new SpecialFolder("Install New Programs", "shell:::{15eae92e-f17a-4431-9f28-805e482dafd4}");
		case 333:
			// Win8.1から
			return new SpecialFolder("My Music", "shell:::{1CF1260C-4DD0-4ebb-811F-33C572699FDE}");
		case 334:
			return new SpecialFolder("User Pinned", "shell:::{1f3427c8-5c10-4210-aa03-2ee45287d668}");
		case 335:
			return new SpecialFolder("This PC", "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}");
		case 336:
			// すべてのコントロール パネル項目
			return new SpecialFolder("All Control Panel Items", "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}");
		case 337:
			// プリンター
			return new SpecialFolder("Printers", "shell:::{2227A280-3AEA-1069-A2DE-08002B30309D}");
		case 338:
			// Win10から
			return new SpecialFolder("Local Pictures", "shell:::{24ad3ad4-a569-4530-98e1-ab02f9417aa8}");
		case 339:
			return new SpecialFolder("All Control Panel Items", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\0");
		case 340:
			return new SpecialFolder("Hardware and Sound", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\4");
		case 341:
			return new SpecialFolder("System and Security", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\10");
		case 342:
			return new SpecialFolder("All Control Panel Items", "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\\11");
		case 343:
			// Win8.1から
			return new SpecialFolder("Downloads", "shell:::{374DE290-123F-4565-9164-39C4925E467B}");
		case 344:
			// Win8.1から
			return new SpecialFolder("My Pictures", "shell:::{3ADD1653-EB32-4cb0-BBD7-DFA0ABB5ACCA}");
		case 345:
			// Win10から
			return new SpecialFolder("Local Music", "shell:::{3dfdf296-dbec-4fb4-81d1-6a3438bcf4de}");
		case 346:
			return new SpecialFolder("Applications", "shell:::{4234d49b-0245-4df3-b780-3893943456e1}");
		case 347:
			// パブリック
			return new SpecialFolder("Public Folder", "shell:::{4336a54d-038b-4685-ab02-99bb52d3fb8b}");
		case 348:
			// すべてのコントロール パネル項目
			return new SpecialFolder("Control Panel command object for Start menu and desktop", "shell:::{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}");
		case 349:
			return new SpecialFolder("UsersFiles", "shell:::{59031a47-3f72-44a7-89c5-5595fe6b30ee}");
		case 350:
			// このデバイス
			return new SpecialFolder("This Device", "shell:::{5b934b42-522b-4c34-bbfe-37a3ef7b9c90}");
		case 351:
			// ごみ箱
			return new SpecialFolder("Recycle Bin", "shell:::{645FF040-5081-101B-9F08-00AA002F954E}");
		case 352:
			// プログラムと機能
			return new SpecialFolder("Programs and Features", "shell:::{7b81be6a-ce2b-4676-a29e-eb907a5126c5}");
		case 353:
			// ネットワーク接続
			return new SpecialFolder("Network Connections", "shell:::{7007ACC7-3202-11D1-AAD2-00805FC1270E}");
		case 354:
			// プリンター
			return new SpecialFolder("Remote Printers", "shell:::{863aa9fd-42df-457b-8e4d-0de1b8015c60}");
		case 355:
			// インターネット
			return new SpecialFolder("Internet Folder", "shell:::{871C5380-42A0-1069-A2EA-08002B30309D}");
		case 356:
			// Win8.1のみ
			return new SpecialFolder("OneDrive", "shell:::{8E74D236-7F35-4720-B138-1FED0B85EA75}");
		case 357:
			// 検索結果
			return new SpecialFolder("CLSID_SearchHome", "shell:::{9343812e-1c37-4a49-a12e-4b2d810d956b}");
		case 358:
			// 同期センター
			return new SpecialFolder("Sync Center Folder", "shell:::{9C73F5E5-7AE7-4E32-A8E8-8D23B85255BF}");
		case 359:
			// ネットワーク接続
			return new SpecialFolder("Network Connections", "shell:::{992CFFA0-F557-101A-88EC-00DD010CCC48}");
		case 360:
			// Win8.1から
			return new SpecialFolder("My Video", "shell:::{A0953C92-50DC-43bf-BE83-3742FED03C9C}");
		case 361:
			// Win8.1から
			return new SpecialFolder("Personal", "shell:::{A8CDFF1C-4878-43be-B5FD-F8091C1C60D0}");
		case 362:
			// Win8.1からWin10 1511まで
			return new SpecialFolder("StartMenuAllPrograms", "shell:::{adfa80e7-9769-4ad9-992c-55dc57e1008c}");
		case 363:
			// Win8.1から
			return new SpecialFolder("ThisPCDesktopFolder", "shell:::{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}");
		case 364:
			// ホームグループ
			return new SpecialFolder("Other Users Folder", "shell:::{B4FB3F98-C1EA-428d-A78A-D1F5659CBA93}");
		case 365:
			// フォント
			return new SpecialFolder("Microsoft Windows Font Folder", "shell:ControlPanelFolder\\::{BD84B380-8CA2-1069-AB1D-08000948F534}");
		case 366:
			// 管理ツール (All Users)
			return new SpecialFolder("Administrative Tools", "shell:::{D20EA4E1-3957-11d2-A40B-0C5020524153}");
		case 367:
			// Win10から
			return new SpecialFolder("Local Documents", "shell:::{d3162b92-9365-467a-956b-92703aca08af}");
		case 368:
			// インストールされた更新プログラム
			return new SpecialFolder("Installed Updates", "shell:::{d450a8a1-9568-45c7-9c0e-b4f9fb4537bd}");
		case 369:
			return new SpecialFolder("Default Programs command object for Start menu", "shell:::{E44E5D18-0652-4508-A4E2-8A090067BCB0}");
		case 370:
			// ゲーム
			// Win10 1709までサポート (?)
			return new SpecialFolder("Games Explorer", "shell:::{ED228FDF-9EA8-4870-83b1-96b02CFE0D52}");
		case 371:
			// ネットワーク
			return new SpecialFolder("Computers and Devices", "shell:::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}");
		case 372:
			// このデバイス (Win10 1703から)
			return new SpecialFolder("This Device", "shell:::{f8278c54-a712-415b-b593-b77a2be0dda9}");
		case 373:
			// Win10から
			return new SpecialFolder("Local Videos", "shell:::{f86fa3ab-70d2-4fc7-9c99-fcbf05467f3a}");
		
		// フォルダーとして使えないshellコマンド
		// Get-Clsid.ps1やGet-ShellCommand.ps1でヒットしないようにするためのもの
		case 374:
			return new SpecialFolder("shell:MAPIFolder", "shell:MAPIFolder", { category: "Unusable" });
		case 375:
			return new SpecialFolder("shell:RecordedTVLibrary", "shell:RecordedTVLibrary");
		
		case 376:
			return new SpecialFolder("", "shell:::{00020D75-0000-0000-C000-000000000046}");
		case 377:
			return new SpecialFolder("Desktop", "shell:::{00021400-0000-0000-C000-000000000046}");
		case 378:
			return new SpecialFolder("Shortcut", "shell:::{00021401-0000-0000-C000-000000000046}");
		case 379:
			// Win10 1507から1703まで
			return new SpecialFolder("", "shell:::{047ea9a0-93bb-415f-a1c3-d7aeb3dd5087}");
		case 380:
			return new SpecialFolder("Open With Context Menu Handler", "shell:::{09799AFB-AD67-11d1-ABCD-00C04FC30936}");
		case 381:
			return new SpecialFolder("Folder Shortcut", "shell:::{0AFACED1-E828-11D1-9187-B532F1E9575D}");
		case 382:
			return new SpecialFolder("", "shell:::{0C39A5CF-1A7A-40C8-BA74-8900E6DF5FCD}");
		case 383:
			return new SpecialFolder("", "shell:::{0D45D530-764B-11d0-A1CA-00AA00C16E65}");
		case 384:
			// Win8から
			return new SpecialFolder("Shell File System Folder", "shell:::{0E5AAE11-A475-4c5b-AB00-C66DE400274E}");
		case 385:
			return new SpecialFolder("Device Center Print Context Menu Extension", "shell:::{0e6daa63-dd4e-47ce-bf9d-fdb72ece4a0d}");
		case 386:
			return new SpecialFolder("IE History and Feeds Shell Data Source for Windows Search", "shell:::{11016101-E366-4D22-BC06-4ADA335C892B}");
		case 387:
			return new SpecialFolder("OpenMediaSharing", "shell:::{17FC1A80-140E-4290-A64F-4A29A951A867}");
		case 388:
			return new SpecialFolder("CLSID_DBFolderBoth", "shell:::{1bef2128-2f96-4500-ba7c-098dc0049cb2}");
		case 389:
			return new SpecialFolder("CompatContextMenu Class", "shell:::{1d27f844-3a1f-4410-85ac-14651078412d}");
		case 390:
			return new SpecialFolder("Windows Security", "shell:::{2559a1f2-21d7-11d4-bdaf-00c04f60b9f0}");
		case 391:
			return new SpecialFolder("Location Folder", "shell:::{267cf8a9-f4e3-41e6-95b1-af881be130ff}");
		case 392:
			return new SpecialFolder("Enhanced Storage Context Menu Handler Class", "shell:::{2854F705-3548-414C-A113-93E27C808C85}");
		case 393:
			return new SpecialFolder("System Restore", "shell:::{3f6bc534-dfa1-4ab4-ae54-ef25a74e0107}");
		case 394:
			return new SpecialFolder("Start Menu Folder", "shell:::{48e7caab-b918-4e58-a94d-505519c795dc}");
		case 395:
			return new SpecialFolder("IGD Property Page", "shell:::{4A1E5ACD-A108-4100-9E26-D2FAFA1BA486}");
		case 396:
			// Win10 1607まで
			return new SpecialFolder("LzhCompressedFolder2", "shell:::{4F289A46-2BBB-4AE8-9EDA-E5E034707A71}");
		case 397:
			// Win10から
			return new SpecialFolder("This PC", "shell:::{5E5F29CE-E0A8-49D3-AF32-7A7BDC173478}");
		case 398:
			return new SpecialFolder("", "shell:::{62AE1F9A-126A-11D0-A14B-0800361B1103}");
		case 399:
			return new SpecialFolder("Search Connector Folder", "shell:::{72b36e70-8700-42d6-a7f7-c9ab3323ee51}");
		case 400:
			return new SpecialFolder("CryptPKO Class", "shell:::{7444C717-39BF-11D1-8CD9-00C04FC29D45}");
		case 401:
			return new SpecialFolder("Temporary Internet Files", "shell:::{7BD29E00-76C1-11CF-9DD0-00A0C9034933}");
		case 402:
			return new SpecialFolder("Temporary Internet Files", "shell:::{7BD29E01-76C1-11CF-9DD0-00A0C9034933}");
		case 403:
			return new SpecialFolder(WIN10_1703 ? "" : "Briefcase", "shell:::{85BBD920-42A0-1069-A2E4-08002B30309D}");
		case 404:
			return new SpecialFolder("Shortcut", "shell:::{85cfccaf-2d14-42b6-80b6-f40f65d016e7}");
		case 405:
			return new SpecialFolder("Mobile Broadband Profile Settings Editor", "shell:::{87630419-6216-4ff8-a1f0-143562d16d5c}");
		case 406:
			return new SpecialFolder("Compressed (zipped) Folder SendTo Target", "shell:::{888DCA60-FC0A-11CF-8F0F-00C04FD7D062}");
		case 407:
			return new SpecialFolder("ActiveX Cache Folder", "shell:::{88C6C381-2E85-11D0-94DE-444553540000}");
		case 408:
			return new SpecialFolder("Libraries delegate folder that appears in Users Files Folder", "shell:::{896664F7-12E1-490f-8782-C0835AFD98FC}");
		case 409:
			// Win10 1607まで
			return new SpecialFolder("Windows Search Service Media Center Namespace Extension Handler", "shell:::{98D99750-0B8A-4c59-9151-589053683D73}");
		case 410:
			return new SpecialFolder("MAPI Shell Context Menu", "shell:::{9D3C0751-A13F-46a6-B833-B46A43C30FE8}");
		case 411:
			return new SpecialFolder("Previous Versions", "shell:::{9DB7A13C-F208-4981-8353-73CC61AE2783}");
		case 412:
			return new SpecialFolder("Mail Service", "shell:::{9E56BE60-C50F-11CF-9A2C-00A0C90A90CE}");
		case 413:
			return new SpecialFolder("Desktop Shortcut", "shell:::{9E56BE61-C50F-11CF-9A2C-00A0C90A90CE}");
		case 414:
			return new SpecialFolder("DevicePairingFolder Initialization", "shell:::{AEE2420F-D50E-405C-8784-363C582BF45A}");
		case 415:
			return new SpecialFolder("CLSID_DBFolder", "shell:::{b2952b16-0e07-4e5a-b993-58c52cb94cae}");
		case 416:
			return new SpecialFolder("Device Center Scan Context Menu Extension", "shell:::{b5a60a9e-a4c7-4a93-ac6e-0b76d1d87dc4}");
		case 417:
			return new SpecialFolder("DeviceCenter Initialization", "shell:::{C2B136E2-D50E-405C-8784-363C582BF43E}");
		case 418:
			// Win10 1507から1607まで
			return new SpecialFolder("", "shell:::{D9AC5E73-BB10-467b-B884-AA1E475C51F5}");
		case 419:
			return new SpecialFolder("delegate folder that appears in Users Files Folder", "shell:::{DFFACDC5-679F-4156-8947-C5C76BC0B67F}");
		case 420:
			return new SpecialFolder("CompressedFolder", "shell:::{E88DCCE0-B7B3-11d1-A9F0-00AA0060FA31}");
		case 421:
			return new SpecialFolder("MyDocs Drop Target", "shell:::{ECF03A32-103D-11d2-854D-006008059367}");
		case 422:
			return new SpecialFolder("Shell File System Folder", "shell:::{F3364BA0-65B9-11CE-A9BA-00AA004AE837}");
		case 423:
			// Win10 1607まで
			return new SpecialFolder("Sticky Notes Namespace Extension for Windows Desktop Search", "shell:::{F3F5824C-AD58-4728-AF59-A1EBE3392799}");
		case 424:
			return new SpecialFolder("Subscription Folder", "shell:::{F5175861-2688-11d0-9C5E-00AA00A45957}");
		case 425:
			return new SpecialFolder("Internet Shortcut", "shell:::{FBF23B40-E3F0-101B-8488-00AA003E56F8}");
		case 426:
			return new SpecialFolder("History", "shell:::{FF393560-C2A7-11CF-BFF4-444553540000}");
		case 427:
			index = doneIteration;
			
			return new SpecialFolder("Windows Photo Viewer Image Verbs", "shell:::{FFE2A43C-56B9-4bf5-9A79-CC6D4285608A}");
		
		case doneIteration:
			return null;
		
		default:
			throw new Error("indexが不正: " + index);
		}
	};
})();
