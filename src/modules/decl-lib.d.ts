declare var clipboardData: DataTransfer;

interface Error {
	/**
	 * 特定のエラーと関連付けられた数値を設定します。値の取得も可能です。
	 * @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc427878(v=msdn.10)
	 */
	number: number;
}

/** @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc427715(v=msdn.10) */
interface ErrorConstructor {
	new(number: number): Error;
	new(number: number, message: string): Error;
}

interface Document {
	/**
	 * Retrieves the document compatibility mode of the document.
	 * @see https://msdn.microsoft.com/en-us/windows/desktop/cc196988
	 */
	readonly documentMode?: number;
}

interface Window {
	/**
	 * Creates a modal dialog box that displays the specified HTML document.
	 * @see https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa741502(v=vs.85)
	 */
	showModalDialog(url: string, argIn?: any, options?: any): any;
	/**
	 * Creates a modeless dialog box that displays the specified HTML document.
	 * @see https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa741355(v=vs.85)
	 */
	showModelessDialog(url?: string, argIn?: any, options?: any): Window;
}


/**
 * Provides access to information about an HTML dialog box.
 * @see https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752684(v=vs.85)
*/
interface HTMLDialog {
	/** Gets the variable or array of variables passed into the modal dialog window. */
	dialogArguments: any;
	/** Sets or gets the height of the modal dialog window. */
	dialogHeight: any;
	dialogHide: string;
	/** Sets or gets the left coordinate of the modal dialog window. */
	dialogLeft: any;
	/** Sets or gets the width of the modal dialog window. */
	dialogWidth: any;
	/** Sets or gets the top coordinate of the modal dialog window. */
	dialogTop: any;
	resizable: string;
	/** Sets or gets the value returned from the modal dialog window. */
	returnValue: any;
	status: string;
	unadorned: string;

	/** Closes the current browser window or HTA. */
	close(): void;
}


/** @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc428132(v=msdn.10) */
declare enum DriveTypeConst {
	UnknownType = 0,
	Removable = 1,
	Fixed = 2,
	Remote = 3,
	CDRom = 4,
	RamDisk = 5,
}

/** @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc428087(v=msdn.10) */
declare enum FileAttribute {
	/** 標準ファイル。どの属性も設定されません。 */
	Normal = 0,
	/** 読み取り専用ファイル。この属性は、値の取得も設定も可能です。 */
	ReadOnly = 1,
	/** 隠しファイル。この属性は、値の取得も設定も可能です。 */
	Hidden = 2,
	/** システム ファイル。この属性は、値の取得も設定も可能です。 */
	System = 4,
	/** ディスク ドライブ ボリューム ラベル。この属性は、値の取得のみ可能です。 */
	Volume = 8,
	/** フォルダまたはディレクトリ。この属性は、値の取得のみ可能です。 */
	Directory = 16,
	/** ファイルが前回のバックアップ以降に変更されているかどうか。この属性は、値の取得も設定も可能です。 */
	Archive = 32,
	/** リンクまたはショートカット。この属性は、値の取得のみ可能です。 */
	Alias = 1024,
	/** 圧縮ファイル。この属性は、値の取得のみ可能です。 */
	Compressed = 2048,
}

/** @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc428042(v=msdn.10) */
declare enum IOMode {
	/** ファイルを読み取り専用として開きます。このファイルには書き込むことができません */
	ForReading = 1,
	/** ファイルを書き込み専用として開きます。既存ファイルがある場合、以前の内容は上書きされます。 */
	ForWriting = 2,
	/** ファイルを開き、ファイルの最後に追加して書き込みます。 */
	ForAppending = 8,
}

/**
 * Specifies unique, system-independent values that identify special folders.
 * @see https://docs.microsoft.com/en-us/windows/win32/api/shldisp/ne-shldisp-shellspecialfolderconstants
 */
declare enum ShellSpecialFolderConstants {
	/** Windows desktop—the virtual folder that is the root of the namespace. */
	ssfDESKTOP = 0,
	/** File system directory that contains the user's program groups (which are also file system directories). */
	ssfPROGRAMS = 2,
	/** Virtual folder that contains icons for the Control Panel applications. */
	ssfCONTROLS = 3,
	/** Virtual folder that contains installed printers. */
	ssfPRINTERS = 4,
	/** File system directory that serves as a common repository for a user's documents. */
	ssfPERSONAL = 5,
	/** File system directory that serves as a common repository for the user's favorite URLs. */
	ssfFAVORITES = 6,
	/**
	 * File system directory that corresponds to the user's Startup program group.
	 * The system starts these programs whenever any user first logs into their profile after a reboot.
	 */
	ssfSTARTUP = 7,
	/** File system directory that contains the user's most recently used documents. */
	ssfRECENT = 8,
	/** File system directory that contains Send To menu items. */
	ssfSENDTO = 9,
	/** Virtual folder that contains the objects in the user's Recycle Bin. */
	ssfBITBUCKET = 10,
	/** File system directory that contains Start menu items. */
	ssfSTARTMENU = 11,
	/**
	 * File system directory used to physically store the file objects that are displayed on the desktop.
	 * It is not to be confused with the desktop folder itself, which is a virtual folder.
	 */
	ssfDESKTOPDIRECTORY = 16,
	/**
	 * My Computer—the virtual folder that contains everything on the local computer:
	 * storage devices, printers, and Control Panel. This folder can also contain mapped network drives.
	 */
	ssfDRIVES = 17,
	/** Network Neighborhood—the virtual folder that represents the root of the network namespace hierarchy. */
	ssfNETWORK = 18,
	/**
	 * A file system folder that contains any link objects in the My Network Places virtual folder.
	 * It is not the same as ssfNETWORK, which represents the network namespace root.
	 */
	ssfNETHOOD = 19,
	/** Virtual folder that contains installed fonts. */
	ssfFONTS = 20,
	/** File system directory that serves as a common repository for document templates. */
	ssfTEMPLATES = 21,
	/** File system directory that contains the programs and folders that appear on the Start menu for all users. */
	ssfCOMMONSTARTMENU = 22,
	/** File system directory that contains the directories for the common program groups that appear on the Start menu for all users. */
	ssfCOMMONPROGRAMS = 23,
	/** File system directory that contains the programs that appear in the Startup folder for all users. */
	ssfCOMMONSTARTUP = 24,
	/** File system directory that contains files and folders that appear on the desktop for all users. */
	ssfCOMMONDESKTOPDIR = 25,
	/** File system directory that serves as a common repository for application-specific data. */
	ssfAPPDATA = 26,
	/** File system directory that contains any link objects in the Printers virtual folder. */
	ssfPRINTHOOD = 27,
	/** File system directory that serves as a data repository for local (non-roaming) applications. */
	ssfLOCALAPPDATA = 28,
	/** File system directory that corresponds to the user's non-localized Startup program group. */
	ssfALTSTARTUP = 29,
	/** File system directory that corresponds to the non-localized Startup program group for all users. */
	ssfCOMMONALTSTARTUP = 30,
	/** File system directory that serves as a common repository for the favorite URLs shared by all users. */
	ssfCOMMONFAVORITES = 31,
	/** File system directory that serves as a common repository for temporary Internet files. */
	ssfINTERNETCACHE = 32,
	/** File system directory that serves as a common repository for Internet cookies. */
	ssfCOOKIES = 33,
	/** File system directory that serves as a common repository for Internet history items. */
	ssfHISTORY = 34,
	/** Application data for all users. */
	ssfCOMMONAPPDATA = 35,
	/** Windows directory. This corresponds to the %windir% or %SystemRoot% environment variables. */
	ssfWINDOWS = 36,
	/** The System folder. */
	ssfSYSTEM = 37,
	/** Program Files folder. */
	ssfPROGRAMFILES = 38,
	/** My Pictures folder. */
	ssfMYPICTURES = 39,
	/** User's profile folder. */
	ssfPROFILE = 40,
	/** System folder. */
	ssfSYSTEMx86 = 41,
	/** Program Files folder. */
	ssfPROGRAMFILESx86 = 42,
}

/** @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc428028(v=msdn.10) */
declare enum SpecialFolderConst {
	/** Windows オペレーティング システムによりセットアップされたファイルの置かれている Windows フォルダが返されます。 */
	WindowsFolder = 0,
	/** ライブラリ、フォント、デバイス ドライバなどの置かれている System フォルダが返されます。 */
	SystemFolder = 1,
	/** 一時ファイルの格納に使用される Temp フォルダが返されます。このパスは、環境変数 TMP より取得します。 */
	TemporaryFolder = 2,
}

/** @see https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/windows-scripting/y6hbz9es(v=vs.84) */
declare enum StandardStreamTypes {
	/** Returns a TextStream object corresponding to the standard input stream. */
	StdIn = 0,
	/** Returns a TextStream object corresponding to the standard output stream. */
	StdOut = 1,
	/** Returns a TextStream object corresponding to the standard error stream. */
	StdErr = 2,
}

/** @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc428044(v=msdn.10) */
declare enum Tristate {
	/** システム デフォルトを使ってファイルを開きます。 */
	TristateUseDefault = -2,
	TristateMixed = -2,
	/** ファイルを Unicode ファイルとして開きます。 */
	TristateTrue = -1,
	/** ファイルを ASCII ファイルとして開きます。 */
	TristateFalse = 0,
}

/** @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364410(v=msdn.10) */
declare enum WshExecStatus {
	/** ジョブはまだ実行中です。 */
	WshRunning = 0,
	/** ジョブの実行が完了しました。 */
	WshFinished = 1,
	WshFailed = 2,
}


/**
 * ファイル システムへアクセスする方法を提供します。
 * @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc428071(v=msdn.10)
 */
interface FileSystemObject {
	/** ローカル コンピュータ上で利用可能なすべての Drive オブジェクトからなる Drives コレクションを返します。 */
	readonly Drives: DriveCollection;

	/** 既存のパスの末尾に名前を追加します。 */
	BuildPath(path: string, name: string): string;
	/** ファイルを別の場所へコピーします。 */
	CopyFile(source: string, destination: string, overwrite?: boolean): void;
	/** フォルダを、ある場所から別の場所に再帰的にコピーします。 */
	CopyFolder(source: string, destination: string, overwrite?: boolean): void;
	/** フォルダを作成します。 */
	CreateFolder(foldername: string): FsoFolder;
	/** 指定した名前のファイルを作成し、そのファイルの読み取り、書き込みに使用できる TextStream オブジェクトを返します。 */
	CreateTextFile(filename: string, overwrite?: boolean, unicode?: boolean): TextStreamWriter;
	/** 指定したファイルを削除します。 */
	DeleteFile(filespec: string, force?: boolean): void;
	/** 指定したフォルダとその中身を削除します。 */
	DeleteFolder(folderspec: string, force?: boolean): void;
	/** 指定したドライブが存在すれば true を返し、存在しなければ false を返します。 */
	DriveExists(drivespec: string): boolean;
	/** 指定したファイルが存在すれば true を返し、存在しなければ false を返します。 */
	FileExists(filespec: string): boolean;
	/** 指定したフォルダが存在すれば true を返し、存在しなければ false を返します。 */
	FolderExists(folderspec: string): boolean;
	/** 指定したパスから、省略されていない完全なパスを返します。 */
	GetAbsolutePathName(pathspec: string): string;
	/** パスの最後の構成要素のベース名 (ファイル拡張子を除いたもの) を表す文字列を返します。 */
	GetBaseName(path: string): string;
	/** 指定したパスのドライブに対応する Drive オブジェクトを返します。 */
	GetDrive(driveSpec: string): Drive;
	/** 指定したパスのドライブ名を表す文字列を返します。 */
	GetDriveName(path: string): string;
	/** パスの最後の構成要素の拡張子名を表す文字列を返します。 */
	GetExtensionName(path: string): string;
	/** 指定したパスにあるファイルに対応する File オブジェクトを返します。 */
	GetFile(filespec: string): FsoFile;
	/** 指定したパスの最後の構成要素のうちドライブ指定以外の部分を返します。 */
	GetFileName(pathspec: string): string;
	/** 指定したファイルのバージョン番号を返します。 */
	GetFileVersion(pathspec: string): string;
	/** 指定したパスにあるフォルダに対応する Folder オブジェクトを返します。 */
	GetFolder(folderspec: string): FsoFolder;
	/** 指定したパスの最後の構成要素の親フォルダ名を表す文字列を返します。 */
	GetParentFolderName(path: string): string;
	/** 指定した特殊フォルダのオブジェクトを返します。 */
	GetSpecialFolder(folderspec: SpecialFolderConst): FsoFolder;
	/** Returns a TextStream object corresponding to the standard input stream. */
	GetStandardStream(standardStreamType: StandardStreamTypes.StdIn, unicode?: boolean): TextStreamReader;
	/** Returns a TextStream object corresponding to the standard output, or error stream. */
	GetStandardStream(standardStreamType: StandardStreamTypes.StdOut | StandardStreamTypes.StdErr, unicode?: boolean): TextStreamWriter;
	/** ランダムに生成される一時ファイルまたは一時フォルダの名前を返します。これらは一時ファイルや一時フォルダを必要とする処理を実行する際に便利です。 */
	GetTempName(): string;
	/** 1 つまたは複数のファイルを、ある場所から別の場所に移動します。 */
	MoveFile(source: string, destination: string): void;
	/** 1 つまたは複数のフォルダを、ある場所から別の場所に移動します。 */
	MoveFolder(source: string, destination: string): void;
	/** 指定したファイルを開き、ファイルの読み込みに使用できる TextStream オブジェクトを返します。 */
	OpenTextFile(filename: string, iomode?: IOMode.ForReading, create?: boolean, format?: Tristate): TextStreamReader;
	/** 指定したファイルを開き、ファイルの書き込みや追加書き込みに使用できる TextStream オブジェクトを返します。 */
	OpenTextFile(filename: string, iomode: IOMode.ForWriting | IOMode.ForAppending, create?: boolean, format?: Tristate): TextStreamWriter;
}

/**
 * ディスク ドライブまたはネットワーク共有の各種プロパティへアクセスする手段を提供します。
 * @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc428067(v=msdn.10)
 */
interface Drive {
	/** 指定したドライブ上またはネットワーク共有上でユーザーが利用可能な領域の量を返します。 */
	readonly AvailableSpace: any;
	/** 物理ローカル ドライブまたはネットワーク共有のドライブ名を返します。 */
	readonly DriveLetter: string;
	/** 指定したドライブの種類を示す値を返します。 */
	readonly DriveType: DriveTypeConst;
	/** 指定したドライブで使用されているファイル システムの種類を返します。 */
	readonly FileSystem: string;
	/** 指定したドライブ上またはネットワーク共有上でユーザーが利用可能な空き領域の量を返します。 */
	readonly FreeSpace: any;
	/** 指定したドライブがレディ状態であれば true を返し、レディ状態でなければ false を返します。 */
	readonly IsReady: boolean;
	/** 指定したドライブのパスを返します。 */
	readonly Path: string;
	/** 指定したドライブのルート フォルダを表す Folder オブジェクトを返します。 */
	readonly RootFolder: FsoFolder;
	/** ディスク ボリュームを一意に識別するために使用する 10 進シリアル番号を返します。 */
	readonly SerialNumber: number;
	/** 指定したドライブのネットワーク共有名を返します。 */
	readonly ShareName: string;
	/** ドライブまたはネットワーク共有の総容量をバイト単位で返します。 */
	readonly TotalSize: any;
	/** 指定したドライブのボリューム名を設定または参照します。 */
	VolumeName: string;
}

/**
 * 利用可能なすべてのドライブの読み取り専用コレクションです。
 * @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc427962(v=msdn.10)
 */
interface DriveCollection {
	/** コレクション内にある項目の数を返します。 */
	readonly Count: number;
	/** 指定されたキーに対応するオブジェクトを返します。 */
	Item(key: any): Drive;
}

/**
 * ファイルのあらゆるプロパティにアクセスする手段を提供します。
 * @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc428069(v=msdn.10)
 */
interface FsoFile {
	/** ファイルの属性を設定または参照します。 */
	Attributes: FileAttribute;
	/** 指定したファイルが作成された日付と時刻を返します。読み取り専用です。 */
	readonly DateCreated: VarDate;
	/** 指定したファイルが最後にアクセスされた日付と時刻を返します。 */
	readonly DateLastModified: VarDate;
	/** 指定したファイルが最後に変更された日付と時刻を返します。 */
	readonly DateLastAccessed: VarDate;
	/** 指定したファイルのあるドライブのドライブ名を返します。 */
	readonly Drive: Drive;
	/** 指定したファイルの名前を設定または参照します。 */
	Name: string;
	/** 指定したファイルの親にあたるフォルダ オブジェクトを返します。 */
	readonly ParentFolder: FsoFolder;
	/** 指定したファイルのパスを返します。 */
	readonly Path: string;
	/** 従来の 8.3 命名規則を必要とするプログラムで使用する短い名前を返します。  */
	readonly ShortName: string;
	/** 従来の 8.3 命名規則を必要とするプログラムで使用する短いパスを返します。 */
	readonly ShortPath: string;
	/** 指定したファイルのサイズをバイト単位で返します。 */
	readonly Size: any;
	/** ファイルの種類に関する情報を返します。  */
	readonly Type: string;

	/** 指定したファイルを別の場所にコピーします。 */
	Copy(destination: string, overwrite?: boolean): void;
	/** 指定したファイルを削除します。 */
	Delete(force?: boolean): void;
	/** 指定されたファイルを別の場所へ移動します。 */
	Move(destination: string): void;
	/** 指定したファイルを開き、ファイルの読み込みに使用できる TextStream オブジェクトを返します。 */
	OpenAsTextStream(iomode?: IOMode.ForReading, format?: Tristate): TextStreamReader;
	/** 指定したファイルを開き、ファイルの書き込みや追加書き込みに使用できる TextStream オブジェクトを返します。 */
	OpenAsTextStream(iomode: IOMode.ForWriting | IOMode.ForAppending, format?: Tristate): TextStreamWriter;
}

/**
 * フォルダ内にあるすべての File オブジェクトのコレクションです。
 * @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc427964(v=msdn.10)
 */
interface FileCollection {
	/** コレクション内にある項目の数を返します。 */
	readonly Count: number;
	/** 指定されたキーに対応するオブジェクトを返します。 */
	Item(key: any): FsoFile;
}

/**
 * フォルダのあらゆるプロパティにアクセスする手段を提供します。
 * @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc428096(v=msdn.10)
 */
interface FsoFolder {
	/** フォルダの属性を設定または参照します。 */
	Attributes: FileAttribute;
	/** 指定したフォルダが作成された日付と時刻を返します。読み取り専用です。 */
	readonly DateCreated: VarDate;
	/** 指定したフォルダが最後にアクセスされた日付と時刻を返します。 */
	readonly DateLastAccessed: VarDate;
	/** 指定したフォルダが最後に変更された日付と時刻を返します。 */
	readonly DateLastModified: VarDate;
	/** 指定したフォルダのあるドライブのドライブ名を返します。 */
	readonly Drive: Drive;
	/** 指定したフォルダ内にあるすべての File オブジェクトからなる Files コレクションを返します。隠しファイル属性やシステム ファイル属性が設定されたものも含まれます。 */
	readonly Files: FileCollection;
	/** 指定したフォルダがルート フォルダであれば true を返し、ルート フォルダでなければ false を返します。 */
	readonly IsRootFolder: boolean;
	/** 指定したフォルダの名前を設定または参照します。 */
	Name: string;
	/** 指定したフォルダの親にあたるフォルダ オブジェクトを返します。 */
	readonly ParentFolder: FsoFolder;
	/** 指定したフォルダのパスを返します。 */
	readonly Path: string;
	/** 従来の 8.3 命名規則を必要とするプログラムで使用する短い名前を返します。 */
	readonly ShortName: string;
	/** 従来の 8.3 命名規則を必要とするプログラムで使用する短いパスを返します。 */
	readonly ShortPath: string;
	/** フォルダに含まれているすべてのファイルおよびサブフォルダの合計サイズをバイト単位で返します。 */
	readonly Size: any;
	/** 指定したフォルダ内にあるすべてのフォルダからなる Folders コレクションを返します。隠しファイル属性やシステム ファイル属性が設定されたものも含まれます。 */
	readonly SubFolders: FolderCollection;
	/** フォルダの種類に関する情報を返します。 */
	readonly Type: string;

	/** 指定したフォルダを別の場所にコピーします。 */
	Copy(destination: string, overwrite?: boolean): void;
	/** 指定した名前のファイルを作成し、作成したファイルの読み取りまたは書き込みに使用できる TextStream オブジェクトを返します。 */
	CreateTextFile(filename: string, overwrite?: boolean, unicode?: boolean): TextStreamWriter;
	/** 指定したフォルダを削除します。 */
	Delete(force?: boolean): void;
	/** 指定されたファイルまたはフォルダを別の場所へ移動します。 */
	Move(destination: string): void;
}

/**
 * Folder オブジェクト内に含まれているすべての Folder オブジェクトのコレクションです。
 * @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc428003(v=msdn.10)
 */
interface FolderCollection {
	/** コレクション内にある項目の数を返します。 */
	readonly Count: number;
	/** 指定されたキーに対応するオブジェクトを返します。 */
	Item(key: any): FsoFolder;

	/** 新しいフォルダを Folders コレクションに追加します。 */
	Add(name: string): FsoFolder;
}


/**
 * Represents the objects in the Shell. Methods are provided to control the Shell and to execute commands within the Shell.
 * @see https://docs.microsoft.com/en-us/windows/win32/shell/shell
 */
interface Shell {
	/** Contains the object's Application object. */
	readonly Application: any;
	/** Gets an object that represents the parent of the current object. */
	readonly Parent: any;

	/** Adds a file to the most recently used (MRU) list. */
	AddToRecent(varFile: any, bstrCategory?: string): void;
	/** Creates a dialog box that enables the user to select a folder and then returns the selected folder's Folder object. */
	BrowseForFolder(hWnd: number, sTitle: string, iOptions: number, vRootFolder?: any): ShlFolder;
	/** Determines if the current user can start and stop the named service. */
	CanStartStopService(sServiceName: string): any;
	/** Cascades all of the windows on the desktop. */
	CascadeWindows(): void;
	/** Runs the specified Control Panel (*.cpl) application. If the application is already open, it will activate the running instance.  */
	ControlPanelItem(bstrDir: string): void;
	/** Ejects the computer from its docking station. */
	EjectPC(): void;
	/** Opens a specified folder in a Windows Explorer window. */
	Explore(vDir: any): void;
	/** Gets the value for a specified Internet Explorer policy. */
	ExplorerPolicy(bstrPolicyName: string): any;
	/** Displays the Run dialog to the user. */
	FileRun(): void;
	/** Displays the Search Results: Computers dialog box. */
	FindComputer(): void;
	/** Displays the Find: All Files dialog box. */
	FindFiles(): void;
	/** Displays the Find Printer dialog box. */
	FindPrinter(sName?: string, sLocation?: string, sModel?: string): void;
	/** Retrieves a global Shell setting. */
	GetSetting(lSetting: number): boolean;
	/** Retrieves system information. */
	GetSystemInformation(sName: string): any;
	/** Displays the Windows Help and Support Center. */
	Help(): void;
	/** Retrieves a group's restriction setting from the registry. */
	IsRestricted(sGroup: string, sRestriction: string): number;
	/** Returns a value that indicates whether a particular service is running. */
	IsServiceRunning(sServiceName: string): any;
	/** Minimizes all of the windows on the desktop. */
	MinimizeAll(): void;
	/** Creates and returns a Folder object for the specified folder. */
	NameSpace(vDir: any): ShlFolder;
	/** Opens the specified folder. */
	Open(vDir: any): void;
	/** Refreshes the contents of the Start menu. */
	RefreshMenu(): void;
	/** Displays the Apps Search pane. */
	SearchCommand(): void;
	/** Starts a named service. */
	ServiceStart(sServiceName: string, vPersistent: any): any;
	/** Stops a named service. */
	ServiceStop(sServiceName: string, vPersistent: any): any;
	/** Displays the Date and Time Properties dialog box. */
	SetTime(): void;
	/** Performs a specified operation on a specified file. */
	ShellExecute(sFile: string, vArguments?: any, vDirectory?: any, vOperation?: any, vShow?: any): void;
	/** Displays a browser bar. */
	ShowBrowserBar(sCLSID: string, vShow: any): any;
	/** Displays the Shut Down Windows dialog box. */
	ShutdownWindows(): void;
	/** Tiles all of the windows on the desktop horizontally. */
	TileHorizontally(): void;
	/** Tiles all of the windows on the desktop vertically. */
	TileVertically(): void;
	/** Displays or hides the desktop. */
	ToggleDesktop(): void;
	/** Displays the Taskbar and Start Menu Properties dialog box. */
	TrayProperties(): void;
	/** Restores all desktop windows to the same state they were in before the last MinimizeAll command. */
	UndoMinimizeALL(): void;
	/** Creates and returns a ShellWindows object. */
	Windows(): any;
	/** Displays the Windows Security dialog box. */
	WindowsSecurity(): void;
	/** Displays your open windows in a 3D stack that you can flip through. */
	WindowSwitcher(): void;
}

/**
 * Represents a Shell folder. This object contains properties and methods that allow you to retrieve information about the folder.
 * @see https://docs.microsoft.com/en-us/windows/win32/shell/folder2-object
 */
interface ShlFolder {
	/** Contains the folder's Application object. */
	readonly Application: Shell;
	/** Contains the offline status of the folder. */
	readonly OfflineStatus: number;
	/** Contains the parent Folder object. */
	readonly ParentFolder: ShlFolder;
	/** Contains the folder's FolderItem object. */
	readonly Self: FolderItem;
	/** Ask if the WebView barricade should be shown or not. */
	ShowWebViewBarricade: boolean;
	/** Contains the title of the folder. */
	readonly Title: string;

	/** Copies an item or items to a folder. */
	CopyHere(vItem: (string | FolderItem | FolderItems), vOptions?: any): void;
	/** Called in response to the web view barricade being dismissed by the user. */
	DismissedWebViewBarricade(): void;
	/** Retrieves details about an item in a folder. */
	GetDetailsOf(vItem: FolderItem, iColumn: number): string;
	/** Retrieves a FolderItems object that represents the collection of items in the folder. */
	Items(): FolderItems;
	/** Moves an item or items to this folder. */
	MoveHere(vItem: (string | FolderItem | FolderItems), vOptions?: any): void;
	/** Creates a new folder. */
	NewFolder(bName: string, vOptions?: any): void;
	/** Creates and returns a FolderItem object that represents a specified item. */
	ParseName(bName: string): FolderItem;
	/** Synchronizes all offline files in the folder. */
	Synchronize(): void;
}

/**
 * Represents an item in a Shell folder. This object contains properties and methods that allow you to retrieve information about the item.
 * @see https://docs.microsoft.com/en-us/windows/win32/shell/shellfolderitem-object
 */
interface FolderItem {
	/** Contains the Application object of the folder item. */
	readonly Application: Shell;
	/** Contains the item's Folder object, if the item is a folder. */
	readonly GetFolder: ShlFolder;
	/** Contains the item's ShellLinkObject object, if the item is a shortcut. */
	readonly GetLink: ShellLinkObject;
	/** Indicates if the item can be hosted inside a browser or Windows Explorer frame. */
	readonly IsBrowsable: boolean;
	/** Indicates if the item is part of the file system. */
	readonly IsFileSystem: boolean;
	/** Indicates if the item is a folder. */
	readonly IsFolder: boolean;
	/** Indicates whether the item is a shortcut. */
	readonly IsLink: boolean;
	/** Sets or gets the date and time that a file was last modified. */
	ModifyDate: VarDate;
	/** Sets or gets the item's name. */
	Name: string;
	/** Gets an object that represents the parent of the item. */
	readonly Parent: ShlFolder;
	/** Contains the item's full path and name. */
	readonly Path: string;
	/** Contains the item's size. */
	readonly Size: number;
	/** Contains a string representation of the item's type. */
	readonly Type: string;

	/** Gets the value of a property from an item's property set. */
	ExtendedProperty(sPropName: string): any;
	/** Executes a verb on the item. */
	InvokeVerb(vVerb?: any): void;
	/** Executes a verb on a Shell item. */
	InvokeVerbEx(vVerb?: any, vArgs?: any): void;
	/** Retrieves the item's FolderItemVerbs object. */
	Verbs(): FolderItemVerbs;
}

/**
 * Represents the collection of items in a Shell folder. This object contains properties and methods that allow you to retrieve information about the collection.
 * @see https://docs.microsoft.com/en-us/windows/win32/shell/folderitems3-object
 */
interface FolderItems {
	/** Contains the Application object of the folder items collection. */
	readonly Application: Shell;
	/** Contains the number of items in the collection. */
	readonly Count: number;
	/** Gets the list of verbs common to all the folder items. */
	readonly Verbs: FolderItemVerbs;

	/** Sets a wildcard filter to apply to the items returned. */
	Filter(grfFlags: number, bstrFilter: string): void;
	/** Executes a verb on a collection of FolderItem objects. */
	InvokeVerbEx(vVerb?: any, vArgs?: any): void;
	/** Retrieves the FolderItem object for a specified item in the collection. */
	Item(iIndex?: any): FolderItem;
}

/**
 * Represents a single verb available to an item. This object contains properties and methods that allow you to retrieve information about the verb.
 * @see https://docs.microsoft.com/en-us/windows/win32/shell/folderitemverb
 */
interface FolderItemVerb {
	/** Contains the verb's name. */
	readonly Name: string;

	/** Executes a verb on the FolderItem associated with the verb. */
	DoIt(): void;
}

/**
 * Represents the collection of verbs for an item in a Shell folder. This object contains properties and methods that allow you to retrieve information about the collection.
 * @see https://docs.microsoft.com/en-us/windows/win32/shell/folderitemverbs
 */
interface FolderItemVerbs {
	/** Contains the number of items in the collection. */
	readonly Count: number;

	/** Retrieves the FolderItemVerb object for a specified item in the collection. */
	Item(iIndex?: number): FolderItemVerb;
}

/**
 * Manages Shell links.
 * @see https://docs.microsoft.com/en-us/windows/win32/shell/ishelllinkdual2-object
 */
interface ShellLinkObject {
	/** Contains a link's arguments. */
	Arguments: string;
	/** Gets or sets the description of the link. */
	Description: string;
	/** Gets or sets the keyboard shortcut for the link. */
	Hotkey: number;
	/** Gets or sets the path to the link object. */
	Path: string;
	/** Gets or sets the initial display state (sized, minimized, or maximized) of the link's command. */
	ShowCommand: number;
	/** Contains the link object's target */
	readonly Target: FolderItem;
	/** Gets or sets the working directory specified in the link. */
	WorkingDirectory: string;

	/** Gets the location of the icon assigned to the link. */
	GetIconLocation(/* out */ sPath: string): number;
	/** Looks for the target of a Shell link, even if the target has been moved or renamed. */
	Resolve(fFlags: number): void;
	/** Saves all changes to the link. */
	Save(sFile?: string): void;
	/** Sets the location of the icon assigned to the link. */
	SetIconLocation(sPath: string, iIndex: number): void;
}


/**
 * Windows のネイティブ シェル機能にアクセスできます。
 * @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364436(v=msdn.10)
 */
interface WshShell {
	/** 現在アクティブになっているディレクトリを取得または変更します。 */
	CurrentDirectory: string;
	/** WshEnvironment オブジェクトを返します。 */
	Environment(type?: string): WshEnvironment;
	/** SpecialFolders オブジェクト (特殊フォルダのコレクション) を返します。 */
	readonly SpecialFolders: WshCollection;

	/** アプリケーション ウィンドウをアクティブにします。 */
	AppActivate(app: any, wait?: any): boolean;
	/** ショートカットまたは URL ショートカットへのオブジェクト参照を作成します。 */
	CreateShortcut(pathname: string): (WshShortcut | WshURLShortcut);
	/** 子コマンドシェルでアプリケーションを実行します。アプリケーションから StdIn/StdOut/StdErr ストリームにアクセスできます。 */
	Exec(command: string): WshScriptExec;
	/** 要求された環境変数を実行中のプロセスから展開し、結果を文字列として返します。 */
	ExpandEnvironmentStrings(string: string): string;
	/** Windows NT イベント ログにイベントを記録します。 */
	LogEvent(type: number, message: string, target?: string): boolean;
	/** ポップアップ メッセージ ボックスにテキストを表示します。 */
	Popup(text: string, secondsToWait?: number, title?: string, type?: number): number;
	/** レジストリから指定されたキーまたは値を削除します。 */
	RegDelete(name: string): void;
	/** レジストリ内のキー名または値名の値を返します。 */
	RegRead(name: string): any;
	/** 新しいキーの作成、新しい値名の既存キーへの追加 (および値の設定)、既存の値名の値変更などを行います。 */
	RegWrite(name: string, value: any, type?: string): void;
	/** 新しいプロセス内でプログラムを実行します。 */
	Run(command: string, windowStyle?: number, waitOnReturn?: boolean): number;
	/** キーボードから入力したときのように、1 つ以上のキー ストロークをアクティブなウィンドウに送ります。 */
	SendKeys(string: string, wait?: number): void;
}

interface WshCollection {
	/** コレクションから、指定されたアイテムを返します。 */
	Item(index: any): any;
	/** コレクション内のアイテム数を返します。 */
	readonly length: number;

	/** オブジェクトの数を返します。 */
	Count(): number;
}

/**
 * Windows 環境変数のコレクションへのアクセスを提供します。
 * @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364435(v=msdn.10)
 */
interface WshEnvironment {
	/** コレクションから、指定されたアイテムを返します。 */
	Item(name: string): string;
	/** ローカル コンピュータ システム上に存在する Windows 環境変数の個数 (Environment コレクション内のアイテム数) を返します。 */
	readonly length: number;

	/** オブジェクトの数を返します。 */
	Count(): number;
	/** 指定された既存の環境変数を削除します。 */
	Remove(name: string): void;
}

/**
 * Exec を使って実行したスクリプトのステータス情報を提供します。また、StdIn、StdOut、および StdErr の各ストリームへのアクセスも提供します。
 * @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364375(v=msdn.10)
 */
interface WshScriptExec {
	/** Exec() メソッドで実行したスクリプトやプログラムによって設定された終了コードを返します。 */
	readonly ExitCode: number;
	/** WshScriptEcec オブジェクトから起動したプロセスのプロセス ID (PID) を返します。 */
	readonly ProcessID: number;
	/** Exec() メソッドで実行したスクリプトに関するステータス情報を返します。 */
	readonly Status: WshExecStatus;
	/** Exec オブジェクトの stderr 出力ストリームへのアクセスを提供します。 */
	readonly StdErr: TextStreamWriter;
	/** Exec オブジェクトの stdin 入力ストリームを公開します。 */
	readonly StdIn: TextStreamReader;
	/** Exec オブジェクトの書き込み専用の stdout 出力ストリームを公開します。 */
	readonly StdOut: TextStreamWriter;

	/** Exec メソッドで開始したプロセスをスクリプト エンジンに終了させます。 */
	Terminate(): void;
}

/**
 * ショートカットへのオブジェクト参照を作成します。
 * @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364438(v=msdn.10)
 */
interface WshShortcut {
	/** ショートカットの引数を設定または識別します。 */
	Arguments: string;
	/** ショートカットの説明を返します。 */
	Description: string;
	/** ショートカット オブジェクトのリンク先への絶対パスを返します。 */
	readonly FullName: string;
	/** ショートカットに対するキーの組み合わせの割り当てを設定および取得できます。 */
	Hotkey: string;
	/** アイコンをショートカットに割り当てるか、割り当てられたアイコンを識別します。 */
	IconLocation: string;
	/** ショートカットの相対パスの割り当てまたは識別を行います。 */
	readonly RelativePath: string;
	/** ショートカットのリンク先の実行可能ファイルのパスを設定および取得できます。 */
	TargetPath: string;
	/** ショートカットに適用するウィンドウ スタイルを設定および取得できます。  */
	WindowStyle: number;
	/** ショートカットに適用する作業ディレクトリを設定および取得できます。  */
	WorkingDirectory: string;

	/** ショートカット オブジェクトを保存します。 */
	Save(): void;
}

/**
 * URL ショートカットへのオブジェクト参照を作成します。
 * @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364464(v=msdn.10)
 */
interface WshURLShortcut {
	/** ショートカット オブジェクトのリンク先への絶対パスを返します。 */
	readonly FullName: string;
	/** ショートカットのリンク先のパスを設定および取得できます。 */
	TargetPath: string;

	/** ショートカット オブジェクトを保存します。 */
	Save(): void;
}


/**
 * コマンド ラインの名前付き引数へのアクセスを提供します。
 * @see https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc364367(v=msdn.10)
 */
interface WshNamed {
	/** WshNamed オブジェクト内のアイテムへのアクセスを提供します。 */
	Item(item: string): string | boolean;
	/** 特定のスクリプトのコマンド ライン パラメータの個数 (引数コレクション内のアイテム数) を返します。 */
	readonly length: number;

	/** WshNamed オブジェクトのスイッチの数を返します。 */
	Count(): number;
	/** WshNamed オブジェクトに特定のキー値が格納されているかどうかを示します。 */
	Exists(item: string): boolean;
}


interface ActiveXObject {
	new(s: "Scripting.FileSystemObject"): FileSystemObject;
	new(s: "Shell.Application"): Shell;
	new(s: "WScript.Shell"): WshShell;
}
