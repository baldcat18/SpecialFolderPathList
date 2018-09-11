/// <reference path="decl-lib.d.ts" />

declare enum Key {
	backspace = 8,
	tab = 9,
	enter = 13,
	shift = 16,
	ctrl = 17,
	alt = 18,
	pause = 19,
	capsLock = 20,
	escape = 27,
	space = 32,
	pageUp = 33,
	pageDown = 34,
	end = 35,
	home = 36,
	leftArrow = 37,
	upArrow = 38,
	rightArrow = 39,
	downArrow = 40,
	insert = 45,
	deleteKey = 46,
	num0 = 48,
	num1 = 49,
	num2 = 50,
	num3 = 51,
	num4 = 52,
	num5 = 53,
	num6 = 54,
	num7 = 55,
	num8 = 56,
	num9 = 57,
	a = 65,
	b = 66,
	c = 67,
	d = 68,
	e = 69,
	f = 70,
	g = 71,
	h = 72,
	i = 73,
	j = 74,
	k = 75,
	l = 76,
	m = 77,
	n = 78,
	o = 79,
	p = 80,
	q = 81,
	r = 82,
	s = 83,
	t = 84,
	u = 85,
	v = 86,
	w = 87,
	x = 88,
	y = 89,
	z = 90,
	leftWindows = 91,
	rightWindows = 92,
	menu = 93,
	numPad0 = 96,
	numPad1 = 97,
	numPad2 = 98,
	numPad3 = 99,
	numPad4 = 100,
	numPad5 = 101,
	numPad6 = 102,
	numPad7 = 103,
	numPad8 = 104,
	numPad9 = 105,
	multiply = 106,
	add = 107,
	subtract = 109,
	decimalPoint = 110,
	divide = 111,
	f1 = 112,
	f2 = 113,
	f3 = 114,
	f4 = 115,
	f5 = 116,
	f6 = 117,
	f7 = 118,
	f8 = 119,
	f9 = 120,
	f10 = 121,
	f11 = 122,
	f12 = 123,
	numLock = 144,
	scrollLock = 145,
	browserBack = 166,
	browserForward = 167,
	semicolon = 186,
	equal = 187,
	comma = 188,
	dash = 189,
	period = 190,
	forwardSlash = 191,
	graveAccent = 192,
	openBracket = 219,
	backSlash = 220,
	closeBracket = 221,
	singleQuote = 222,
}


// 追加したメソッドだとわかるようにプリフィックスxを付けている

interface Number {
	xToHex(): string;
}

interface String {
	xExpand(): string;
	xFormat(...args: any[]): string;
}

interface HTMLAnchorElement {
	xFolder: SpecialFolder;
}


declare var global: {
	Setting: object;
	SpecialFolders: object;
	ssfDESKTOPDIRECTORY: ShellSpecialFolderConstants.ssfDESKTOPDIRECTORY;
	ssfDRIVES: ShellSpecialFolderConstants.ssfDRIVES;
	ssfPROFILE: ShellSpecialFolderConstants.ssfPROFILE;
	TemporaryFolder: SpecialFolderConst.TemporaryFolder;
	Version: Function;
	window?: Window;
	WScript?: object;
};

declare const Setting: {
	debug: boolean;
	[name: string]: number | boolean;
	
	/** サポートされない古いOS上で実行した時に強制終了させる */
	abortIfOldOS: boolean;
	/** ファイル フォルダー(ディレクトリ)の情報だけを返す */
	fileFolderOnly: boolean;
	/** フォルダーのカテゴリの名前を表示する */
	viewCategory: boolean;
	
	/** 64ビットOSでも32ビット版のmshta.exeで実行する */
	htaFoeceWow64: boolean;
	/** HTAのウインドウの左上隅のX座標。htaTopにも値を指定した場合のみ有効 */
	htaLeft?: number;
	/** HTAのウインドウの左上隅のY座標。htaLeftにも値を指定した場合のみ有効 */
	htaTop?: number;
	/** HTAのウインドウの幅。htaHeightにも値を指定した場合のみ有効 */
	htaWidth?: number;
	/** HTAのウインドウの高さ。htaWidthにも値を指定した場合のみ有効 */
	htaHeight?: number;
	
	/** WSH版で実行環境ではサポートしないフォルダーも出力する */
	wshForceWriteAllData?: boolean;
	/** WSH版でBOMを出力する */
	wshWriteBom?: boolean;
	/** WSH版でFolderData.dirも出力する。デバッグモード限定 */
	wshWriteDir?: boolean;
	/** WSH版でフォルダーの表示名も出力する */
	wshWriteDisplayName?: boolean;
	/** WSH版でフォルダーの種類も出力する */
	wshWriteType?: boolean;
	
	/** @deprecated 移行先はfileFolderOnly */
	directoryOnly?: boolean;
};

declare function getRegValue(name: string, defaultValue:string, expand: true): string;


interface Version {
	major: number;
	minor: number;
	build: number;
	revision: number;
	
	compareTo(value: Version): number;
	equals(obj: Version): boolean;
	isGreaterThan(obj: Version): boolean;
	toString(fieldCount?: number): string;
}

interface VersionConstructor {
	new (major: number, minor: number, build?: number, revision?: number): Version;
	
	prototype: Version;
}

declare const Version: VersionConstructor;


interface SpecialFolder {
	readonly category?: string;
	readonly title: string;
	readonly dir: string | number;
	readonly folderItem: FolderItem;
	readonly path: string;
	readonly isFileFolder: boolean;
	
	execCmd(verb?: string): void;
	execExplorer(verb?: string): void;
	execPowershell(verb?: string): void;
	execWsl(verb?: string): void;
	getType(): string;
	hasProperties(): boolean;
	open(): void;
	showProperties(): void;
}

/** プロパティの表示方法 */
type PropertyTypes = 0 | 1 | 2;

interface SpecialFolderOption {
	/** フォルダーのカテゴリ名 */
	category?: string;
	/** ツール上に表示するフォルダーのパス。FolderItem#Pathを使用したくないときに指定する */
	path?: string;
	/** プロパティの表示方法 */
	propertyType?: PropertyTypes;
	/** プロパティを表示するのに別のFolderItemオブジェクトを使いたい場合に指定する */
	folderItemForProperties?: FolderItem;
}

interface SpecialFolderArgument {
	title: string;
	dir: string | number;
	folderItem: FolderItem;
	path: string;
	option: SpecialFolderOption;
}

interface FolderIteratorIndex {
	current: number;
}

declare const SpecialFolders: {
	item(itemIndex: number): SpecialFolder;
	iterator(): { next: () => IteratorResult<SpecialFolder>; }
}


interface DialogItem {
	id: string;
	caption: string;
	key: string;
	isConsole?: boolean;
	isExtended?: boolean;
	isAlwaysVisible?: boolean;
	isVisible?: boolean;
}

interface DialogArgument {
	items: DialogItem[];
	sendItem(item: string): void;
}


/** @type {ShellSpecialFolderConstants.ssfDESKTOPDIRECTORY} */
declare const ssfDESKTOPDIRECTORY = 16;
/** @type {ShellSpecialFolderConstants.ssfDRIVES} */
declare const ssfDRIVES = 17;
/** @type {ShellSpecialFolderConstants.ssfPROFILE} */
declare const ssfPROFILE = 40;
/** @type {SpecialFolderConst.TemporaryFolder} */
declare const TemporaryFolder = 2;
