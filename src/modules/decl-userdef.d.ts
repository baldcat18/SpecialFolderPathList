/// <reference path="decl-lib.d.ts" />

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


/** グローバル変数参照用 */
declare var G: {
	Setting: object;
	Version: Function;
	window?: Window;
	WScript?: typeof WScript;
	SpecialFolders: object;
	ssfPROFILE: number;
	TemporaryFolder: number;
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

declare function getRegValue(name: string, defaultValue: string, expand: true): string;


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
	new(major: number, minor: number, build?: number, revision?: number): Version;

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
	dir: string;
	folderItem: FolderItem;
	path: string;
	option: SpecialFolderOption;
}

interface FolderIteratorIndex {
	current: number;
}

declare const SpecialFolders: {
	item(itemIndex: number): SpecialFolder;
	iterator(): { next: () => IteratorResult<SpecialFolder, undefined>; }
}


interface DialogItem {
	id: string;
	caption: string;
	isConsole?: boolean;
	isExtended?: boolean;
	isAlwaysVisible?: boolean;
	isVisible?: boolean;
	key?: number;
}

interface DialogArgument {
	items: DialogItem[];
	sendItem(item: string): void;
}


declare const ssfPROFILE: ShellSpecialFolderConstants.ssfPROFILE;
declare const TemporaryFolder: SpecialFolderConst.TemporaryFolder;
