/// <reference path="common.js" />

/*
 * ツールの挙動を設定する。
 * このファイルはバージョンアップの時に上書きされるので
 * カスタマイズしたい場合はsettings-custom.jsと言う名前のファイルを作りそちらを編集する。
 */


// ファイル フォルダー(ディレクトリ)の情報だけを返してほしい時は値をtrueにする
Setting.fileFolderOnly = false;

// フォルダーのカテゴリの名前を表示する場合はtrueにする
Setting.viewCategory = false;

// サポートされない古いOS上で実行した時に強制終了させたい時はtrueにする
Setting.abortIfOldOS = false;

// 64ビットOSでも32ビット版のmshta.exeで実行したい時はtrueにする
Setting.htaFoeceWow64 = false;

// HTAのウインドウの幅と高さを指定する
// htaWidthとhtaHeightの両方を指定した場合のみ有効
// Setting.htaWidth = 600;
// Setting.htaHeight = 400;

// HTAのウインドウの左上隅の座標を指定する
// htaLeftとhtaTopの両方を指定した場合のみ有効
// Setting.htaLeft = 100;
// Setting.htaTop = 100;
