//----------------------------------------------------------
// Excel-CSV変換
//
//  概要：
//    xlsxファイルパスをinputに、対象ファイルを展開し、
//    シートごとに出力フォルダパスで指定された場所にcsv出力する
//
//  引数：
//    ・入力xlsxファイルパス
//    ・出力フォルダパス
//
//  リターンコード：
//    ・81：引数不足
//    ・82：入力ファイルパス不正
//    ・83：出力フォルダ作成失敗
//    ・99：その他例外
//----------------------------------------------------------
const XLSX = require("xlsx");
const FS = require("fs");
const PATH = require("path");

// メイン処理呼び出し
main(process.argv);

// --------------------------------------
// メイン処理
//
// 引数：実行時引数
// 戻り値：リターンコード
// --------------------------------------
async function main(args) {
  let check = 0;

  // 引数チェック
  check = CheckArgs(args);
  if (check != 0) {
    return check;
  }

  // 入力ファイルパス
  const inFile = args[2];
  // 出力フォルダパス
  const outPath = args[3];

  // 事前処理
  check = PreProc(outPath);

  // 開始メッセージ
  console.log("Excel-CSV変換処理を開始します。");
  console.log("入力ファイル：" + inFile);

  console.log("...");

  // 変換処理
  const outFiles = Excel2Csv(inFile, outPath);
  if (outFiles == null) {
    return 99;
  }

  // 終了メッセージ
  console.log("Excel-CSV変換処理が完了しました。\n出力ファイル：");
  outFiles.forEach((outFile) => {
    console.log("・" + outFile);
  });

  return 0;
}

// --------------------------------------
// 引数チェック
//
// 引数：実行時引数
// 戻り値：リターンコード
//  81：引数不足、82：入力ファイルパス不正
// --------------------------------------
function CheckArgs(args) {
  if (args.length < 4) {
    console.log(
      "引数が不足しています、第1引数：入力xlsxファイル、第2引数：出力フォルダパス"
    );
    return 81;
  }

  // 入力ファイル存在チェック
  if (!FS.existsSync(args[2])) {
    console.log("入力xlsxファイルが存在しません。");
    return 82;
  }

  return 0;
}

// --------------------------------------
// 事前処理
//
// 引数：出力フォルダパス
// 戻り値：リターンコード
//  83：出力フォルダ作成失敗
// --------------------------------------
function PreProc(outPath) {
  try {
    // 出力フォルダ先フォルダが存在しない場合、作成
    if (!FS.existsSync(outPath)) {
      FS.mkdirSync(outPath);
    }
  } catch (e) {
    console.log("出力フォルダの作成に失敗しました。");
    console.log(e.message);
    return 83;
  }

  return 0;
}

// --------------------------------------
// エクセル-CSV変換
//
// 引数：
// ・入力xlsxファイルパス
// ・出力フォルダパス
// 戻り値：出力ファイル一覧
// --------------------------------------
function Excel2Csv(inFile, outPath) {
  const outFiles = [];

  try {
    // エクセルファイル読み込み
    const excelBook = XLSX.readFile(inFile);

    // エクセルファイル名prefix
    const xlsxFNM = PATH.basename(inFile);
    const prefix = PATH.parse(xlsxFNM).name;

    // エクセルブックに含まれるシートを走査
    excelBook.SheetNames.forEach((sheetName) => {
      // 出力ファイル名
      const outFile = outPath + "\\" + prefix + "_" + sheetName + ".csv";

      // シート
      const sheet = excelBook.Sheets[sheetName];

      // CSVオプションを設定
      const csvOptions = {
        FS: ",",
        RS: "\r\n",
        forceQuotes: true,
      };
      // CSV変換
      const csv = XLSX.utils.sheet_to_csv(sheet, csvOptions);

      // ファイル出力
      FS.writeFileSync(outFile, csv);
      // 出力先を格納
      outFiles.push(outFile);
    });

    return outFiles;
  } catch (e) {
    console.log("Excel-CSV変換に失敗しました。");
    console.log(e.message);
    return null;
  }
}
