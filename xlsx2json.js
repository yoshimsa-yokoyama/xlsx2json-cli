#!/usr/bin/env node

'use strict';

const fs = require('fs');
const path = require('path');
const { Command } = require('commander');
const chalk = require('chalk');
const rimraf = require('rimraf');
const XLSX = require('xlsx');
const merge = require('deepmerge');

const packageJson = require('./package.json');

// ワークブックのシート名一覧
let tgts = [];
// オーバーライド用のオブジェクトのキー配列
let overrideSheetKeys = [];

// デバッグ用
const space = 0;

// エラー文言
const ERROR_TEXT = {
  ON_EXEL_FILE_PARSE: 'エクセルファイルがパースできません。',
  NO_EXEL_FILE_ASSIGN: '指定されたファイルはエクセル（.xlsx）ではありません。',
  NO_EXEL_FILE_EXIST: '指定されたファイルが存在していません。',
  NO_REF_SHEET_EXIST: '参照しているシートが存在しません。',
  NO_MULTI_SHEET_REF: '1つ以上のシートを参照しているか、構文が間違っています。',
  NO_REF_ATTR_IN_OPT: 'シートの参照を行うには$refオプションは必須です。',
  NO_MULTI_OVERRIDE_ROW: 'オーバーライド用シートが持てる値は一行だけです。'
}

// 区切り文字
const HIERARCHY_DELIMITER = '.';
// TODO: 引数化
const REFERENCE_DELIMITER = ':';
// TODO: 引数化
const REF_VAL_DELIMITER = ',';
// TODO: 引数化
const REF_OPT_DELIMITER = '$';
// TODO: 引数化
const TOP_LV_OPT_DELIMITER = '$';
// TODO: 引数化
const OVERRIDE_DELIMITER = '!';
// TODO: 引数化
const OWNKEY_REF_REGEX = /\[[A-Za-z0-9|_]+\]/g;
// TODO: 引数化
const FILENAME_PREFIX = '';

// 生成されたJSONのキャッシュ
const jsonCache = {};
// 参照が不完全なシート名の配列
const incompleteSheetQueue = [];
// writeFile用コールバック
const cb = (err) => {
  if (err) console.log(err);
};

/**********************************************************
 *
 * カンマ区切り半角数字を整数の配列に変換
 *
 **********************************************************/
const convertStrToArr = (str, delimiter) => {
  return str
    .toString()
    .split(delimiter);
};


/**********************************************************
 *
 * 参照されているシートを取得
 *
 **********************************************************/
const getReferredSheetName = (key) => {
  let referredSheetName = '';
  // 参照先シートの数をカウント
  const ptnMatchCount = key.match(new RegExp(`\\${REFERENCE_DELIMITER}`, 'g')).length;
  // 参照時のオプション指定の有無を評価
  const refOption = key.match(new RegExp(`\\${REF_OPT_DELIMITER}`,'g'));
  const refOptionCount = refOption !== null ? refOption.length : 0;

  // 参照先のシート名を取得
  if (ptnMatchCount.length > 1) {
    // 1シート以上の参照があったらエラー
    throw new Error(ERROR_TEXT.NO_MULTI_SHEET_REF);
  } else {
    if (refOptionCount === 0) {
      // NOTE: indexOfだと REFERENCE_DELIMITER も含んでしまうので + 1 処理を行う
      referredSheetName = key.substr(key.indexOf(REFERENCE_DELIMITER) + 1);
    } else {
      // NOTE: オプションが有る場合は REFERENCE_DELIMITER(:) 〜 REF_OPT_DELIMITER($)
      // の間で参照先シート名を取得
      referredSheetName = key.substring(key.indexOf(REFERENCE_DELIMITER) + 1, key.indexOf(REF_OPT_DELIMITER));
    }
  }

  return referredSheetName;
};


/**********************************************************
 *
 * シートのオプションを取得
 *
 **********************************************************/
const getTopLvOptions = key => {
  const topLvOptions = {};
  // 頭のNOTE TOP_LV_OPT_DELIMITER が入ってきてしまうので + 1
  const trimedKey = key.substr(key.indexOf(TOP_LV_OPT_DELIMITER) + 1);

  trimedKey.split(TOP_LV_OPT_DELIMITER).forEach(optstr => {
    const opt = optstr.split('=');
    let topLvOptKey = opt[0];
    let topLvOptVal = opt[1];

    topLvOptions[topLvOptKey] = topLvOptVal;
  });

  return topLvOptions;
};


/**********************************************************
 *
 * シート参照時のオプションを取得
 *
 **********************************************************/
const getReferOptions = (key, record) => {
  const referOptions = {};
  // 頭のNOTE REF_OPT_DELIMITERが入ってきてしまうので + 1
  const trimedKey = key.substr(key.indexOf(REF_OPT_DELIMITER) + 1);

  trimedKey.split(REF_OPT_DELIMITER).forEach(optstr => {
    const opt = optstr.split('=');
    let refOptKey = opt[0];
    let refOptVal = opt[1];

    // オプジョンの値のなかの[]がついているものを、
    // 参照元のオブジェクトの値で書き換える
    const ownKeyMatch = refOptVal.match(OWNKEY_REF_REGEX);
    if (ownKeyMatch !== null) {
      ownKeyMatch.forEach(itm => {
        refOptVal = refOptVal.replace(itm, record[itm.replace(/\[|\]/g, '')]);
      });
    }

    referOptions[refOptKey] = refOptVal;
  });

  // 必須パラメータのバリデーション
  if (referOptions.ref === undefined) {
    throw new Error(ERROR_TEXT.NO_REF_ATTR_IN_OPT);
  }

  return referOptions;
};


/**********************************************************
 *
 * 階層オブジェクトを作成
 *
 **********************************************************/
const createHierarchy = (originKey, singleRecord) => {
  // 階層構造のキーを取得
  let hierarchyKeys = convertStrToArr(originKey, HIERARCHY_DELIMITER);
  // 配列をreverseする前に新しくキーとなる値を保存
  const newOriginKey = hierarchyKeys[0];

  // 階層構造のキーの配列を深層からreduceして
  // 階層構造オブジェクトを作成
  const obj = hierarchyKeys.reverse().reduce((acc,cur) => {
    let vessel = {};

    vessel[cur] = acc;
    return vessel;
  }, singleRecord[originKey]);

  // すでに階層オブジェクトが存在した場合はマージ
  if (singleRecord.hasOwnProperty(newOriginKey)) {
    singleRecord[newOriginKey] = merge(singleRecord[newOriginKey], obj[newOriginKey]);
  } else {
    singleRecord[newOriginKey] = obj[newOriginKey];
  }

  // 階層構造化前のキーは削除
  if (singleRecord.hasOwnProperty(originKey)) {
    delete singleRecord[originKey];
  }

  return singleRecord;
}


/**********************************************************
 *
 * 他シートへの参照からオブジェクトを作成
 *
 **********************************************************/
const importReference = (originKey, singleRecord, referringSheetName) => {
  const newOriginKey = originKey.substr(0, originKey.indexOf(REFERENCE_DELIMITER));
  const referredSheetName = getReferredSheetName(originKey);
  const referOptions = getReferOptions(originKey, singleRecord);

  // 参照しているシートが存在しているかを確認
  if (!tgts.includes(referredSheetName)) {
    throw new Error(ERROR_TEXT.NO_REF_SHEET_EXIST)
  } else if (
    jsonCache.hasOwnProperty(referredSheetName) &&
    jsonCache[referredSheetName].hasOwnProperty('data'))
  {
    // すでにJSONが生成されている場合
    // 突合する値を参照元から取得
    let refKeys = convertStrToArr(singleRecord[originKey], REF_VAL_DELIMITER);

    // ref_prefix か ref_suffix が存在している場合は ref に付与
    // OWNKEY_REF_REGEX( [A-Za-z0-9|_]+ )を自分のオブジェクトの該当の値で置き換える
    refKeys = refKeys.map( key => {
      if (referOptions.hasOwnProperty('ref_prefix')) key = `${referOptions.ref_prefix}${key}`;
      if (referOptions.hasOwnProperty('ref_suffix')) key = `${key}${referOptions.ref_suffix}`;

      const ownKeyMatch = key.match(OWNKEY_REF_REGEX);
      if (ownKeyMatch !== null) {
        ownKeyMatch.forEach(itm => {
          // []を除いた値に参照するキーを差し替え
          key = key.replace(itm, singleRecord[itm.replace(/\[|\]/g, '')]);
        });
      }
      return key;
    });

    let newValue = jsonCache[referredSheetName].data.reduce((acc,cur,idx,src) => {
      // NOTE: refKeysが文字列の配列なので、toString()
      // TODO: 現在の参照は配列にしか対応していない

      // deepcopyの作成
      const tmpCur = JSON.parse(JSON.stringify(cur));

      // valueから不可視キー（__*)を削除
      Object.keys(tmpCur).forEach(key => {
        if (/^(__)[A-Za-z0-9_]*$/.test(key)) {
          delete tmpCur[key];
        }
      });

      if (refKeys.includes(cur[referOptions.ref].toString())) {
        acc.push(tmpCur);
      }
      return acc;
    }, []);

    singleRecord[newOriginKey] = newValue;

    // 参照前のキーは削除
    if (singleRecord.hasOwnProperty(originKey)) {
      delete singleRecord[originKey];
    }
    console.log(`[ ${referringSheetName}.${newOriginKey} ]への[ ${referredSheetName} ] の紐付けに成功しました...`);
  } else {
    console.log(`[ ${referredSheetName} ] がまだ存在してないため、[ ${originKey} ]のパースに失敗しました...`);
    // シートがまだ無かった場合、後で紐付けをし直すために
    // シート名と該当のキーを保存
    if (!incompleteSheetQueue.includes(`${referringSheetName},${originKey}`)) {
      incompleteSheetQueue.push(`${referringSheetName},${originKey}`);
    }
  }

  return singleRecord;
}


/**********************************************************
 *
 * ワークシートのメイン処理
 *
 **********************************************************/
const xlsx2Json = (xlsxFilePath, outdir = path.resolve('./')) => {
  let workbook;

  console.log(`${chalk.white.bgBlue.bold(`output directory: ${outdir}`)}`);

  // エクセルファイルをJSONへ書き出し
  try {
    workbook = XLSX.readFile(xlsxFilePath);
  } catch (e) {
    // 何らかの原因でエクセルの展開に失敗した場合はエラー
    throw new Error(ERROR_TEXT.ON_EXEL_FILE_PARSE);
    console.log(e);
  }

  // シート名を格納
  tgts = workbook.SheetNames;

  overrideSheetKeys = tgts.filter(itm => /^\![A-Za-z_]*/.test(itm))

  // シートを一枚ずつ取得して処理
  tgts.forEach((tgt) => {
    // ワークシートの内容を取得してJSONへ変換
    let recordArray = XLSX.utils.sheet_to_json(workbook.Sheets[tgt]);

    console.log('')
    console.log('')
    console.log('======================================');
    console.log(`シート[ ${tgt} ]の処理を開始`)
    console.log('======================================');
    console.log('')

    // レコードを一行ずつ取得して処理
    recordArray = recordArray.map(record => {

      // 1レコードのキーを順番に処理
      Object.keys(record).forEach(key => {
        // keyの書式から作成するオブジェクトの種類を取得
        const hasHierarchy = key.includes(HIERARCHY_DELIMITER);
        const hasReference = key.includes(REFERENCE_DELIMITER);

        console.log(`key >>> [ ${key} ]`);
        console.log('hierarchy >>> ', hasHierarchy);
        console.log('reference >>> ', hasReference);

        if (hasHierarchy) {
          // 階層構造を持っていた場合
          record = createHierarchy(key, record);
        } else if (hasReference) {
          // 参照を持っていた場合
          record = importReference(key, record, tgt);
        }
      });

      return record;
    });

    jsonCache[tgt] = {
      data: recordArray
    };

    console.log('');
    console.log('');
    console.log(jsonCache[tgt]);
    console.log('');
    console.log('');
    console.log('======================================');
  });

  // 参照が完了してないデータがある場合、
  // 再度インポートを試みる
  if(incompleteSheetQueue.length > 0) {
    incompleteSheetQueue.forEach(refCode => {

      const sheetName = refCode.split(',')[0];
      const keyName = refCode.split(',')[1];

      console.log('')
      console.log('完了してないシートがあります... ', sheetName);
      console.log('対象のキーは以下です... ', keyName);
      console.log('生成後のデータから再度紐付けを行います...');

      if (
        jsonCache.hasOwnProperty(sheetName) &&
        jsonCache[sheetName].hasOwnProperty('data')
      ) {
        // NOTE: jsonCacheの内容を直接自身の変更で上書きするとエラーになるので、
        //       一度JSON.stringifyでDeepCopyを行う
        const tmpJsonData = JSON.stringify(jsonCache[sheetName], undefined, space);

        jsonCache[sheetName].data = JSON.parse(tmpJsonData).data.map(record => {
          // 一度作り終わったデータから参照を紐付け直す
          record = importReference(keyName, record, sheetName);
          return record;
        });
      }
    })
  }

  // オーバーライド用シートがあったらマージ
  if (overrideSheetKeys.length > 0) {
    overrideSheetKeys.forEach(key => {
      // オーバーライド用シートの構造チェック
      if (jsonCache[key].data.length > 1) {
        throw new Error(ERROR_TEXT.NO_MULTI_OVERRIDE_ROW);
      }
      // オーバーライド用シート名から結合するシートのキーを取得
      // NOTE: オプションが付いているシート名は単純にオーバーライド用
      //       シート名をキーとして参照できないため
      const tgtKey = tgts.find(tgt => new RegExp(`^${key.substr(1)}\\$`).test(tgt));

      // オーバーライド用シートのデータにオーバーライド先のシートのデータをマージ
      jsonCache[tgtKey] = merge(jsonCache[key].data[0], jsonCache[tgtKey]);
    });
  }

  // キャッシュした完成データをファイルに出力
  tgts.forEach(tgt => {
    // __*と!*のシートは出力はしない
    if (!/^(__)[A-Za-z0-9_]*$|^\![A-Za-z0-9_]*/.test(tgt)) {

      const topLvOption = tgt.match(new RegExp(`\\${TOP_LV_OPT_DELIMITER}`,'g'));
      const topLvOptionCount = topLvOption !== null ? topLvOption.length : 0;

      if (topLvOptionCount > 0) {
        // シートにオプションが付いていた場合
        const trimedTgt = tgt.substr(0, tgt.indexOf(TOP_LV_OPT_DELIMITER));
        const topLvOptions = getTopLvOptions(tgt);

        // トップレベルキーの置き換え
        // NOTE: jsonCacheにはシート名そのままをいれおき、最後に書き出すときに
        //       ファイル名を特定・トップレベルキーも置き換える
        if (topLvOptions.hasOwnProperty('key')) {
          const tmpObj = {};

          Object.keys(jsonCache[tgt]).forEach(itm => {
            if (itm === 'data') {
              tmpObj[topLvOptions.key] = jsonCache[tgt].data;
            } else {
              tmpObj[itm] = jsonCache[tgt][itm];
            }
          });

          fs.writeFile(
            `${outdir}/${FILENAME_PREFIX}${trimedTgt}.json`,
            JSON.stringify(tmpObj, undefined, space),
            cb
          );
        }
      } else {
        fs.writeFile(
          `${outdir}/${FILENAME_PREFIX}${tgt}.json`,
          JSON.stringify(jsonCache[tgt], undefined, space),
          cb
        );
      }
    }
  });

  console.log('cached: ', jsonCache);

  console.log('')
  console.log('======================================');
  console.log('JSONの生成が完了しました！');
  console.log('======================================');
}


/**********************************************************
 *
 * コマンド開始
 *
 **********************************************************/
const init = () => {
  const program = new Command(packageJson.name)
    .version(packageJson.version)
    .option('-o, --outdir <outdir>', 'JSON output directory')
    .parse(process.argv);

  // コマンド名の表示
  console.log(`${chalk.white.bgBlue.bold(`HELLO ${program.name()}`)}`);

  // 指定ファイルの評価(拡張子で評価）
  if (
    program.args.length > 0 &&
    /^(\.\/)?[^\.]+\.(xlsx)$/.test(program.args[0])
  ) {
    const rootDir = path.resolve('./');
    const xlsxFilePath = `${rootDir}/${program.args[0]}`;

    // ファイルの存在を確認
    try {
      if (fs.existsSync(xlsxFilePath)) {

        // オプションを解析
        const options = program.opts()

        if (options.outdir !== undefined) console.log(`${chalk.white.bgBlue.bold(`output dir: ${options.outdir}`)}`);

        // 実行
        xlsx2Json(xlsxFilePath, options.outdir);
      }
    } catch(err) {
      console.error(err)
      // エクセルファイルが存在しなかった場合はエラー
      throw new Error(ERROR_TEXT.NO_EXEL_FILE_EXIST);
    }
  } else {
    // エクセルファイルが指定されなかった場合はエラー
    throw new Error(ERROR_TEXT.NO_EXEL_FILE_ASSIGN);
  }
}

init();
