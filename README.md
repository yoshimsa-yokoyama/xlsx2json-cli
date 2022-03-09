# @asamihsoy/xlsx2json-cli

※ will be translated into English in near future

## インストール・起動方法

``` bash
npm i -D @asamihsoy/xlsx2json-cli
npx xlsx2json-cli ./sample.xlsx
```

※ npm scriptには `xlsx2json-cli ./sample.xlsx -o ./assets/src/` という形で指定できます。

## 基本の使い方

xlsxファイルを用意して `xlsx2json-cli` を実行することで、JSONへ変換できます。

##### 【INPUT】ファイル名: rostar.xlsx / シート名: member

|id|name              |company          |department|
|--|------------------|-----------------|----------|
|1 |@asamihsoy|fugafuga Inc     |division1 |
|2 |@person2          |hogehoge Inc     |division2 |

↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓

```
npx xlsx2json-cli ./member.xlsx
```

↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓

##### 【OUTPUT】ファイル名: member.json
```
{
"data": [{
    "id": 1,
    "name": "@asamihsoy",
    "company": "fugafuga Inc",
    "department": "division1"
  },{
    "id": 2,
    "name": "@person2",
    "company": "hogehoge Inc",
    "department": "division2"
  }]
}
```

---

## コマンドライン使用時のオプション確認

```
npx xlsx2json-cli -h
```

|フラグ|備考                                          |
|------|----------------------------------------------|
|`-o`  |実行場所からJSONの格納ディレクトリへの相対パス|

## その他の使い方

シート名や見出しを一定の構文で記述することで、JSONのアウトプットを操作することができます。

|アクション                      |設定箇所|設定構文                    |設定値                |説明                                                                                                                                                                               |
|--------------------------------|--------|----------------------------|----------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|JSONのトップレベルキーの変更    |シート名|`sheetName$key=altKey`      |任意の文字列          |JSONのトップレベルのキーを変更できます。（※デフォルト値は `data` です。）                                                                                                          |
|JSONのトップレベルへのキーの追加|シート名|`!sheetName`                |マージしたいシート名  |`!`を頭につけたシートはJSON化されず、!のあとのsheetNameと同じ名前のシートのJSONのTOPレベルにマージされます。※ そのため、シート内のレコードは1行しか反映されません。                |
|参照用途限定のシートを作成      |シート名|`__sheetName`               |参照用にしたいシート名|`__`を頭につけたシートはJSON化されず、参照用のみに使うことができます。                                                                                                             |
|他シートの展開                  |見出し  |`jsonKey:sheetName1`        |シート名              |JSONの出力キーに`:`でシート名をつなぐことで、該当シートから生成されるJSONのキー（jsonKey）の中にシート1（sheetName1）の値を展開できます。 ※1キーに紐付けられるのは1シートのみです。|
|ネストの作成                    |見出し  |`jsonKey1.jsonKey2.jsonKey3`|キー名                |JSONの出力キーを.でつなげることで階層構造を作成可能です。かぶるキーがある場合はオブジェクトはマージされます。                                                                      |

※ 機能は適宜追加していきます。
※ シート名／キー名は半角英数字のみサポートしています。
※ `!`と`__`が頭についたシートはJSONとして出力されません(シート名は半角英数字である必要があります。)
