# CRUD図 解析ツール

## 1. 目的

システムで利用しているデータベースのテーブル・項目と、各機能（Oracle PL/SQL・VB.NET DAC）で使用しているCRUD操作との相関図（マトリックス）を自動生成する。

## 2. システム構成

### 2.1 データベース

| 項目             | 内容                                                                |
| ---------------- | ------------------------------------------------------------------- |
| RDBMS            | Oracle 19c                                                          |
| 対象オブジェクト | Package, Trigger, View, Materialized View, Table, Synonym, Sequence |

### 2.2 アプリケーション構成

```
VB.NET アプリケーション
├── PR層（プレゼンテーション層）
│   └── 画面・UI
└── BC層（ビジネスロジック層）
    ├── ビジネスロジック
    └── DAC（データアクセス層） ← CRUD解析対象（SQL埋め込み）
        └── *dac*.vb ファイル

Oracle データベース
├── Table（テーブル定義）       ← DDL解析対象
├── Index（インデックス定義）   ← DDL解析対象
├── Package Body               ← CRUD解析対象（ストアドプロシージャ）
├── Function（スタンドアロン）  ← CRUD解析対象
├── Procedure（スタンドアロン） ← CRUD解析対象
├── Trigger                    ← CRUD解析対象
├── View                       ← CRUD解析対象
├── Materialized View          ← CRUD解析対象
├── Sequence                   ← 対象外
└── Synonym                    ← 対象外
```

### 2.3 ソース管理

| ソース種別              | 管理方法              | ファイル拡張子 |
| ----------------------- | --------------------- | -------------- |
| Oracle テーブル定義     | SVN                   | .sql           |
| Oracle インデックス定義 | SVN                   | .sql           |
| Oracle パッケージ等     | SVN                   | .sql           |
| Oracle Function         | SVN                   | .sql           |
| Oracle Procedure        | SVN                   | .sql           |
| VB.NET ソース           | SVN（ローカルコピー） | .vb            |

## 3. ツール概要

### 3.1 解析対象

| 解析対象        | ソース種別    | 対象ファイル                         | 抽出内容                             |
| --------------- | ------------- | ------------------------------------ | ------------------------------------ |
| テーブル定義DDL | .sql ファイル | CREATE TABLE 文                      | テーブル名、カラム名、データ型、制約 |
| インデックスDDL | .sql ファイル | CREATE INDEX 文                      | インデックス名、テーブル名、カラム名 |
| Oracle PL/SQL   | .sql ファイル | Package Body, Function, Procedure 等 | INSERT/SELECT/UPDATE/DELETE/MERGE 文 |
| Oracle Trigger  | .sql ファイル | Trigger                              | INSERT/SELECT/UPDATE/DELETE/MERGE 文 |
| Oracle View     | .sql ファイル | View, Materialized View              | INSERT/SELECT/UPDATE/DELETE/MERGE 文 |
| VB.NET DAC      | .vb ファイル  | ファイル名に `dac` を含むもの        | 埋め込みSQL文                        |

### 3.2 解析フロー

```
[テーブル定義DDL]    ──→ Parse-TableDdl.ps1  ──→ テーブル定義/インデックス定義
[インデックス定義DDL]──→ Parse-TableDdl.ps1  ──┘          │
                                                           ↓ SELECT * 展開 / 未使用カラム検出
[Oracle SQLファイル]  ──→ Parse-OracleSql.ps1 ──┐          │
                                                 ├──→ [統合] ──→ Export-CrudExcel.ps1 ──→ CrudMatrix.xlsx
[VB.NET DACファイル]  ──→ Parse-VbNetDac.ps1  ──┘          │
                                                            └──→ crud_results.json（中間データ）
```

**実行フロー詳細:**

1. **設定読み込み**: `config.psd1` からパス・パターン・除外条件を読み込み
2. **Oracle SQL 解析**: Package Body / Function / Procedure / Trigger / View の DML文を抽出
3. **VB.NET DAC 解析**: `*dac*.vb` ファイルを走査し、埋め込みSQL文字列を抽出
4. **DDL 解析**: CREATE TABLE / CREATE INDEX 文からテーブル定義・インデックス定義を抽出
5. **SELECT \* 展開**: テーブル定義を使い `SELECT *` を個別カラムに展開
6. **未使用カラム検出**: テーブル定義にあるがCRUD結果に登場しないカラムを検出
7. **結果統合**: Oracle・VB.NETの結果をマージ
8. **JSON出力**: 中間データとしてJSON形式で保存
9. **Excel出力**: CRUDマトリックス＋テーブル定義＋インデックス定義をExcel出力

### 3.3 抽出ロジック

#### Oracle PL/SQL 解析

| SQL種別                             | 抽出項目             | CRUD区分           |
| ----------------------------------- | -------------------- | ------------------ |
| INSERT INTO table (col1, col2, ...) | テーブル名、カラム名 | C (Create)         |
| SELECT col1, col2 FROM table        | テーブル名、カラム名 | R (Read)           |
| UPDATE table SET col1 = ...         | テーブル名、カラム名 | U (Update)         |
| DELETE FROM table                   | テーブル名           | D (Delete)         |
| MERGE INTO table                    | テーブル名           | CU (Create/Update) |

#### VB.NET DAC 解析

以下のパターンで埋め込みSQLを検出:

| パターン                               | 説明                       |
| -------------------------------------- | -------------------------- |
| `Dim sql As String = "SELECT ..."`     | 単一行の文字列リテラル     |
| `"SELECT " & "col1 " & "FROM ..."`     | `&` 演算子による文字列連結 |
| `sb.Append("SELECT ...")`              | StringBuilder パターン     |
| `"SELECT ..." _` + 改行 + `"FROM ..."` | VB.NET 行継続文字 `_`      |

## 4. CRUDマトリックス仕様

### 4.1 凡例

| 記号 | 意味           | 対応SQL |
| ---- | -------------- | ------- |
| C    | Create（登録） | INSERT  |
| R    | Read（参照）   | SELECT  |
| U    | Update（更新） | UPDATE  |
| D    | Delete（削除） | DELETE  |
| CU   | Create/Update  | MERGE   |
| -    | 該当なし       | -       |

### 4.2 出力Excelシート構成（全6シート）

#### シート1: テーブル×機能サマリー

テーブル単位で各機能のCRUD操作を一覧化する。

| テーブル名       | PKG:受注PKG.登録 | PKG:受注PKG.検索 | FNC:GET_PRICE.(MAIN) | VB:OrderDac.Insert |
| ---------------- | :--------------: | :--------------: | :------------------: | :----------------: |
| ORDER_TBL        |        C         |        R         |          -           |         C          |
| ORDER_DETAIL_TBL |        C         |        R         |          -           |         C          |
| CUSTOMER_TBL     |        -         |        R         |          -           |         -          |
| STOCK_TBL        |        U         |        -         |          R           |         U          |

#### シート2: 項目別詳細

テーブル×項目×機能の詳細CRUDマトリックス。`SELECT *` はテーブル定義から個別カラムに展開される。

| テーブル名   | 項目名        | PKG:受注PKG.登録 | PKG:受注PKG.検索 | VB:OrderDac.Insert |
| ------------ | ------------- | :--------------: | :--------------: | :----------------: |
| ORDER_TBL    | ORDER_ID      |        C         |        R         |         C          |
| ORDER_TBL    | ORDER_DATE    |        C         |        R         |         C          |
| ORDER_TBL    | CUSTOMER_ID   |        C         |        R         |         C          |
| ORDER_TBL    | STATUS        |        C         |        -         |         C          |
| CUSTOMER_TBL | CUSTOMER_NAME |        -         |        R         |         -          |

#### シート3: テーブル定義

CREATE TABLE DDLから抽出したテーブル・カラム定義の一覧。

| テーブル名 | No  | カラム名    | データ型     | NULL許可 | DEFAULT | ソースファイル |
| ---------- | --- | ----------- | ------------ | -------- | ------- | -------------- |
| ORDER_TBL  | 1   | ORDER_ID    | NUMBER(10)   | NOT NULL | NO      | order_tbl.sql  |
| ORDER_TBL  | 2   | ORDER_DATE  | DATE         | NOT NULL | NO      | order_tbl.sql  |
| ORDER_TBL  | 3   | CUSTOMER_ID | NUMBER(10)   | NOT NULL | NO      | order_tbl.sql  |
| ORDER_TBL  | 4   | STATUS      | VARCHAR2(20) | NULL     | YES     | order_tbl.sql  |

#### シート4: インデックス定義

CREATE INDEX DDLから抽出したインデックス定義の一覧。

| テーブル名 | インデックス名 | 一意性    | カラム位置 | カラム名    | ソースファイル |
| ---------- | -------------- | --------- | ---------- | ----------- | -------------- |
| ORDER_TBL  | PK_ORDER       | UNIQUE    | 1          | ORDER_ID    | order_idx.sql  |
| ORDER_TBL  | IDX_ORDER_DATE | NONUNIQUE | 1          | ORDER_DATE  | order_idx.sql  |
| ORDER_TBL  | IDX_ORDER_CUST | NONUNIQUE | 1          | CUSTOMER_ID | order_idx.sql  |

#### シート5: 未使用カラム

テーブル定義に存在するが、どのCRUD操作からも参照されていないカラム一覧。

| テーブル名 | カラム名    | データ型      | NULL許可 |
| ---------- | ----------- | ------------- | -------- |
| ORDER_TBL  | REMARKS     | VARCHAR2(500) | NULL     |
| ORDER_TBL  | DELETE_FLAG | CHAR(1)       | NULL     |

#### シート6: 生データ

全解析結果のフラットな一覧。フィルタリング・再集計に利用可能。

| ソース種別 | ソースファイル | オブジェクト種別 | オブジェクト名 | プロシージャ/メソッド | 機能名                         | テーブル名 | 項目名   | 操作 |
| ---------- | -------------- | ---------------- | -------------- | --------------------- | ------------------------------ | ---------- | -------- | ---- |
| Oracle     | order_pkg.sql  | PACKAGE          | ORDER_PKG      | INSERT_ORDER          | PACKAGE:ORDER_PKG.INSERT_ORDER | ORDER_TBL  | ORDER_ID | C    |
| Oracle     | get_price.sql  | FUNCTION         | GET_PRICE      | (MAIN)                | FUNCTION:GET_PRICE.(MAIN)      | STOCK_TBL  | PRICE    | R    |
| VB.NET     | OrderDac.vb    | DAC              | OrderDac       | InsertOrder           | VB:OrderDac.InsertOrder        | ORDER_TBL  | ORDER_ID | C    |

## 5. ディレクトリ構成

```
CrudMatrixAnalyzer/
├── Readme.md                          ← 本ドキュメント
└── scripts/
    ├── config.psd1                    ← 設定ファイル
    ├── Run-CrudAnalysis.ps1           ← メイン実行スクリプト
    ├── Parse-OracleSql.ps1            ← Oracle SQL解析（Package/Function/Procedure/Trigger/View）
    ├── Parse-VbNetDac.ps1             ← VB.NET DAC解析
    ├── Parse-TableDdl.ps1             ← テーブル定義・インデックス定義解析
    ├── Export-CrudExcel.ps1           ← Excel出力（6シート構成）
    └── output/                        ← 出力ディレクトリ（自動生成）
        ├── crud_results.json          ← 中間データ（JSON）
        └── CrudMatrix.xlsx            ← CRUDマトリックス（Excel）
```

## 6. スクリプト仕様

### 6.1 config.psd1（設定ファイル）

PowerShell Data File 形式の設定ファイル。実行前にパスを環境に合わせて修正する。

| セクション | キー             | 説明                                  | 設定例                                                  |
| ---------- | ---------------- | ------------------------------------- | ------------------------------------------------------- |
| Oracle     | SourcePath       | Oracle SQLファイルのディレクトリパス  | `C:\SVN\Oracle\packages`                                |
| Oracle     | FilePattern      | 対象ファイルパターン                  | `*.sql`                                                 |
| Oracle     | ExcludePatterns  | 除外ファイルパターン                  | `@("*_bak*", "*_old*")`                                 |
| Oracle     | ObjectTypes      | 対象オブジェクト種別                  | `@("PACKAGE BODY", "TRIGGER", "FUNCTION", "PROCEDURE")` |
| Ddl        | TableSourcePath  | テーブル定義DDLのディレクトリパス     | `C:\SVN\Oracle\tables`                                  |
| Ddl        | IndexSourcePath  | インデックス定義DDLのディレクトリパス | `C:\SVN\Oracle\indexes`                                 |
| Ddl        | FilePattern      | 対象ファイルパターン                  | `*.sql`                                                 |
| Ddl        | ExcludePatterns  | 除外ファイルパターン                  | `@("*_bak*", "*_old*")`                                 |
| Ddl        | ExpandSelectStar | SELECT * をカラム展開するか           | `$true`                                                 |
| VbNet      | SourcePath       | VB.NETソースのディレクトリパス        | `C:\SVN\VbNet\src`                                      |
| VbNet      | DacFilePattern   | DACファイルのパターン                 | `*dac*.vb`                                              |
| VbNet      | ExcludePatterns  | 除外ファイルパターン                  | `@("*.Designer.vb")`                                    |
| Output     | JsonPath         | JSON出力パス                          | `.\output\crud_results.json`                            |
| Output     | ExcelPath        | Excel出力パス                         | `.\output\CrudMatrix.xlsx`                              |
| Output     | SummarySheetName | サマリーシート名                      | `テーブル×機能サマリー`                                 |
| Output     | DetailSheetName  | 詳細シート名                          | `項目別詳細`                                            |
| -          | ExcludeTables    | 除外テーブル名                        | `@("DUAL", "PLAN_TABLE")`                               |
| -          | ExcludeSchemas   | 除外スキーマプレフィックス            | `@("SYS", "SYSTEM")`                                    |

### 6.2 Run-CrudAnalysis.ps1（メイン実行）

| パラメータ | 型     | 必須 | デフォルト      | 説明                                                |
| ---------- | ------ | ---- | --------------- | --------------------------------------------------- |
| ConfigPath | string | ×    | `.\config.psd1` | 設定ファイルパス                                    |
| ExportMode | string | ×    | `Module`        | Excel出力方式（`Module` / `COM`）                   |
| SkipOracle | switch | ×    | $false          | Oracle SQL解析をスキップ                            |
| SkipVbNet  | switch | ×    | $false          | VB.NET DAC解析をスキップ                            |
| SkipDdl    | switch | ×    | $false          | DDL（テーブル定義・インデックス定義）解析をスキップ |

### 6.3 Parse-OracleSql.ps1（Oracle SQL解析）

**主要関数:**

| 関数名                     | 説明                                                           |
| -------------------------- | -------------------------------------------------------------- |
| `Parse-OracleSqlDirectory` | ディレクトリ内の全SQLファイルを一括解析                        |
| `Parse-OracleSqlFile`      | 単一SQLファイルを解析                                          |
| `Get-OracleObjectInfo`     | パッケージ名・オブジェクト種別を抽出（Function/Procedure対応） |
| `Get-ProcedureBlocks`      | プロシージャ/ファンクション単位にブロック分割                  |
| `Extract-PackageBodySection` | PACKAGE BODY部分のみを抽出（SPEC+BODY同一ファイル対応）      |
| `Extract-TableAndColumns`  | SQL文からテーブル名・カラム名を抽出                            |
| `Remove-SqlComments`       | SQLコメント（-- と /* */）を除去                               |

**解析データ構造:**

```
@{
    SourceType  = "Oracle"           # ソース種別
    SourceFile  = "order_pkg.sql"    # ファイル名
    ObjectType  = "PACKAGE"          # オブジェクト種別
    ObjectName  = "ORDER_PKG"        # オブジェクト名
    ProcName    = "INSERT_ORDER"     # プロシージャ名
    FeatureName = "PACKAGE:ORDER_PKG.INSERT_ORDER"  # 機能名（マトリックスの列名）
    TableName   = "ORDER_TBL"        # テーブル名
    ColumnName  = "ORDER_ID"         # 項目名
    Operation   = "C"                # CRUD操作
}
```

### 6.4 Parse-TableDdl.ps1（テーブル定義・インデックス定義解析）

**主要関数:**

| 関数名                      | 説明                                                        |
| --------------------------- | ----------------------------------------------------------- |
| `Remove-SqlCommentsForDdl`  | SQLコメント除去（DDL用、単体実行対応）                      |
| `Parse-TableDdlDirectory`   | ディレクトリ内の全DDLファイルを一括解析                     |
| `Parse-CreateTable`         | CREATE TABLE 文からテーブル名・カラム名・データ型を抽出     |
| `Parse-CreateIndex`         | CREATE INDEX 文からインデックス名・テーブル名・カラムを抽出 |
| `Extract-ColumnDefinitions` | テーブル定義のカラム部分を個別に解析                        |
| `Expand-SelectStar`         | テーブル定義を使い SELECT * を個別カラムに展開              |
| `Find-UnusedColumns`        | テーブル定義にあるがCRUD結果にないカラムを検出              |

**テーブル定義データ構造:**

```
@{
    Schema      = "MYSCHEMA"         # スキーマ名
    TableName   = "ORDER_TBL"        # テーブル名
    ColumnName  = "ORDER_ID"         # カラム名
    DataType    = "NUMBER(10)"       # データ型
    Nullable    = "NOT NULL"         # NULL許可
    HasDefault  = "NO"               # DEFAULT有無
    OrdinalPos  = 1                  # カラム順序
    SourceFile  = "order_tbl.sql"    # DDLファイル名
}
```

**インデックス定義データ構造:**

```
@{
    IndexSchema = "MYSCHEMA"         # インデックススキーマ
    IndexName   = "PK_ORDER"         # インデックス名
    TableSchema = "MYSCHEMA"         # テーブルスキーマ
    TableName   = "ORDER_TBL"        # テーブル名
    ColumnName  = "ORDER_ID"         # カラム名
    ColumnPos   = 1                  # カラム位置
    Uniqueness  = "UNIQUE"           # 一意性（UNIQUE/NONUNIQUE）
    SourceFile  = "order_idx.sql"    # DDLファイル名
}
```

### 6.5 Parse-VbNetDac.ps1（VB.NET DAC解析）

**主要関数:**

| 関数名                     | 説明                                     |
| -------------------------- | ---------------------------------------- |
| `Assert-SqlParserLoaded`   | Extract-TableAndColumns の依存チェック   |
| `Parse-VbNetDacDirectory`  | ディレクトリ内の全DACファイルを一括解析  |
| `Parse-VbNetDacFile`       | 単一DACファイルを解析                    |
| `Extract-VbNetSqlStrings`  | VB.NETコードから埋め込みSQL文字列を抽出  |
| `Get-VbNetClassAndMethods` | クラス名・メソッド名を抽出しブロック分割 |

**埋め込みSQL検出パターン:**

| #   | パターン      | VB.NETコード例                               |
| --- | ------------- | -------------------------------------------- |
| 1   | 単一行文字列  | `Dim sql As String = "SELECT col1 FROM tbl"` |
| 2   | 文字列連結    | `"SELECT " & "col1 " & "FROM tbl"`           |
| 3   | StringBuilder | `sb.Append("SELECT col1 FROM tbl")`          |
| 4   | 行継続文字    | `"SELECT col1 " _ ` + 改行 + `& "FROM tbl"`  |

### 6.6 Export-CrudExcel.ps1（Excel出力）

**主要関数:**

| 関数名                       | 説明                                          |
| ---------------------------- | --------------------------------------------- |
| `Export-CrudExcelWithModule` | ImportExcel モジュールを使用したExcel出力     |
| `Export-CrudExcelWithCOM`    | Excel COM オートメーションを使用したExcel出力 |
| `Export-CrudJson`            | 中間データをJSON形式で出力                    |
| `Build-CrudSummaryMatrix`    | テーブル×機能のサマリーマトリックスを構築     |
| `Build-CrudDetailMatrix`     | テーブル×項目×機能の詳細マトリックスを構築    |
| `Build-TableDefSheet`        | テーブル定義シートを構築                      |
| `Build-IndexDefSheet`        | インデックス定義シートを構築                  |
| `Build-UnusedColumnsSheet`   | 未使用カラムシートを構築                      |
| `Build-RawDataSheet`         | 生データシートを構築                          |

**Excel出力方式:**

| 方式           | 説明                                 | 前提条件                                             |
| -------------- | ------------------------------------ | ---------------------------------------------------- |
| Module（推奨） | ImportExcel PowerShellモジュール使用 | 初回実行時に自動インストール（要インターネット接続） |
| COM            | Excel COMオートメーション使用        | Microsoft Excel がインストール済みであること         |

## 7. 使用方法

### 7.1 前提条件

| #   | 条件                          | 備考                                                                   |
| --- | ----------------------------- | ---------------------------------------------------------------------- |
| 1   | Windows OS                    | PowerShell 5.1 以上                                                    |
| 2   | PowerShell 実行ポリシー       | `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` |
| 3   | Oracle SQLソースファイル      | SVNからチェックアウト済みであること                                    |
| 4   | VB.NETソースファイル          | ローカルに配置済みであること                                           |
| 5   | Excel出力（Module方式の場合） | インターネット接続（ImportExcel 初回インストール時）                   |
| 6   | Excel出力（COM方式の場合）    | Microsoft Excel インストール済み                                       |

### 7.2 セットアップ手順

#### Step 1: 設定ファイルの編集

`scripts/config.psd1` を環境に合わせて編集する。

```powershell
@{
    Oracle = @{
        SourcePath = "C:\SVN\Oracle\packages"  # ← 実際のパスに変更
        FilePattern = "*.sql"
    }
    Ddl = @{
        TableSourcePath  = "C:\SVN\Oracle\tables"   # ← 実際のパスに変更
        IndexSourcePath  = "C:\SVN\Oracle\indexes"   # ← 実際のパスに変更
        FilePattern      = "*.sql"
        ExpandSelectStar = $true
    }
    VbNet = @{
        SourcePath = "C:\SVN\VbNet\src"        # ← 実際のパスに変更
        DacFilePattern = "*dac*.vb"
    }
    # ... 以下省略
}
```

#### Step 2: ImportExcel モジュールのインストール（Module方式の場合）

```powershell
Install-Module ImportExcel -Scope CurrentUser
```

### 7.3 実行方法

#### 基本実行（全解析 + Excel出力）

```powershell
cd C:\path\to\CRUD\scripts
.\Run-CrudAnalysis.ps1
```

#### COM方式でExcel出力

```powershell
.\Run-CrudAnalysis.ps1 -ExportMode COM
```

#### Oracle解析のみ実行

```powershell
.\Run-CrudAnalysis.ps1 -SkipVbNet
```

#### VB.NET解析のみ実行

```powershell
.\Run-CrudAnalysis.ps1 -SkipOracle
```

#### DDL解析をスキップして実行

```powershell
.\Run-CrudAnalysis.ps1 -SkipDdl
```

#### 設定ファイルを指定して実行

```powershell
.\Run-CrudAnalysis.ps1 -ConfigPath "C:\config\my_config.psd1"
```

### 7.4 実行結果の確認

実行後、`scripts/output/` に以下が出力される。

| ファイル            | 内容                             |
| ------------------- | -------------------------------- |
| `CrudMatrix.xlsx`   | CRUDマトリックス（6シート構成）  |
| `crud_results.json` | 中間データ（デバッグ・再加工用） |

### 7.5 実行ログ例

```
============================================
  CRUD解析ツール
============================================

[設定] 設定ファイル読み込み完了: .\config.psd1

--- Oracle SQL 解析 ---
[Oracle] 解析開始: C:\SVN\Oracle\packages
[Oracle] 対象ファイル数: 45
[Oracle] 解析完了: 1234 件のCRUDエントリを検出

--- VB.NET DAC 解析 ---
[VB.NET] 解析開始: C:\SVN\VbNet\src
[VB.NET] 対象DACファイル数: 32
[VB.NET] 解析完了: 876 件のCRUDエントリを検出

--- DDL 解析 ---
[DDL] テーブル定義解析開始: C:\SVN\Oracle\tables
[DDL] 対象ファイル数: 58
[DDL] 解析完了: テーブル 58 件, カラム 742 件, インデックス 95 件

--- SELECT * 展開 ---
[DDL] SELECT * 展開: 23 箇所を個別カラムに展開

--- 未使用カラム分析 ---
[分析] 未使用カラム: 47 件（12 テーブル）

--- 解析結果サマリー ---
  検出テーブル数（CRUD） : 58
  検出機能数             : 128
  総CRUDエントリ         : 2380
  テーブル定義数（DDL）  : 58 テーブル / 742 カラム
  インデックス定義数     : 95
  未使用カラム数         : 47

--- Excel出力 ---
[Excel] サマリーシート作成中...
[Excel] 詳細シート作成中...
[Excel] テーブル定義シート作成中...
[Excel] インデックス定義シート作成中...
[Excel] 未使用カラムシート作成中...
[Excel] 生データシート作成中...
[Excel] 出力完了: .\output\CrudMatrix.xlsx

============================================
  CRUD解析完了
============================================
```

## 8. 制約・注意事項

### 8.1 解析の制約

| #   | 制約                                         | 影響                                                           | 対処方法                     |
| --- | -------------------------------------------- | -------------------------------------------------------------- | ---------------------------- |
| 1   | 動的SQL（EXECUTE IMMEDIATE）は解析対象外     | 動的に組み立てるSQLのテーブル・カラムは検出不可                | 生データシートで手動補完     |
| 2   | サブクエリ内のテーブルは一部未検出の場合あり | ネストが深いSQLは漏れる可能性あり                              | JSONデータを目視確認         |
| 3   | エイリアスからの逆引きは非対応               | `SELECT a.col1 FROM table1 a` でカラムの所属テーブル特定が困難 | FROM句のテーブル全てに紐付け |
| 4   | `SELECT *` の展開はDDL依存                   | テーブル定義DDLがない場合は `*` のまま記録される               | DDLファイルをSVNから取得     |
| 5   | VB.NETの複雑な文字列組立は未検出の場合あり   | String.Format、補間文字列の一部パターン                        | 生データ確認後に手動補完     |
| 6   | WITH句（CTE）内のテーブルは一部未検出        | 再帰クエリ等の複雑なCTE                                        | 手動確認                     |

### 8.2 運用上の注意

- **初回実行時**: 出力結果を目視でサンプル確認し、解析精度を検証すること
- **定期実行**: ソース変更後にCRUD図を最新化する場合は再実行する
- **手動補完**: 生データシートを確認し、漏れがあれば手動で追加する
- **大量ファイル**: ファイル数が多い場合は `SkipOracle` / `SkipVbNet` で分割実行可能

## 9. 今後の拡張予定

| #   | 拡張内容                                   | 優先度 |
| --- | ------------------------------------------ | ------ |
| 1   | EXECUTE IMMEDIATE 内SQL解析                | 中     |
| 2   | Synonym 解決（実テーブル名への変換）       | 中     |
| 3   | Sequence 使用箇所の検出                    | 中     |
| 4   | View定義の再帰解析（参照先テーブルの特定） | 低     |
| 5   | 差分検出（前回実行結果との比較）           | 低     |
