r# SQL抽出ツール仕様・実装計画

## Overview

**プログラムソースコードからSQL文を自動抽出し、単独実行可能なSQLファイルとして出力するPowerShellツール**の仕様と実装計画です。

対象言語は **PL/SQL** と **VB.NET** の2つ。動的SQLの文字列連結を解析してSQL部分のみを抽出し、IF分岐がある場合はすべての分岐パスを展開してSQL文を網羅的に収集します。

> **設計方針:** 各言語ごとにパーサーモジュールを分離し、共通のSQL整形・出力エンジンを通して結果を生成する構成とします。正規表現ベースの軽量パーサーで実用的な精度を実現します。

## Preferences

- **実装言語:** PowerShell
- **対象言語:** PL/SQL、VB.NET
- **動的SQL対応:** 文字列連結・編集箇所を除外し、SQL部分のみ抽出
- **IF分岐対応:** プログラム制御構文を除外し、全分岐パスを展開
- **出力形式:** 単独実行可能なSQLファイル（.sql）

## Implementation Plan

### Step 1: ツール全体構成

#### アーキテクチャ

    入力ソースファイル (.sql / .vb)
        │
        ├─ 言語判定モジュール（拡張子 or パラメータで判定）
        │
        ├─ PL/SQL パーサー
        │     ├─ 静的SQL抽出
        │     ├─ 動的SQL（EXECUTE IMMEDIATE等）解析
        │     └─ IF分岐展開
        │
        ├─ VB.NET パーサー
        │     ├─ 静的SQL抽出（CommandText代入等）
        │     ├─ 動的SQL（文字列連結）解析
        │     └─ If分岐展開
        │
        └─ SQL整形・出力エンジン
              ├─ SQL文の正規化・フォーマット
              ├─ バインド変数のプレースホルダ化
              └─ .sqlファイル出力

#### ファイル構成

| ファイル | 役割 |
|---|---|
| `script/Extract-Sql.ps1` | メインスクリプト（エントリポイント） |
| `script/PlSqlParser.psm1` | PL/SQLパーサーモジュール |
| `script/VbNetParser.psm1` | VB.NETパーサーモジュール |
| `script/SqlFormatter.psm1` | SQL整形・出力モジュール |

### Step 2: 入力仕様

#### コマンドライン仕様

    .\Extract-Sql.ps1
      -InputPath <string>       # 入力ファイル or フォルダパス（必須）
      -OutputDir <string>       # 出力先ディレクトリ（省略時: ./output）
      -Language <string>        # 言語指定: "plsql" | "vbnet" | "auto"（省略時: auto）
      -Encoding <string>        # 文字コード（省略時: UTF-8）

#### 言語自動判定ルール

| 拡張子 | 判定言語 |
|---|---|
| `.sql`, `.pls`, `.pck`, `.pkb`, `.pks` | PL/SQL |
| `.vb`, `.vbnet` | VB.NET |

> 📁 `-InputPath` にフォルダを指定した場合は、配下のファイルを再帰的に走査し、対象拡張子のファイルをすべて処理します。

### Step 3: PL/SQLパーサー仕様

#### 抽出対象

1. **静的SQL文**
   - `SELECT`, `INSERT`, `UPDATE`, `DELETE`, `MERGE` で始まるSQL文
   - `CURSOR` 宣言内のSELECT文
   - `CREATE`, `ALTER`, `DROP` 等のDDL文
2. **動的SQL（EXECUTE IMMEDIATE）**
   - `EXECUTE IMMEDIATE` に続く文字列リテラル・変数を解析
   - 文字列連結（`||`）で組み立てられたSQLを結合して抽出
3. **DBMS_SQL / OPEN FOR 文**
   - `DBMS_SQL.PARSE()` の引数からSQL抽出
   - `OPEN cursor FOR` に続くSQL文

#### 動的SQL解析ルール

    -- パターン例
    v_sql := 'SELECT col1, col2 FROM ' || v_table || ' WHERE id = ' || v_id;
    EXECUTE IMMEDIATE v_sql;

**処理:**

- 文字列リテラル（`'...'`）部分 → SQLとして抽出
- 変数部分（`v_table`, `v_id`）→ プレースホルダ `/*:変数名*/` に置換
- 連結演算子（`||`）→ 除去して文字列部分を結合

**出力例:**

    SELECT col1, col2 FROM /*:v_table*/ WHERE id = /*:v_id*/

#### IF分岐の展開ルール

    IF condition1 THEN
        v_sql := v_sql || ' AND status = ''A''';
    ELSIF condition2 THEN
        v_sql := v_sql || ' AND status = ''B''';
    ELSE
        v_sql := v_sql || ' AND status = ''C''';
    END IF;

**処理方針:**

- `IF / ELSIF / ELSE / END IF` の制御構文を除去
- **すべての分岐内のSQL断片を展開**して出力
- 分岐ごとにコメントで区切りを入れる

**出力例:**

    -- [Branch 1] condition1
     AND status = 'A'
    -- [Branch 2] condition2
     AND status = 'B'
    -- [Branch 3] ELSE
     AND status = 'C'

### Step 4: VB.NETパーサー仕様

#### 抽出対象

1. **CommandText代入**
   - `cmd.CommandText = "SELECT ..."` パターン
   - `New SqlCommand("SELECT ...", conn)` パターン
2. **文字列変数へのSQL組み立て**
   - `Dim sql As String = "SELECT ..."` パターン
   - `sql &= "..."` / `sql = sql & "..."` / `sql += "..."` パターン
   - `StringBuilder.Append("...")` / `.AppendLine("...")` パターン

#### 動的SQL解析ルール

    Dim sql As String = "SELECT col1, col2 FROM " & tableName & " WHERE id = " & id

**処理:**

- 文字列リテラル（`"..."`）部分 → SQLとして抽出
- 変数部分（`tableName`, `id`）→ プレースホルダ `/*:変数名*/` に置換
- 連結演算子（`&`, `+`）→ 除去して文字列部分を結合
- `String.Format` → `{0}`, `{1}` をプレースホルダに変換
- 補間文字列 `$"...{var}..."` → 変数部分をプレースホルダに変換

#### IF分岐の展開ルール

    If condition Then
        sb.Append(" AND status = 'A'")
    Else
        sb.Append(" AND status = 'B'")
    End If

**処理方針（PL/SQLと同一）:**

- `If / ElseIf / Else / End If` の制御構文を除去
- すべての分岐内のSQL断片を展開して出力
- 分岐ごとにコメントで区切り

> ⚠️ **VB.NET固有の注意点:**
> - 行継続文字 `_` への対応（複数行にまたがる文字列連結）
> - `vbCrLf` / `Environment.NewLine` → 改行として処理
> - `""` （ダブルクォートのエスケープ）→ 単一の `"` に変換

### Step 5: SQL整形・出力仕様

#### 出力ファイル命名規則

    {元ファイル名}_{連番:3桁}.sql

例: `OrderProc.pkb` から3つのSQLを抽出した場合

- `OrderProc_001.sql`
- `OrderProc_002.sql`
- `OrderProc_003.sql`

#### 出力SQLのフォーマット

各SQLファイルには以下のヘッダコメントを付与:

    -- ============================================
    -- Source: OrderProc.pkb
    -- Line: 45-62
    -- Type: SELECT (Dynamic SQL)
    -- Extracted: 2026-04-14 21:30:00
    -- ============================================
    
    SELECT col1, col2
    FROM /*:v_table*/
    WHERE id = /*:v_id*/
      -- [Branch 1] condition1
      AND status = 'A'
      -- [Branch 2] ELSE
      AND status = 'B'
    ;

#### 整形ルール

- [ ] SQLキーワードを大文字に統一（`SELECT`, `FROM`, `WHERE` 等）
- [ ] インデントを統一（スペース2つ）
- [ ] 末尾にセミコロン `;` を付与
- [ ] 連続する空白・改行を正規化
- [ ] 抽出元の行番号をコメントに記録

### Step 6: プレースホルダ変換ルール

動的SQLの変数部分はSQLとして実行可能な形を維持するため、以下のルールで変換します。

#### 変換パターン一覧

| 元のコード | 変換後 | 説明 |
|---|---|---|
| `v_table`（テーブル名位置） | `/*:v_table*/DUAL` | テーブル位置はダミーテーブル付与 |
| `v_id`（WHERE条件値） | `/*:v_id*/` | 値位置はコメントのみ |
| `TO_DATE(v_date, 'YYYYMMDD')` | そのまま維持 | SQL関数はそのまま |
| `{0}`, `{1}`（String.Format） | `/*:param0*/`, `/*:param1*/` | パラメータ名で置換 |
| `$"{varName}"`（補間文字列） | `/*:varName*/` | 変数名で置換 |

> 💡 **設計意図:** `/*:変数名*/` 形式のコメントを使うことで、SQLとして構文的に有効な状態を保ちつつ、元の変数名を参照できるようにします。

### Step 7: エラー処理・ログ出力

#### ログ出力

処理結果をコンソールとログファイルに出力:

    [INFO]  Processing: OrderProc.pkb (PL/SQL)
    [INFO]  Found 3 SQL statements (2 static, 1 dynamic)
    [INFO]  Output: output/OrderProc_001.sql
    [INFO]  Output: output/OrderProc_002.sql
    [INFO]  Output: output/OrderProc_003.sql
    [WARN]  Line 120: Unresolved variable in SQL - skipped
    [INFO]  Summary: 3 files processed, 8 SQLs extracted, 1 warning

#### エラーハンドリング

| 状況 | 動作 |
|---|---|
| 入力ファイルが存在しない | エラーメッセージを出力して終了 |
| 言語判定不可 | 警告を出力してスキップ |
| SQL抽出ゼロ件 | 情報メッセージ出力（正常終了） |
| 構文解析が不完全 | 警告付きで解析可能な部分だけ出力 |
| 出力先書き込み不可 | エラーメッセージを出力して終了 |

### Step 8: PowerShell実装方針

#### 技術的な実装アプローチ

- **正規表現ベース**のパーサーで実装（AST構築は行わない軽量設計）
- **状態マシン**でネスト構造（IF内IF等）を管理
- PowerShell **5.1以上** 対応（.NET Framework依存なし）

#### 主要クラス/関数設計

    # メインスクリプト
    Extract-Sql.ps1
      ├─ Get-SourceLanguage()       # 言語判定
      ├─ Invoke-PlSqlParser()       # PL/SQLパース実行
      ├─ Invoke-VbNetParser()       # VB.NETパース実行
      └─ Export-SqlFiles()          # SQL出力
    
    # 共通ユーティリティ
      ├─ Merge-DynamicSql()         # 動的SQL文字列結合
      ├─ Expand-IfBranches()        # IF分岐展開
      ├─ Convert-ToPlaceholder()    # 変数→プレースホルダ変換
      └─ Format-SqlStatement()      # SQL整形

#### テスト方針

- [ ] PL/SQL: 静的SQL / 動的SQL / EXECUTE IMMEDIATE / IF分岐 / ネストIF
- [ ] VB.NET: CommandText / StringBuilder / 文字列連結 / If分岐 / 行継続
- [ ] エッジケース: 空ファイル / SQLなし / コメント内SQL（無視すべき）/ 複数SQLの混在

---

> 🚀 **今後の拡張候補（対象外）:**
> - C# / Java 対応
> - バインド変数の型推定
> - SQL実行計画の取得
> - GUI（WPF）対応

## Architecture

    flowchart TD
        A["入力ソースファイル .sql / .vb"] --> B{"言語判定"}
        B -->|PL/SQL| C["PL/SQL パーサー"]
        B -->|VB.NET| D["VB.NET パーサー"]
        C --> C1["静的SQL抽出"]
        C --> C2["動的SQL解析"]
        C --> C3["IF分岐展開"]
        D --> D1["静的SQL抽出"]
        D --> D2["動的SQL解析"]
        D --> D3["If分岐展開"]
        C1 & C2 & C3 --> E["SQL整形エンジン"]
        D1 & D2 & D3 --> E
        E --> F["プレースホルダ変換"]
        F --> G["SQLフォーマット"]
        G --> H["出力: .sql ファイル群"]