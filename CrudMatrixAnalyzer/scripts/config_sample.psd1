@{
    # ===================================================
    # CRUD解析ツール 設定ファイル
    # ===================================================

    # --- Oracle SQL ソース設定 ---
    Oracle = @{
        # Oracle SQLファイルの格納ディレクトリ（SVNワーキングコピー等）
        # Package Body, Trigger, View, Function, Procedure を含む
        SourcePath       = "C:\App\CrudMatrixAnalyzer\test_data\package"

        # 対象ファイルパターン
        FilePattern      = "pkg_*.sql"

        # 除外パターン（バックアップ等）
        ExcludePatterns  = @("*_bak*", "*_old*", "*_backup*")

        # 対象オブジェクト種別
        ObjectTypes      = @("PACKAGE BODY", "TRIGGER", "VIEW", "MATERIALIZED VIEW", "FUNCTION", "PROCEDURE")

        # プロシージャ単位の抽出デバッグ（Run-CrudAnalysis.ps1 の -DebugOracle と同じ）
        DebugLog         = $false

        # WITH が別メソッド等で組み立てられ、同一フラグメントに WITH が無いときに FROM で使う CTE 名を明示（任意・大文字小文字無視）
        KnownCteNames    = @()

        # ファイルの文字コード: "auto"（自動判定）/ "utf-8" / "shift_jis"
        # 自動判定は BOM付きUTF-8 → UTF-8 → Shift-JIS の順で試みる
        # 文字化けする場合は "shift_jis" を明示指定してください（Oracle SQL は Shift-JIS が多い）
        SourceEncoding   = "shift_jis"
    }

    # --- Oracle DDL（テーブル定義・インデックス定義）設定 ---
    Ddl = @{
        # テーブル定義SQLファイルの格納ディレクトリ
        TableSourcePath  = "c:\App\CrudMatrixAnalyzer\test_data\tables"

        # インデックス定義SQLファイルの格納ディレクトリ
        IndexSourcePath  = "c:\App\CrudMatrixAnalyzer\test_data\tables"

        # 対象ファイルパターン
        FilePattern      = "create_*.sql"

        # 除外パターン
        ExcludePatterns  = @("*_bak*", "*_old*", "*_backup*")

        # SELECT * 展開を行うか
        ExpandSelectStar = $true

        # ファイルの文字コード: "auto" / "utf-8" / "shift_jis"（DDL も Shift-JIS が多い）
        SourceEncoding   = "shift_jis"
    }

    # --- VB.NET ソース設定 ---
    VbNet = @{
        # VB.NETソースのルートディレクトリ
        SourcePath       = "C:\App\CrudMatrixAnalyzer\test_data\vbnet"

        # DACファイルパターン（ファイル名にdacを含むもの）
        DacFilePattern   = "*dac*.vb"

        # 除外パターン
        ExcludePatterns  = @("*.Designer.vb", "*AssemblyInfo*", "*.g.vb")

        # ファイルの文字コード: "auto" / "utf-8" / "shift_jis"（VB.NET は UTF-8 が多い）
        SourceEncoding   = "utf-8"
    }

    # --- 出力設定 ---
    Output = @{
        # 中間JSON出力パス
        JsonPath         = "..\..\test_data\output\
        crud_results.json"

        # Excel出力パス
        ExcelPath        = "..\..\test_data\output\CrudMatrix.xlsx"

        # サマリーシート名
        SummarySheetName = "テーブル×機能サマリー"

        # 詳細シート名
        DetailSheetName  = "項目別詳細"
    }

    # --- 除外テーブル ---
    ExcludeTables = @("DUAL", "PLAN_TABLE", "ALL_OBJECTS", "USER_OBJECTS", "ALL_TAB_COLUMNS", "TABLE")

    # --- 除外スキーマプレフィックス ---
    ExcludeSchemas = @("SYS", "SYSTEM", "DBMS_")
}
