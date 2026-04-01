@{
    # ===================================================
    # CRUD解析ツール 設定ファイル
    # ===================================================

    # --- Oracle SQL ソース設定 ---
    Oracle = @{
        # Oracle SQLファイルの格納ディレクトリ（SVNワーキングコピー等）
        # Package Body, Trigger, View, Function, Procedure を含む
        SourcePath       = "C:\SVN\Oracle\packages"

        # 対象ファイルパターン
        FilePattern      = "*.sql"

        # 除外パターン（バックアップ等）
        ExcludePatterns  = @("*_bak*", "*_old*", "*_backup*")

        # 対象オブジェクト種別
        ObjectTypes      = @("PACKAGE BODY", "TRIGGER", "VIEW", "MATERIALIZED VIEW", "FUNCTION", "PROCEDURE")
    }

    # --- Oracle DDL（テーブル定義・インデックス定義）設定 ---
    Ddl = @{
        # テーブル定義SQLファイルの格納ディレクトリ
        TableSourcePath  = "C:\SVN\Oracle\tables"

        # インデックス定義SQLファイルの格納ディレクトリ
        IndexSourcePath  = "C:\SVN\Oracle\indexes"

        # 対象ファイルパターン
        FilePattern      = "*.sql"

        # 除外パターン
        ExcludePatterns  = @("*_bak*", "*_old*", "*_backup*")

        # SELECT * 展開を行うか
        ExpandSelectStar = $true
    }

    # --- VB.NET ソース設定 ---
    VbNet = @{
        # VB.NETソースのルートディレクトリ
        SourcePath       = "C:\SVN\VbNet\src"

        # DACファイルパターン（ファイル名にdacを含むもの）
        DacFilePattern   = "*dac*.vb"

        # 除外パターン
        ExcludePatterns  = @("*.Designer.vb", "*AssemblyInfo*", "*.g.vb")
    }

    # --- 出力設定 ---
    Output = @{
        # 中間JSON出力パス
        JsonPath         = ".\output\crud_results.json"

        # Excel出力パス
        ExcelPath        = ".\output\CrudMatrix.xlsx"

        # サマリーシート名
        SummarySheetName = "テーブル×機能サマリー"

        # 詳細シート名
        DetailSheetName  = "項目別詳細"
    }

    # --- 除外テーブル ---
    ExcludeTables = @("DUAL", "PLAN_TABLE", "ALL_OBJECTS", "USER_OBJECTS", "ALL_TAB_COLUMNS")

    # --- 除外スキーマプレフィックス ---
    ExcludeSchemas = @("SYS", "SYSTEM", "DBMS_")
}
