@{
    # VB.NET プロジェクト関連ファイルとして計上する拡張子（小文字・ドット付き）
    # product-requirements.md §3「バイナリファイルを除く テキスト系の全拡張子 を対象とする」に基づく。
    VbNetExtensions = @(
        '.vb', '.vbproj', '.sln', '.resx', '.config', '.xml',
        '.xsd', '.xsl', '.xslt', '.txt', '.ini', '.bat', '.cmd',
        '.ps1', '.psm1', '.psd1', '.md', '.csv'
    )

    # PL/SQL ソースファイルとして計上する拡張子（小文字・ドット付き）
    PlSqlExtensions = @(
        '.sql', '.pkb', '.pks', '.prc', '.fnc', '.trg', '.vw'
    )

    # バイナリ・スキャン対象外の拡張子（Language=SKIPPED として配列に含める）
    ExcludeExtensions = @(
        '.exe', '.dll', '.pdb', '.lib', '.obj', '.bin', '.dat',
        '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.ico', '.tif', '.tiff',
        '.zip', '.7z', '.rar', '.cab', '.msi', '.gz', '.tar',
        '.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx'
    )

    # 除外するフォルダ名（パスセグメント単位で完全一致。前フィルタで降りない）
    ExcludeFolders = @(
        'bin', 'obj',
        '.git', '.vs',
        'node_modules',
        'output', 'logs',
        '.steering'
    )
}
