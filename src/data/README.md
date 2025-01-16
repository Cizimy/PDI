# データ層

このディレクトリには、データアクセスと検証を担当するモジュールが含まれています。

## モジュール構成

### modDBHandler.bas
データベース操作の中央管理システム。ADO（ActiveX Data Objects）を使用したデータベースアクセスを提供します。

```vb
' データベース接続と操作の例
Public Sub ProcessDatabaseOperation()
    ' 初期化と接続
    modDBHandler.Initialize
    If modDBHandler.Connect Then
        ' トランザクション開始
        modDBHandler.BeginTransaction
        
        On Error GoTo ErrorHandler
        
        ' データの取得
        Dim rs As Object
        Set rs = modDBHandler.ExecuteQuery("SELECT * FROM Users WHERE Active = ?", Array(True))
        
        ' データの更新
        Dim affected As Long
        affected = modDBHandler.ExecuteCommand("UPDATE Users SET LastLogin = ? WHERE UserID = ?", _
                                             Array(Now, 123))
        
        ' トランザクションのコミット
        modDBHandler.CommitTransaction
        Exit Sub
        
    ErrorHandler:
        ' エラー発生時はロールバック
        modDBHandler.RollbackTransaction
        modErrorHandler.HandleError Err.Number, Err.Description, "ProcessDatabaseOperation", etDatabase
    End If
End Sub
```

### modFileSystem.bas
ファイルシステム操作の中央管理システム。ファイルの読み書き、コピー、移動などの操作を提供します。

```vb
' ファイル操作の例
Public Sub ProcessFileOperation()
    ' ファイルの読み書き
    If modFileSystem.FileExists("data.txt") Then
        Dim content As String
        content = modFileSystem.ReadFile("data.txt")
        modFileSystem.WriteFile "backup.txt", content
    End If
    
    ' ディレクトリ操作
    If Not modFileSystem.DirectoryExists("logs") Then
        modFileSystem.CreateDirectory "logs"
    End If
    
    ' ファイル一覧の取得
    Dim files As Variant
    files = modFileSystem.GetFiles("logs", "*.log")
    
    ' ファイル情報の取得
    Dim fileInfo As Object
    Set fileInfo = modFileSystem.GetFileInfo("data.txt")
    Debug.Print "File Size: " & fileInfo("Size")
End Sub
```

### modDataValidator.bas
データ入力の検証機能を提供します。様々な形式のデータ検証に対応しています。

```vb
' データ検証の例
Public Function ValidateUserInput(ByVal userName As String, _
                                ByVal email As String, _
                                ByVal age As Variant) As Boolean
                                
    Dim result As ValidationResult
    
    ' 必須入力の検証
    result = modDataValidator.ValidateRequired(userName, "ユーザー名")
    If Not result.IsValid Then
        MsgBox result.ErrorMessage
        Exit Function
    End If
    
    ' メールアドレス形式の検証
    result = modDataValidator.ValidateEmail(email, "メールアドレス")
    If Not result.IsValid Then
        MsgBox result.ErrorMessage
        Exit Function
    End If
    
    ' 数値範囲の検証
    result = modDataValidator.ValidateNumberRange(age, "年齢", 0, 120)
    If Not result.IsValid Then
        MsgBox result.ErrorMessage
        Exit Function
    End If
    
    ValidateUserInput = True
End Function
```

## モジュール間の連携

1. エラー処理とログ記録
```vb
Private Sub HandleDataOperation()
    On Error GoTo ErrorHandler
    
    ' データベース操作
    modDBHandler.BeginTransaction
    
    ' ファイル操作
    Dim data As String
    data = modFileSystem.ReadFile("input.txt")
    
    ' データ検証
    Dim result As ValidationResult
    result = modDataValidator.ValidateRequired(data, "入力データ")
    
    If result.IsValid Then
        ' 検証OKならデータベースに保存
        modDBHandler.ExecuteCommand "INSERT INTO Data (Content) VALUES (?)", Array(data)
        modDBHandler.CommitTransaction
    Else
        modDBHandler.RollbackTransaction
    End If
    
    Exit Sub
    
ErrorHandler:
    modDBHandler.RollbackTransaction
    modErrorHandler.HandleError Err.Number, _
                               Err.Description, _
                               "HandleDataOperation", _
                               etDatabase
End Sub
```

2. 設定に基づくデータベース接続
```vb
Public Sub InitializeDataLayer()
    ' 設定からデータベース接続情報を取得
    Dim server As String
    Dim database As String
    server = modConfigManager.GetValue("database", "server")
    database = modConfigManager.GetValue("database", "database")
    
    ' データベース接続の初期化
    modDBHandler.Initialize
    
    ' ログレベルの設定
    Dim logLevel As String
    logLevel = modConfigManager.GetValue("logging", "level")
    modLogger.Configure "", logLevel
End Sub
```

## エラーコード体系

データ層で使用するエラーコードの範囲：

- vbObjectError + 1000: ファイルが存在しない
- vbObjectError + 1001: コピー元ファイルが存在しない
- vbObjectError + 1002: 移動元ファイルが存在しない
- vbObjectError + 1003: ディレクトリが存在しない
- vbObjectError + 1004: ファイル情報の取得に失敗

## セキュリティ考慮事項

1. データベースセキュリティ
   - パラメータ化クエリの使用（SQLインジェクション対策）
   - 接続情報の暗号化
   - 最小権限の原則

2. ファイルシステムセキュリティ
   - アクセス権限の確認
   - パスのバリデーション
   - セキュアな一時ファイル処理

3. 入力データのセキュリティ
   - 不正な入力値の検証
   - エスケープ処理
   - 文字エンコーディングの制御

## パフォーマンス最適化

1. データベース操作
   - コネクションプーリング
   - バッチ処理の活用
   - インデックスの適切な使用

2. ファイル操作
   - バッファリングの活用
   - 非同期処理の検討
   - ファイルハンドルの適切な管理

3. データ検証
   - キャッシュの活用
   - 正規表現の最適化
   - バリデーションの順序の最適化

## 今後の拡張計画

1. 新機能の追加
   - NoSQLデータベースサポート
   - クラウドストレージ連携
   - データ暗号化機能

2. パフォーマンス改善
   - 非同期処理の実装
   - キャッシュ機構の強化
   - バッチ処理の最適化

3. セキュリティ強化
   - 監査ログの実装
   - アクセス制御の強化
   - 暗号化機能の拡充
