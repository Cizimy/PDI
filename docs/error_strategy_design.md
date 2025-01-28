# DefaultErrorStrategy クラス設計ドキュメント

## 概要
`IErrorStrategy`インターフェースを実装する`clsDefaultErrorStrategy`クラスは、エラーからの回復戦略を提供します。このクラスは、段階的な回復処理とフォールバックメカニズムを実装し、エラー処理の信頼性と柔軟性を確保します。

## クラス定義

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDefaultErrorStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
```

## プライベート変数

```vb
Private mMaxRetryCount As Long           ' 最大リトライ回数
Private mRetryInterval As Long           ' リトライ間隔（ミリ秒）
Private mPriority As Long                ' 戦略の優先度
Private mFallbackStrategy As IErrorStrategy ' フォールバック戦略
Private mMetrics As Dictionary           ' パフォーマンスメトリクス
Private mProgress As RecoveryProgress    ' 回復の進捗状況
Private mLogger As Object                ' ロギング用オブジェクト
Private mIsInitialized As Boolean        ' 初期化フラグ
```

## 定数定義

```vb
Private Const DEFAULT_MAX_RETRY_COUNT As Long = 3
Private Const DEFAULT_RETRY_INTERVAL_MS As Long = 1000
Private Const MAX_RECOVERY_ATTEMPTS As Long = 5
Private Const RECOVERY_TIMEOUT_MS As Long = 30000
```

## 主要メソッド

### RecoverFromError
エラーからの回復を試みるメイン機能を提供します。

```vb
Public Function RecoverFromError(ByRef errorInfo As ErrorInfo, _
                               ByVal errorManager As clsErrorManager, _
                               Optional ByVal context As RecoveryContext, _
                               Optional ByVal options As RecoveryOptions) As Boolean
```

#### 実装の要点
1. エラー情報の検証
2. 回復コンテキストの初期化
3. タイムアウト処理の実装
4. 段階的な回復処理の実行
5. メトリクスの収集
6. イベントの発行
7. リソースの適切な解放

### CreateRecoveryChain
回復戦略のチェーンを作成します。

```vb
Public Function CreateRecoveryChain(ByVal strategies As Collection, _
                                  ByVal errorManager As clsErrorManager) As RecoveryChain
```

#### 実装の要点
1. 戦略の優先度に基づくソート
2. チェーンの検証
3. エラーマネージャーとの連携
4. メタデータの設定

## プロパティ

### MaxRetryCount
```vb
Public Property Let MaxRetryCount(ByVal value As Long)
Public Property Get MaxRetryCount() As Long
```

### RetryInterval
```vb
Public Property Let RetryInterval(ByVal value As Long)
Public Property Get RetryInterval() As Long
```

### Priority
```vb
Public Property Let Priority(ByVal value As Long)
Public Property Get Priority() As Long
```

### FallbackStrategy
```vb
Public Property Set FallbackStrategy(ByVal value As IErrorStrategy)
Public Property Get FallbackStrategy() As IErrorStrategy
```

## エラー処理

### エラーログ記録
```vb
Private Sub LogError(ByVal source As String, _
                    ByVal message As String, _
                    Optional ByVal details As String = "")
```

### メトリクス収集
```vb
Private Sub UpdateMetrics(ByVal metricName As String, _
                         ByVal value As Variant)
```

## イベント処理

```vb
Public Event RecoveryChainStarted(ByVal chainId As String, _
                                 ByVal strategies As Collection)
Public Event RecoveryChainCompleted(ByVal chainId As String, _
                                   ByVal successCount As Long, _
                                   ByVal failureCount As Long)
Public Event FallbackStrategyActivated(ByVal errorInfo As ErrorInfo, _
                                      ByVal fallbackStrategy As IErrorStrategy, _
                                      ByVal reason As FallbackReason)
```

## 初期化とクリーンアップ

### Initialize
```vb
Private Sub Class_Initialize()
    ' デフォルト値の設定
    mMaxRetryCount = DEFAULT_MAX_RETRY_COUNT
    mRetryInterval = DEFAULT_RETRY_INTERVAL_MS
    Set mMetrics = New Dictionary
    Set mProgress = New RecoveryProgress
    Set mLogger = CreateObject("Scripting.FileSystemObject")
    mIsInitialized = True
End Sub
```

### Cleanup
```vb
Private Sub Class_Terminate()
    Set mFallbackStrategy = Nothing
    Set mMetrics = Nothing
    Set mProgress = Nothing
    Set mLogger = Nothing
    mIsInitialized = False
End Sub
```

## 実装の注意点

1. スレッドセーフティ
   - 共有リソースへのアクセスを適切に同期
   - 競合状態を防ぐための排他制御

2. リソース管理
   - メモリリークの防止
   - ファイルハンドルの適切なクローズ
   - データベース接続の管理

3. エラー処理
   - 階層化されたエラーハンドリング
   - 適切なエラー情報の伝播
   - ログ記録の確実な実行

4. パフォーマンス
   - 効率的なリソース使用
   - 適切なキャッシュ戦略
   - 最適化されたアルゴリズム

5. メンテナンス性
   - 明確なコメント
   - モジュール化された設計
   - テスト可能なコード構造

## テスト戦略

1. ユニットテスト
   - 各メソッドの独立したテスト
   - エッジケースの検証
   - エラー条件のテスト

2. 統合テスト
   - ErrorManagerとの連携テスト
   - イベント処理の検証
   - リソース管理の確認

3. パフォーマンステスト
   - 負荷テスト
   - メモリ使用量の監視
   - タイミング検証

## 依存関係

- IErrorStrategy インターフェース
- ErrorManager クラス
- RecoveryChain クラス
- RecoveryProgress クラス
- ErrorInfo クラス
- その他の補助クラス

## セキュリティ考慮事項

1. 入力検証
   - パラメータの妥当性確認
   - 不正な値の検出と処理

2. リソースアクセス
   - 適切な権限管理
   - セキュアなリソースハンドリング

3. エラー情報
   - センシティブ情報の保護
   - 適切なエラーメッセージの構築

## 拡張性

1. カスタム戦略
   - 新しい回復戦略の追加が容易
   - 既存戦略のカスタマイズ可能

2. メトリクス拡張
   - 新しい測定項目の追加
   - カスタムレポート機能

3. イベント処理
   - 新しいイベントタイプの追加
   - イベントハンドラーのカスタマイズ

## 今後の改善点

1. 非同期処理のサポート
2. より詳細なメトリクス収集
3. 高度な回復戦略の実装
4. パフォーマンス最適化
5. セキュリティ強化