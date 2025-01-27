VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommonEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredefinedId = False
Attribute VB_Exposed = True
Option Explicit

'@Description("システム全体で使用される共通の列挙型を定義するクラス")

'@Description("ログレベルを定義する列挙型")
'@Remarks("ログの重要度を表し、フィルタリングやハンドリングの基準として使用")
Public Enum LogLevel
    llTrace = 0    ' 最も詳細なデバッグ情報（開発時のみ使用）
    llDebug = 1    ' デバッグ情報（開発・テスト環境で使用）
    llInfo = 2     ' 一般的な情報（通常の操作ログ）
    llWarning = 3  ' 警告（潜在的な問題）
    llError = 4    ' エラー（回復可能な問題）
    llFatal = 5    ' 致命的なエラー（システム停止レベルの問題）
End Enum

'@Description("エラー通知レベルを定義する列挙型")
'@Remarks("エラーの重大度に応じた通知方法の選択に使用")
Public Enum ErrorNotificationLevel
    enlNone = 0      ' 通知なし（ログも記録しない）
    enlSilent = 1    ' サイレント（ログのみ記録）
    enlInfo = 2      ' 情報（ユーザーに情報として通知）
    enlWarning = 3   ' 警告（ユーザーに警告として通知）
    enlError = 4     ' エラー（ユーザーにエラーとして通知）
    enlCritical = 5  ' 重大（即時通知が必要な重大な問題）
End Enum

'@Description("バリデーション優先度を定義する列挙型")
'@Remarks("バリデーションルールの実行順序の制御に使用")
Public Enum ValidationPriority
    vpLowest = 0    ' 最低優先度（他のすべての検証の後に実行）
    vpLow = 1       ' 低優先度（基本的な検証の後に実行）
    vpNormal = 2    ' 通常優先度（標準的な実行順序）
    vpHigh = 3      ' 高優先度（基本的な検証の前に実行）
    vpHighest = 4   ' 最高優先度（最初に実行される検証）
End Enum

'@Description("バリデーション状態を定義する列挙型")
'@Remarks("バリデーションの現在の実行状態を表す")
Public Enum ValidationState
    vsNotStarted = 0   ' 未開始
    vsPending = 1      ' 保留中
    vsRunning = 2      ' 実行中
    vsPaused = 3       ' 一時停止
    vsCompleted = 4    ' 完了
    vsFailed = 5       ' 失敗
    vsCancelled = 6    ' キャンセル
End Enum

'@Description("エラー回復状態を定義する列挙型")
'@Remarks("エラーからの回復処理の状態を表す")
Public Enum RecoveryState
    rsNotStarted = 0    ' 未開始
    rsAttempting = 1    ' 回復試行中
    rsSucceeded = 2     ' 回復成功
    rsFailed = 3        ' 回復失敗
    rsPartial = 4       ' 部分的に回復
    rsAborted = 5       ' 中断
End Enum

'@Description("通知配信状態を定義する列挙型")
'@Remarks("通知の配信状態を表す")
Public Enum DeliveryState
    dsQueued = 0        ' キュー待ち
    dsProcessing = 1    ' 処理中
    dsDelivered = 2     ' 配信済み
    dsFailed = 3        ' 配信失敗
    dsRetrying = 4      ' 再試行中
    dsThrottled = 5     ' スロットル中
    dsExpired = 6       ' 期限切れ
End Enum

'@Description("リソース状態を定義する列挙型")
'@Remarks("システムリソースの状態を表す")
Public Enum ResourceState
    rsAvailable = 0     ' 利用可能
    rsLow = 1          ' 残り少ない
    rsCritical = 2     ' 危機的状態
    rsExhausted = 3    ' 枯渇
    rsUnknown = 4      ' 不明
End Enum

'@Description("キャッシュ状態を定義する列挙型")
'@Remarks("キャッシュエントリの状態を表す")
Public Enum CacheState
    csValid = 0        ' 有効
    csExpired = 1      ' 期限切れ
    csInvalid = 2      ' 無効
    csUpdating = 3     ' 更新中
End Enum

'@Description("バリデーション依存関係の種類を定義する列挙型")
'@Remarks("バリデーションルール間の依存関係を表す")
Public Enum ValidationDependencyType
    vdtRequired = 0    ' 必須（依存先が成功する必要がある）
    vdtOptional = 1    ' オプション（依存先の結果に影響されない）
    vdtExclusive = 2   ' 排他的（依存先が失敗する必要がある）
    vdtConditional = 3 ' 条件付き（条件に応じて依存関係が変化）
End Enum

'@Description("パフォーマンスメトリクスの種類を定義する列挙型")
'@Remarks("システムの各種パフォーマンス指標を表す")
Public Enum MetricType
    mtResponseTime = 0   ' 応答時間
    mtThroughput = 1    ' スループット
    mtErrorRate = 2     ' エラー率
    mtResourceUsage = 3  ' リソース使用率
    mtQueueLength = 4   ' キュー長
    mtLatency = 5       ' レイテンシ
End Enum

'@Description("バリデーション検証レベルを定義する列挙型")
'@Remarks("バリデーションの詳細度を表す")
Public Enum ValidationLevel
    vlBasic = 0        ' 基本的な検証
    vlNormal = 1       ' 通常の検証
    vlStrict = 2       ' 厳密な検証
    vlComplete = 3     ' 完全な検証
End Enum